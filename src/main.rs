use chrono::{Local, NaiveDateTime, TimeZone};
use filetime::{self, FileTime};
use lopdf::Document;
use std::fs;
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use zip::{ZipArchive, ZipWriter, write::FileOptions};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("========================================");
    println!("   PDF & Word 综合处理工具 (v0.4)");
    println!("========================================");

    // 1. 获取输入路径
    print!("请输入文件或文件夹路径: ");
    io::stdout().flush()?;
    let mut input_str = String::new();
    io::stdin().read_line(&mut input_str)?;
    let input_path = PathBuf::from(input_str.trim().replace("\"", ""));

    if !input_path.exists() {
        println!("❌ 错误：路径不存在！");
        wait_for_keypress();
        return Ok(());
    }

    // 2. 选择功能模式
    println!("\n请选择操作模式:");
    println!(" [1] 仅清理元数据 (删除作者、公司、创建时间等内部信息)");
    println!(" [2] 仅修改系统时间 (创建/修改/访问时间)");
    println!(" [3] 全套处理 (清理元数据 + 修改系统时间)");
    print!("请输入选项 (1/2/3): ");
    io::stdout().flush()?;
    let mut mode_str = String::new();
    io::stdin().read_line(&mut mode_str)?;
    let mode = mode_str.trim();

    // 3. 如果需要修改时间，则提前询问
    let mut times = None;
    if mode == "2" || mode == "3" {
        println!("\n--- 时间设置 (格式: YYYY-MM-DD HH:MM:SS，直接回车使用当前时间) ---");
        let t_create = ask_time("创建时间 (Created)");
        let t_modify = ask_time("修改时间 (Modified)");
        let t_access = ask_time("访问时间 (Accessed)");
        times = Some((t_create, t_modify, t_access));
    }

    // 4. 收集文件
    let mut files = Vec::new();
    if input_path.is_file() {
        files.push(input_path);
    } else {
        for entry in fs::read_dir(&input_path)? {
            let path = entry?.path();
            if path.is_file() {
                files.push(path);
            }
        }
    }

    // 5. 核心处理循环
    for file in files {
        let ext = file
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        if file.to_string_lossy().contains("_processed") {
            continue;
        }

        let file_display = file.file_name().unwrap().to_string_lossy();

        match ext.as_str() {
            "pdf" | "docx" => {
                print!("正在处理 {}: {}... ", ext.to_uppercase(), file_display);

                // 执行清理模式或直接复制
                let target_path = if mode == "1" || mode == "3" {
                    match ext.as_str() {
                        "pdf" => process_pdf(&file).ok(),
                        "docx" => process_docx(&file).ok(),
                        _ => None,
                    }
                } else {
                    // 仅修改时间模式下，生成一个副本以保持一致性
                    let mut out = file.clone();
                    out.set_file_name(format!(
                        "{}_processed.{}",
                        file.file_stem().unwrap().to_string_lossy(),
                        ext
                    ));
                    fs::copy(&file, &out).ok().map(|_| out)
                };

                // 执行时间修改
                if let Some(path) = target_path {
                    if let Some((c, m, a)) = times {
                        let _ = apply_precise_timestamps(&path, c, a, m);
                    }
                    println!("✅");
                } else {
                    println!("❌ 处理失败");
                }
            }
            _ => {}
        }
    }

    println!("\n✨ 任务执行完毕！");
    wait_for_keypress();
    Ok(())
}

// ================= 系统调用部分 (Win32 API) =================

fn apply_precise_timestamps(path: &Path, c: FileTime, a: FileTime, m: FileTime) -> io::Result<()> {
    #[cfg(windows)]
    {
        use std::os::windows::io::AsRawHandle;
        let file = fs::OpenOptions::new().write(true).open(path)?;
        let handle = file.as_raw_handle() as *mut std::ffi::c_void;

        let to_win_ft = |ft: FileTime| {
            let unix_nanos = ft.unix_seconds() * 1_000_000_000 + ft.nanoseconds() as i64;
            let intervals = unix_nanos / 100 + 116_444_736_000_000_000;
            WinFileTime {
                dw_low: (intervals & 0xFFFFFFFF) as u32,
                dw_high: (intervals >> 32) as u32,
            }
        };

        #[repr(C)]
        struct WinFileTime {
            dw_low: u32,
            dw_high: u32,
        }

        unsafe extern "system" {
            fn SetFileTime(
                h: *mut std::ffi::c_void,
                c: *const WinFileTime,
                a: *const WinFileTime,
                m: *const WinFileTime,
            ) -> i32;
        }

        let (ft_c, ft_a, ft_m) = (to_win_ft(c), to_win_ft(a), to_win_ft(m));
        unsafe {
            if SetFileTime(handle, &ft_c, &ft_a, &ft_m) == 0 {
                return Err(io::Error::last_os_error());
            }
        }
    }
    Ok(())
}

// ================= 文档清理部分 =================

fn process_docx(input_path: &Path) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!(
        "{}_processed.docx",
        input_path.file_stem().unwrap().to_string_lossy()
    ));

    let file = fs::File::open(input_path)?;
    let mut archive = ZipArchive::new(file)?;
    let out_file = fs::File::create(&output_path)?;
    let mut zip_out = ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let mut inner_file = archive.by_index(i)?;
        let options: FileOptions<'_, ()> =
            FileOptions::default().compression_method(inner_file.compression());
        zip_out.start_file(inner_file.name(), options)?;

        let mut buffer = Vec::new();
        inner_file.read_to_end(&mut buffer)?;

        if inner_file.name() == "docProps/core.xml" || inner_file.name() == "docProps/app.xml" {
            let mut content = String::from_utf8_lossy(&buffer).into_owned();
            let tags = [
                "dc:creator",
                "cp:lastModifiedBy",
                "cp:keywords",
                "dc:description",
                "dc:title",
                "Company",
                "Manager",
            ];
            for tag in tags {
                let s = format!("<{}>", tag);
                let e = format!("</{}>", tag);
                if let (Some(si), Some(ei)) = (content.find(&s), content.find(&e)) {
                    content.replace_range((si + s.len())..ei, "");
                }
            }
            zip_out.write_all(content.as_bytes())?;
        } else {
            zip_out.write_all(&buffer)?;
        }
    }
    zip_out.finish()?;
    Ok(output_path)
}

fn process_pdf(input_path: &Path) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!(
        "{}_processed.pdf",
        input_path.file_stem().unwrap().to_string_lossy()
    ));
    let mut doc = Document::load(input_path)?;
    if let Ok(info_id) = doc
        .trailer
        .get(b"Info")
        .and_then(|info| info.as_reference())
    {
        if let Ok(dict) = doc.get_object_mut(info_id).and_then(|o| o.as_dict_mut()) {
            let keys: &[&[u8]] = &[
                b"Author",
                b"Creator",
                b"Producer",
                b"Title",
                b"Subject",
                b"Keywords",
                b"Company",
                b"CreationDate",
                b"ModDate",
            ];
            for &k in keys {
                dict.remove(k);
            }
        }
    }
    doc.save(&output_path)?;
    Ok(output_path)
}

// ================= 辅助函数 =================

fn ask_time(label: &str) -> FileTime {
    print!("{}: ", label);
    io::stdout().flush().unwrap();
    let mut s = String::new();
    io::stdin().read_line(&mut s).unwrap();
    let s = s.trim();
    if s.is_empty() {
        FileTime::now()
    } else {
        match NaiveDateTime::parse_from_str(s, "%Y-%m-%d %H:%M:%S") {
            Ok(dt) => {
                let local_dt = Local.from_local_datetime(&dt).single().expect("解析错误");
                FileTime::from_unix_time(local_dt.timestamp(), 0)
            }
            Err(_) => FileTime::now(),
        }
    }
}

fn wait_for_keypress() {
    println!("\n按下回车键退出...");
    let _ = io::stdin().read_line(&mut String::new());
}
