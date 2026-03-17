use chrono::{Local, NaiveDateTime, TimeZone};
use filetime::{self, FileTime};
use lopdf::{Document, Object};
use rust_xlsxwriter::*;
use std::collections::BTreeMap;
use std::fs;
use std::io::{self, Read, Write};
use std::path::{Path, PathBuf};
use zip::{ZipArchive, ZipWriter, write::FileOptions};

// --- 纸张标准尺寸常量 (单位: mm) ---
const A0_SHORT: f64 = 841.0;
const A0_LONG: f64 = 1189.0;
const A1_SHORT: f64 = 594.0;
const A1_LONG: f64 = 841.0;
const A2_SHORT: f64 = 420.0;
const A2_LONG: f64 = 594.0;
const A3_SHORT: f64 = 297.0;
const A3_LONG: f64 = 420.0;
const A4_SHORT: f64 = 210.0;
const A4_LONG: f64 = 297.0;

/// 存储单个 PDF 的统计结果
#[derive(Default, Debug)]
struct PdfStat {
    filename: String,
    total_pages: u32,
    size_distribution: BTreeMap<String, u32>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("========================================");
    println!("   工程图纸统计与文档清理工具 v5.0");
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
    println!(" [1] 仅统计图纸 (导出Excel，不改动文件)");
    println!(" [2] 仅清理元数据 (仅针对PDF/Word)");
    println!(" [3] 仅修改系统时间 (创建/修改/访问)");
    println!(" [4] 全套处理 (统计+清理+修改时间)");
    print!("请输入选项 (1/2/3/4): ");
    io::stdout().flush()?;
    let mut mode_str = String::new();
    io::stdin().read_line(&mut mode_str)?;
    let mode = mode_str.trim();

    // 3. 时间设置逻辑
    let mut times = None;
    if mode == "3" || mode == "4" {
        println!("\n--- 系统时间修改设置 (格式: 2023-01-01 12:00:00) ---");
        let t_create = ask_time("设置创建时间");
        let t_modify = ask_time("设置修改时间");
        let t_access = ask_time("设置访问时间");
        times = Some((t_create, t_modify, t_access));
    }

    // 4. 扫描文件
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

    let mut all_stats = Vec::new();
    println!("\n--- 处理日志 ---");

    for file in files {
        let ext = file
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or("")
            .to_lowercase();
        if file.to_string_lossy().contains("_processed") {
            continue;
        }
        let file_display = file.file_name().unwrap().to_string_lossy().into_owned();

        match ext.as_str() {
            "pdf" => {
                // 执行统计 (模式 1, 4)
                if mode == "1" || mode == "4" {
                    if let Ok(stat) = analyze_pdf_sizes(&file) {
                        println!("[日志] 统计完成: {} ({}页)", file_display, stat.total_pages);
                        all_stats.push(stat);
                    }
                }
                // 执行清理 (模式 2, 4)
                if mode == "2" || mode == "4" {
                    if let Ok(out) = process_pdf(&file) {
                        println!(
                            "[日志] 已清理元数据并生成: {}",
                            out.file_name().unwrap().to_string_lossy()
                        );
                        if let Some((c, m, a)) = times {
                            let _ = apply_precise_timestamps(&out, c, a, m);
                        }
                    }
                }
                // 直接改时间 (模式 3)
                if mode == "3" {
                    if let Some((c, m, a)) = times {
                        let _ = apply_precise_timestamps(&file, c, a, m);
                        println!("[日志] 已更新文件时间: {}", file_display);
                    }
                }
            }
            "docx" => {
                if mode == "2" || mode == "4" {
                    if let Ok(out) = process_docx(&file) {
                        println!(
                            "[日志] 已清理Word并生成: {}",
                            out.file_name().unwrap().to_string_lossy()
                        );
                        if let Some((c, m, a)) = times {
                            let _ = apply_precise_timestamps(&out, c, a, m);
                        }
                    }
                }
                if mode == "3" {
                    if let Some((c, m, a)) = times {
                        let _ = apply_precise_timestamps(&file, c, a, m);
                        println!("[日志] 已更新文件时间: {}", file_display);
                    }
                }
            }
            _ => {}
        }
    }

    // 5. 生成 Excel 统计表
    if !all_stats.is_empty() {
        export_stats_to_excel(&all_stats)?;
        println!("\n✅ 统计报表已导出: 图纸尺寸统计结果.xlsx");
    }

    println!("\n✨ 任务全部执行完毕！");
    wait_for_keypress();
    Ok(())
}

// ================= 核心：图纸尺寸识别 (修复 as_f64 报错) =================

fn analyze_pdf_sizes(path: &Path) -> Result<PdfStat, Box<dyn std::error::Error>> {
    let doc = Document::load(path)?;
    let mut stat = PdfStat {
        filename: path.file_name().unwrap().to_string_lossy().into_owned(),
        total_pages: doc.get_pages().len() as u32,
        size_distribution: BTreeMap::new(),
    };

    // 辅助闭包：安全地提取 PDF 对象中的数值
    let get_val = |obj: &Object| -> f64 {
        match obj {
            Object::Real(f) => *f as f64,
            Object::Integer(i) => *i as f64,
            _ => 0.0,
        }
    };

    for page_id in doc.get_pages().values() {
        if let Ok(page) = doc.get_object(*page_id).and_then(|o| o.as_dict()) {
            if let Ok(mb) = page.get(b"MediaBox").and_then(|o| o.as_array()) {
                // 解决之前 as_f64() 不存在的问题
                let x1 = get_val(&mb[0]);
                let y1 = get_val(&mb[1]);
                let x2 = get_val(&mb[2]);
                let y2 = get_val(&mb[3]);

                let w = (x2 - x1).abs() / 2.8346; // pt 转 mm
                let h = (y2 - y1).abs() / 2.8346;

                let (short, long) = if w < h { (w, h) } else { (h, w) };
                let size_key = identify_drawing_size(short, long);
                *stat.size_distribution.entry(size_key).or_insert(0) += 1;
            }
        }
    }
    Ok(stat)
}

fn identify_drawing_size(short: f64, long: f64) -> String {
    let (base_name, base_long) = if (short - A0_SHORT).abs() < 20.0 {
        ("A0", A0_LONG)
    } else if (short - A1_SHORT).abs() < 15.0 {
        ("A1", A1_LONG)
    } else if (short - A2_SHORT).abs() < 10.0 {
        ("A2", A2_LONG)
    } else if (short - A3_SHORT).abs() < 8.0 {
        ("A3", A3_LONG)
    } else if (short - A4_SHORT).abs() < 5.0 {
        ("A4", A4_LONG)
    } else {
        return "非标/其他".to_string();
    };

    let ratio = long / base_long;
    let multiplier = (ratio * 4.0).round() / 4.0; // 0.25 步进

    if (multiplier - 1.0).abs() < 0.125 {
        base_name.to_string()
    } else {
        format!("{}x{}", base_name, multiplier)
    }
}

// ================= Excel 导出 =================

fn export_stats_to_excel(stats: &[PdfStat]) -> Result<(), Box<dyn std::error::Error>> {
    let mut workbook = Workbook::new();
    let worksheet = workbook.add_worksheet();

    let mut size_headers = std::collections::BTreeSet::new();
    for s in stats {
        for key in s.size_distribution.keys() {
            size_headers.insert(key.clone());
        }
    }
    let sorted_headers: Vec<_> = size_headers.into_iter().collect();

    let head_fmt = Format::new()
        .set_bold()
        .set_bg_color(Color::Silver)
        .set_border(FormatBorder::Thin);

    worksheet.write_with_format(0, 0, "文件名", &head_fmt)?;
    worksheet.write_with_format(0, 1, "总页数", &head_fmt)?;
    for (i, header) in sorted_headers.iter().enumerate() {
        worksheet.write_with_format(0, (i + 2) as u16, header, &head_fmt)?;
    }

    for (row_idx, stat) in stats.iter().enumerate() {
        let r = (row_idx + 1) as u32;
        worksheet.write(r, 0, &stat.filename)?;
        worksheet.write(r, 1, stat.total_pages)?;
        for (col_idx, header) in sorted_headers.iter().enumerate() {
            let count = stat.size_distribution.get(header).unwrap_or(&0);
            worksheet.write(r, (col_idx + 2) as u16, *count)?;
        }
    }

    workbook.save("图纸尺寸统计结果.xlsx")?;
    Ok(())
}

// ================= Win32 时间修改 =================

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
            SetFileTime(handle, &ft_c, &ft_a, &ft_m);
        }
    }
    Ok(())
}

// ================= 文档清理 =================

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
                let local_dt = Local.from_local_datetime(&dt).single().expect("解析失败");
                FileTime::from_unix_time(local_dt.timestamp(), 0)
            }
            Err(_) => {
                println!("  ⚠️ 格式错误，默认使用当前时间");
                FileTime::now()
            }
        }
    }
}

fn wait_for_keypress() {
    println!("\n按下回车键退出...");
    let _ = io::stdin().read_line(&mut String::new());
}
