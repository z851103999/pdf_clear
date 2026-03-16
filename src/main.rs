use std::path::{Path, PathBuf};
use lopdf::Document;
use std::fs;
use std::io::{self, Read, Write};
use zip::{ZipArchive, ZipWriter, write::FileOptions};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("========================================");
    println!("   PDF & Word 批量元数据清理工具");
    println!("========================================");

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

    let mut files = Vec::new();
    if input_path.is_file() {
        files.push(input_path);
    } else {
        for entry in fs::read_dir(&input_path)? {
            let path = entry?.path();
            if path.is_file() { files.push(path); }
        }
    }

    for file in files {
        let ext = file.extension().and_then(|s| s.to_str()).unwrap_or("").to_lowercase();
        if file.to_string_lossy().contains("_cleaned") { continue; }

        match ext.as_str() {
            "pdf" => {
                print!("正在清理 PDF: {}... ", file.file_name().unwrap().to_string_lossy());
                if let Ok(_) = process_pdf(&file) { println!("✅"); }
            },
            "docx" => {
                print!("正在清理 Word: {}... ", file.file_name().unwrap().to_string_lossy());
                if let Ok(_) = process_docx(&file) { println!("✅"); }
            },
            _ => {}
        }
    }

    println!("\n所有任务已完成！");
    wait_for_keypress();
    Ok(())
}

fn process_docx(input_path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!("{}_cleaned.docx", input_path.file_stem().unwrap().to_string_lossy()));

    let file = fs::File::open(input_path)?;
    let mut archive = ZipArchive::new(file)?;
    
    let out_file = fs::File::create(&output_path)?;
    let mut zip_out = ZipWriter::new(out_file);

    for i in 0..archive.len() {
        let mut inner_file = archive.by_index(i)?;
        let name = inner_file.name().to_string();
        
        // 修复：显式指定 FileOptions 的类型为 ()
        let options: FileOptions<()> = FileOptions::default()
            .compression_method(inner_file.compression());

        zip_out.start_file(&name, options)?;

        if name == "docProps/core.xml" || name == "docProps/app.xml" {
            let mut content = String::new();
            inner_file.read_to_string(&mut content)?;
            
            let tags = [
                "dc:creator", "cp:lastModifiedBy", "cp:keywords", 
                "dc:description", "dc:title", "Company", "Manager"
            ];

            let mut final_xml = content;
            for tag in tags {
                let start_tag = format!("<{}>", tag);
                let end_tag = format!("</{}>", tag);
                
                if let Some(start_idx) = final_xml.find(&start_tag) {
                    if let Some(end_idx) = final_xml.find(&end_tag) {
                        let content_start = start_idx + start_tag.len();
                        if content_start < end_idx {
                            final_xml.replace_range(content_start..end_idx, "");
                        }
                    }
                }
            }
            zip_out.write_all(final_xml.as_bytes())?;
        } else {
            let mut buffer = Vec::new();
            inner_file.read_to_end(&mut buffer)?;
            zip_out.write_all(&buffer)?;
        }
    }
    zip_out.finish()?;
    Ok(())
}

fn process_pdf(input_path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!("{}_cleaned.pdf", input_path.file_stem().unwrap().to_string_lossy()));
    
    let mut doc = Document::load(input_path)?;
    
    // 修复：正确处理 Result 和 Option 的嵌套
    // 使用 .ok() 将 Result 转为 Option，使 if let Some 能够工作
    if let Some(info_id) = doc.trailer.get(b"Info").ok().and_then(|info| info.as_reference().ok()) {
        if let Ok(info_dict) = doc.get_object_mut(info_id).and_then(|obj| obj.as_dict_mut()) {
            let keys: &[&[u8]] = &[b"Author", b"Creator", b"Producer", b"Title", b"Subject", b"Keywords", b"Company"];
            for &k in keys { info_dict.remove(k); }
        }
    }

    let mut meta_id = None;
    if let Some(root_id) = doc.trailer.get(b"Root").ok().and_then(|o| o.as_reference().ok()) {
        if let Ok(root) = doc.get_object(root_id).and_then(|o| o.as_dict()) {
            meta_id = root.get(b"Metadata").ok().and_then(|o| o.as_reference().ok());
        }
        if let Some(id) = meta_id {
            if let Ok(root_mut) = doc.get_object_mut(root_id).and_then(|o| o.as_dict_mut()) {
                root_mut.remove(b"Metadata");
            }
            doc.objects.remove(&id);
        }
    }

    doc.save(&output_path)?;
    Ok(())
}

fn wait_for_keypress() {
    println!("\n按下回车键退出...");
    let _ = io::stdin().read_line(&mut String::new());
}