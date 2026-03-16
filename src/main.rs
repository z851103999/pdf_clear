use std::path::{Path, PathBuf};
use lopdf::Document;
use std::fs;
use std::io::{self, Read, Write};
use zip::{ZipArchive, ZipWriter, write::FileOptions};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("========================================");
    println!("   PDF & Word 批量元数据清理工具");
    println!("========================================");

    // --- 1. 获取用户输入的路径 ---
    print!("请输入文件或文件夹路径: ");
    io::stdout().flush()?;
    
    let mut input_str = String::new();
    io::stdin().read_line(&mut input_str)?;
    // 处理路径：去除两端的空格，并去掉因拖拽文件可能产生的双引号
    let input_path = PathBuf::from(input_str.trim().replace("\"", ""));

    if !input_path.exists() {
        println!("❌ 错误：路径不存在！");
        wait_for_keypress();
        return Ok(());
    }

    // --- 2. 收集所有待处理的文件 ---
    let mut files = Vec::new();
    if input_path.is_file() {
        files.push(input_path);
    } else {
        // 如果是文件夹，遍历其中的所有文件
        for entry in fs::read_dir(&input_path)? {
            let path = entry?.path();
            if path.is_file() { files.push(path); }
        }
    }

    // --- 3. 循环处理每一个文件 ---
    for file in files {
        let ext = file.extension().and_then(|s| s.to_str()).unwrap_or("").to_lowercase();
        // 自动跳过已经清理过的文件，避免重复处理
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
            _ => {} // 忽略其他格式
        }
    }

    println!("\n所有任务已完成！");
    wait_for_keypress();
    Ok(())
}

/// Word (DOCX) 处理函数
/// 原理：DOCX 是 ZIP 包，元数据存在 docProps 文件夹下的 XML 中
fn process_docx(input_path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!("{}_cleaned.docx", input_path.file_stem().unwrap().to_string_lossy()));

    let file = fs::File::open(input_path)?;
    let mut archive = ZipArchive::new(file)?;
    
    let out_file = fs::File::create(&output_path)?;
    let mut zip_out = ZipWriter::new(out_file);

    // 遍历原始压缩包中的每一个文件
    for i in 0..archive.len() {
        let mut inner_file = archive.by_index(i)?;
        let name = inner_file.name().to_string();
        
        // 必须显式指定 FileOptions<()> 兼容 zip 库新版本的泛型要求
        let options: FileOptions<()> = FileOptions::default()
            .compression_method(inner_file.compression());

        zip_out.start_file(&name, options)?;

        // 如果是存储元数据的 XML 文件，则进行内容清洗
        if name == "docProps/core.xml" || name == "docProps/app.xml" {
            let mut content = String::new();
            inner_file.read_to_string(&mut content)?;
            
            // 需要清空的 XML 标签列表
            let tags = [
                "dc:creator", "cp:lastModifiedBy", "cp:keywords", 
                "dc:description", "dc:title", "Company", "Manager"
            ];

            let mut final_xml = content;
            for tag in tags {
                let start_tag = format!("<{}>", tag);
                let end_tag = format!("</{}>", tag);
                
                // 定位标签位置：如果找到成对的标签，则清空中间的文本
                if let Some(start_idx) = final_xml.find(&start_tag) {
                    if let Some(end_idx) = final_xml.find(&end_tag) {
                        let content_start = start_idx + start_tag.len();
                        if content_start < end_idx {
                            // 保持标签结构不变，仅抹除内容
                            final_xml.replace_range(content_start..end_idx, "");
                        }
                    }
                }
            }
            zip_out.write_all(final_xml.as_bytes())?;
        } else {
            // 对于非元数据文件（如正文文字、图片等），直接原样拷贝
            let mut buffer = Vec::new();
            inner_file.read_to_end(&mut buffer)?;
            zip_out.write_all(&buffer)?;
        }
    }
    zip_out.finish()?;
    Ok(())
}

/// PDF 处理函数
/// 原理：删除 Trailer 中的 Info 字典键值对，并物理删除 Root 下的 Metadata 对象流
fn process_pdf(input_path: &Path) -> Result<(), Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!("{}_cleaned.pdf", input_path.file_stem().unwrap().to_string_lossy()));
    
    let mut doc = Document::load(input_path)?;
    
    // 1. 清理 Info 字典 (兼容不同版本的 Result/Option 嵌套)
    if let Some(info_id) = doc.trailer.get(b"Info").ok().and_then(|info| info.as_reference().ok()) {
        if let Ok(info_dict) = doc.get_object_mut(info_id).and_then(|obj| obj.as_dict_mut()) {
            let keys: &[&[u8]] = &[b"Author", b"Creator", b"Producer", b"Title", b"Subject", b"Keywords", b"Company"];
            for &k in keys { 
                info_dict.remove(k); 
            }
        }
    }

    // 2. 清理现代 PDF 常用 XML Metadata (存储在 Root 节点)
    let mut meta_id = None;
    // 先查找 Root 字典中是否存在 Metadata 引用
    if let Some(root_id) = doc.trailer.get(b"Root").ok().and_then(|o| o.as_reference().ok()) {
        if let Ok(root) = doc.get_object(root_id).and_then(|o| o.as_dict()) {
            meta_id = root.get(b"Metadata").ok().and_then(|o| o.as_reference().ok());
        }
        
        // 如果找到了 Metadata 对象 ID
        if let Some(id) = meta_id {
            // 首先从 Root 字典中移除该键
            if let Ok(root_mut) = doc.get_object_mut(root_id).and_then(|o| o.as_dict_mut()) {
                root_mut.remove(b"Metadata");
            }
            // 然后在 PDF 数据库中彻底物理删除该对象
            doc.objects.remove(&id);
        }
    }

    doc.save(&output_path)?;
    Ok(())
}

/// 保持窗口开启，防止程序运行完立即闪退
fn wait_for_keypress() {
    println!("\n按下回车键退出...");
    let _ = io::stdin().read_line(&mut String::new());
}