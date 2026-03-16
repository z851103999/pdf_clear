use std::path::{Path, PathBuf};
use lopdf::Document;
use std::fs;
use std::io::{self, Write};

fn main() -> Result<(), Box<dyn std::error::Error>> {
    println!("========================================");
    println!("     PDF 批量元数据清理工具 (v1.1)");
    println!("========================================");

    // --- 交互环节 ---
    print!("请输入 PDF 文件或文件夹的路径: ");
    io::stdout().flush()?;
    
    let mut input_str = String::new();
    io::stdin().read_line(&mut input_str)?;
    let input_path = PathBuf::from(input_str.trim().replace("\"", "")); // 处理拖入文件时可能带有的双引号

    if !input_path.exists() {
        println!("❌ 错误：路径不存在！");
        wait_for_keypress();
        return Ok(());
    }

    // --- 收集待处理文件 ---
    let mut files_to_process = Vec::new();
    if input_path.is_file() {
        if is_pdf(&input_path) {
            files_to_process.push(input_path);
        }
    } else {
        // 扫描文件夹
        for entry in fs::read_dir(&input_path)? {
            let path = entry?.path();
            if path.is_file() && is_pdf(&path) {
                // 排除已经是 _cleaned 的文件，防止循环处理
                if !path.to_string_lossy().contains("_cleaned.pdf") {
                    files_to_process.push(path);
                }
            }
        }
    }

    if files_to_process.is_empty() {
        println!("⚠️ 未发现可处理的 PDF 文件。");
        wait_for_keypress();
        return Ok(());
    }

    println!("🚀 找到 {} 个文件，准备开始清理...", files_to_process.len());

    // --- 循环处理 ---
    let mut success_count = 0;
    for file in files_to_process {
        match process_pdf(&file) {
            Ok(out) => {
                println!("✅ 已清理: {}", out.file_name().unwrap().to_string_lossy());
                success_count += 1;
            }
            Err(e) => println!("❌ 失败: {} (原因: {})", file.display(), e),
        }
    }

    println!("\n任务结束！成功处理 {} 个文件。", success_count);
    wait_for_keypress();
    Ok(())
}

/// 核心处理逻辑：清理单个 PDF
fn process_pdf(input_path: &Path) -> Result<PathBuf, Box<dyn std::error::Error>> {
    let mut output_path = input_path.to_path_buf();
    output_path.set_file_name(format!(
        "{}_cleaned.pdf",
        input_path.file_stem().unwrap_or_default().to_string_lossy()
    ));

    let mut doc = Document::load(input_path)?;

    // 1. 清理 Info 字典
    if let Some(info_id) = doc.trailer.get(b"Info").and_then(|info| info.as_reference()).ok() {
        if let Ok(info_dict) = doc.get_object_mut(info_id).and_then(|obj| obj.as_dict_mut()) {
            let keys_to_remove: &[&[u8]] = &[
                b"Title", b"Author", b"Subject", b"Keywords",
                b"Creator", b"Producer", b"CreationDate", b"ModDate", b"Trapped"
            ];
            for &key in keys_to_remove {
                info_dict.remove(key);
            }
        }
    }

    // 2. 清理 Root 中的 Metadata 流
    let mut metadata_id = None;
    if let Some(root_id) = doc.trailer.get(b"Root").and_then(|obj| obj.as_reference()).ok() {
        if let Ok(root_dict) = doc.get_object(root_id).and_then(|obj| obj.as_dict()) {
            metadata_id = root_dict.get(b"Metadata").and_then(|obj| obj.as_reference()).ok();
        }
        if let Some(id) = metadata_id {
            if let Ok(root_dict) = doc.get_object_mut(root_id).and_then(|obj| obj.as_dict_mut()) {
                root_dict.remove(b"Metadata");
            }
            doc.objects.remove(&id);
        }
    }

    doc.save(&output_path)?;
    Ok(output_path)
}

fn is_pdf(path: &Path) -> bool {
    path.extension().and_then(|s| s.to_str()).map(|s| s.to_lowercase() == "pdf").unwrap_or(false)
}

fn wait_for_keypress() {
    println!("\n按下回车键退出...");
    let _ = io::stdin().read_line(&mut String::new());
}