use std::path::PathBuf;
use clap::Parser;
use lopdf::Document;
use std::fs;

#[derive(Parser, Debug)]
#[command(name = "pdf_clear")]
struct Args {
    #[arg(required = true)]
    input: PathBuf,

    #[arg(short, long)]
    output: Option<PathBuf>,
}

fn main() -> Result<(), Box<dyn std::error::Error>> {
    let args = Args::parse();

    if !args.input.exists() {
        return Err(format!("输入文件不存在: {}", args.input.display()).into());
    }

    let output_path = args.output.unwrap_or_else(|| {
        let mut path = args.input.clone();
        path.set_file_name(format!(
            "{}_cleaned.pdf",
            path.file_stem().unwrap_or_default().to_string_lossy()
        ));
        path
    });

    let mut doc = Document::load(&args.input)
        .map_err(|e| format!("无法加载PDF文件: {}", e))?;

    // --- 1. 清理 Info 字典 ---
    if let Some(info_id) = doc.trailer.get(b"Info").and_then(|info| info.as_reference()).ok() {
        if let Ok(info_dict) = doc.get_object_mut(info_id).and_then(|obj| obj.as_dict_mut()) {
            // 显式指定类型以解决 [u8; N] 长度不匹配问题
            let keys_to_remove: &[&[u8]] = &[
                b"Title", b"Author", b"Subject", b"Keywords",
                b"Creator", b"Producer", b"CreationDate", b"ModDate",
                b"Trapped", b"Company", b"Manager", b"Category"
            ];
            for &key in keys_to_remove {
                info_dict.remove(key);
            }
        }
    }

    // --- 2. 清理 Root 中的 Metadata (分步操作以避免多次可变借用) ---
    let mut metadata_id_to_delete = None;

    // 步骤 A: 查找 ID (只读借用)
    if let Some(root_id) = doc.trailer.get(b"Root").and_then(|obj| obj.as_reference()).ok() {
        if let Ok(root_dict) = doc.get_object(root_id).and_then(|obj| obj.as_dict()) {
            metadata_id_to_delete = root_dict.get(b"Metadata").and_then(|obj| obj.as_reference()).ok();
        }

        // 步骤 B: 从字典中移除键 (可变借用)
        if metadata_id_to_delete.is_some() {
            if let Ok(root_dict) = doc.get_object_mut(root_id).and_then(|obj| obj.as_dict_mut()) {
                root_dict.remove(b"Metadata");
            }
        }
    }

    // 步骤 C: 物理删除对象 (此时对 doc 的其他借用已结束)
    if let Some(id) = metadata_id_to_delete {
        doc.objects.remove(&id);
        println!("✅ 已物理删除 Metadata 对象流");
    }

    // --- 3. 保存文件 ---
    if let Some(parent) = output_path.parent() {
        if !parent.exists() { fs::create_dir_all(parent)?; }
    }

    doc.save(&output_path)?;
    println!("✅ 处理完成: {}", output_path.display());

    Ok(())
}