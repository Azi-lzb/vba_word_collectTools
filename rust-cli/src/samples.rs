//! Sample forms generated directly as OOXML, without Word or Python.
use anyhow::{Context, Result, ensure};
use clap::ValueEnum;
use std::{
    collections::BTreeMap,
    fs::File,
    io::{Read, Write},
    path::{Path, PathBuf},
};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

#[derive(Clone, Copy, Debug, ValueEnum)]
pub enum SampleKind {
    Legacy,
    ContentControl,
    Both,
}

fn legacy(index: usize) -> String {
    fn field(name: &str, value: &str) -> String {
        format!(
            r#"<w:r><w:fldChar w:fldCharType="begin"><w:ffData><w:name w:val="{name}"/><w:enabled/><w:calcOnExit w:val="0"/><w:textInput/></w:ffData></w:fldChar></w:r><w:r><w:instrText xml:space="preserve"> FORMTEXT </w:instrText></w:r><w:r><w:fldChar w:fldCharType="separate"/></w:r><w:r><w:t>{value}</w:t></w:r><w:r><w:fldChar w:fldCharType="end"/></w:r>"#
        )
    }
    format!(
        r#"<w:p><w:r><w:t xml:space="preserve">姓名： </w:t></w:r>{}</w:p><w:p><w:r><w:t xml:space="preserve">意见： </w:t></w:r>{}</w:p>"#,
        field("UserName", &format!("用户_{index}")),
        field("Feedback", &format!("示例意见 {index}"))
    )
}

fn content_control(index: usize) -> String {
    format!(
        r#"<w:p><w:r><w:t xml:space="preserve">姓名： </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="姓名"/><w:tag w:val="UserName"/><w:id w:val="{id}"/><w:text/></w:sdtPr><w:sdtContent><w:r><w:t>测试员工_{index}</w:t></w:r></w:sdtContent></w:sdt></w:p><w:p><w:r><w:t xml:space="preserve">部门： </w:t></w:r><w:sdt><w:sdtPr><w:alias w:val="部门"/><w:tag w:val="Department"/><w:id w:val="{id2}"/><w:dropDownList><w:listItem w:displayText="财务" w:value="Fin"/><w:listItem w:displayText="技术部" w:value="Tech"/></w:dropDownList></w:sdtPr><w:sdtContent><w:r><w:t>技术部</w:t></w:r></w:sdtContent></w:sdt></w:p>"#,
        id = index * 10,
        id2 = index * 10 + 1
    )
}

fn protect(path: &Path) -> Result<()> {
    // The initial DOCX has already been saved; add protection and save again.
    let mut source = ZipArchive::new(File::open(path)?)?;
    let temp = tempfile::NamedTempFile::new_in(path.parent().unwrap_or(Path::new(".")))?;
    let mut output = ZipWriter::new(temp.reopen()?);
    let opts = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for index in 0..source.len() {
        let mut entry = source.by_index(index)?;
        let name = entry.name().to_owned();
        let mut bytes = Vec::new();
        entry.read_to_end(&mut bytes)?;
        if name == "[Content_Types].xml" {
            let xml=String::from_utf8(bytes)?.replace("</Types>","<Override PartName=\"/word/settings.xml\" ContentType=\"application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml\"/></Types>");
            bytes = xml.into_bytes();
        } else if name == "word/_rels/document.xml.rels" {
            let xml=String::from_utf8(bytes)?.replace("</Relationships>","<Relationship Id=\"rIdSettings\" Type=\"http://schemas.openxmlformats.org/officeDocument/2006/relationships/settings\" Target=\"settings.xml\"/></Relationships>");
            bytes = xml.into_bytes();
        }
        output.start_file(name, opts)?;
        output.write_all(&bytes)?;
    }
    output.start_file("word/settings.xml", opts)?;
    output.write_all(br#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><w:settings xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:documentProtection w:edit="forms" w:enforcement="1"/></w:settings>"#)?;
    output.finish()?;
    drop(source);
    temp.persist(path).map_err(|e| e.error)?;
    Ok(())
}

pub fn generate(kind: SampleKind, directory: &Path) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for index in 1..=3 {
        if matches!(kind, SampleKind::Legacy | SampleKind::Both) {
            files.push((
                directory.join(format!("旧式窗体示例_{index}.docx")),
                legacy(index),
            ));
        }
        if matches!(kind, SampleKind::ContentControl | SampleKind::Both) {
            files.push((
                directory.join(format!("内容控件示例_{index}.docx")),
                content_control(index),
            ));
        }
    }
    for (path, _) in &files {
        ensure!(
            !path.exists(),
            "示例文件已存在，请选择其他文件夹：{}",
            path.display()
        );
    }
    std::fs::create_dir_all(directory)
        .with_context(|| format!("无法创建文件夹 {}", directory.display()))?;
    let mut generated = Vec::new();
    for (path, body) in files {
        crate::native::write_minimal_docx(&path, &body, &BTreeMap::new(), "")?;
        protect(&path).with_context(|| format!("无法保护示例 {}", path.display()))?;
        generated.push(path);
    }
    Ok(generated)
}
