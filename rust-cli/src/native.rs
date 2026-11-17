use std::{
    collections::{BTreeMap, HashMap, HashSet},
    fs::{self, File},
    io::Read,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result, ensure};
use roxmltree::{Document, Node};
use rust_xlsxwriter::{Format, Workbook};
use zip::{ZipArchive, ZipWriter, write::SimpleFileOptions};

use crate::{
    config::Config,
    model::{FormKind, OutputLayout, Priority, Record, Report, Source, TableData, TableMode},
};

const W_NS: &str = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const R_NS: &str = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";

pub fn extract(inputs: &[PathBuf], form_kind: FormKind, config: &Config) -> Result<Report> {
    let paths = expand_inputs(inputs)?;
    ensure!(
        !paths.is_empty(),
        "No .docx or .docm input documents were found"
    );

    let mut report = Report {
        form_kind: Some(form_kind),
        template_path: None,
        ..Report::default()
    };

    for path in paths {
        report.sources.push(Source {
            file: path.to_string_lossy().into_owned(),
            display: if form_kind == FormKind::Legacy {
                mapped_name(&path, config)
            } else {
                path.file_name()
                    .unwrap_or_default()
                    .to_string_lossy()
                    .into_owned()
            },
        });
        let extension = path
            .extension()
            .and_then(|s| s.to_str())
            .unwrap_or_default();
        ensure!(
            extension.eq_ignore_ascii_case("docx") || extension.eq_ignore_ascii_case("docm"),
            "Pure Rust supports .docx and .docm, not {}",
            path.display()
        );
        if report.template_path.is_none() {
            report.template_path = Some(path.to_string_lossy().to_string());
        }
        let xml = read_document_xml(&path)?;
        let document = Document::parse(&xml)
            .with_context(|| format!("Malformed WordprocessingML in {}", path.display()))?;
        for namespace in document.root_element().namespaces() {
            let prefix = namespace.name().unwrap_or_default().to_owned();
            report
                .namespaces
                .entry(prefix)
                .or_insert_with(|| namespace.uri().to_owned());
        }
        let records = match form_kind {
            FormKind::Legacy => extract_legacy(&xml, &document, &path, config)?,
            FormKind::ContentControl => extract_content_controls(&xml, &document, &path, config)?,
        };
        report.records.extend(records);
    }

    Ok(report)
}

fn expand_inputs(inputs: &[PathBuf]) -> Result<Vec<PathBuf>> {
    let mut paths = Vec::new();
    for input in inputs {
        if input.is_file() {
            paths.push(input.clone());
            continue;
        }
        ensure!(input.is_dir(), "Input does not exist: {}", input.display());
        let mut stack = vec![input.clone()];
        while let Some(directory) = stack.pop() {
            for entry in fs::read_dir(&directory)
                .with_context(|| format!("Could not read {}", directory.display()))?
            {
                let entry = entry?;
                let path = entry.path();
                if path.is_dir() {
                    stack.push(path);
                } else if is_word_file(&path) {
                    paths.push(path);
                }
            }
        }
    }
    let mut seen = HashSet::new();
    paths.retain(|p| seen.insert(p.clone()));
    Ok(paths)
}

fn is_word_file(path: &Path) -> bool {
    path.extension()
        .and_then(|s| s.to_str())
        .is_some_and(|s| s.eq_ignore_ascii_case("docx") || s.eq_ignore_ascii_case("docm"))
}

fn read_document_xml(path: &Path) -> Result<String> {
    let file = File::open(path).with_context(|| format!("Could not open {}", path.display()))?;
    let mut archive = ZipArchive::new(file)
        .with_context(|| format!("{} is not an OOXML Word package", path.display()))?;
    let mut xml = Vec::new();
    archive
        .by_name("word/document.xml")
        .with_context(|| format!("{} has no word/document.xml", path.display()))?
        .read_to_end(&mut xml)?;
    crate::xml::decode(&xml)
}

fn extract_legacy(
    xml: &str,
    document: &Document<'_>,
    path: &Path,
    config: &Config,
) -> Result<Vec<Record>> {
    let field_data: Vec<_> = document
        .descendants()
        .filter(|node| is_w(*node, "ffData"))
        .collect();
    let field_chars: Vec<_> = document
        .descendants()
        .filter(|node| is_w(*node, "fldChar"))
        .collect();
    let mut previous_end = 0usize;
    let mut records = Vec::new();
    let source_file = path.to_string_lossy().to_string();
    let display_source = mapped_name(path, config);

    for data in field_data {
        let Some(begin) = ancestor_field_char(data, "begin") else {
            continue;
        };
        let mut depth = 1;
        let mut separate = None;
        let mut end = None;
        for marker in field_chars
            .iter()
            .copied()
            .filter(|node| node.range().start > begin.range().start)
        {
            match field_char_type(marker) {
                Some("begin") => depth += 1,
                Some("separate") if depth == 1 => separate = Some(marker),
                Some("end") => {
                    depth -= 1;
                    if depth == 0 {
                        end = Some(marker);
                        break;
                    }
                }
                _ => {}
            }
        }
        let Some(end) = end else { continue };

        let field_name = child_value(data, "name")
            .unwrap_or("未命名窗体域")
            .to_owned();
        let (question_xml, question_text) =
            fragment_for_range(xml, document, previous_end, begin.range().start);
        let (mut answer_xml, mut answer_text) = fragment_for_range(
            xml,
            document,
            separate
                .map(|node| node.range().end)
                .unwrap_or(end.range().start),
            end.range().start,
        );
        if answer_xml.is_empty() {
            answer_xml = "<w:p></w:p>".into();
        }
        if let Some(checkbox) = child(data, "checkBox") {
            let value = child(checkbox, "checked").or_else(|| child(checkbox, "default"));
            let checked = value.is_some_and(|node| {
                !matches!(node.attribute((W_NS, "val")), Some("0" | "false" | "off"))
            });
            answer_text = if checked { "1" } else { "0" }.into();
        } else if let Some(dropdown) = child(data, "ddList") {
            let index = child_value(dropdown, "result")
                .or_else(|| child_value(dropdown, "default"))
                .and_then(|v| v.parse::<usize>().ok())
                .unwrap_or(0);
            if let Some(entry) = dropdown
                .children()
                .filter(|node| is_w(*node, "listEntry"))
                .nth(index)
            {
                answer_text = entry.attribute((W_NS, "val")).unwrap_or("").into();
            }
        }
        let label = clean_prompt(&question_text);
        let group_key = if label.is_empty() {
            field_name.clone()
        } else {
            label.clone()
        };
        records.push(Record {
            group_key,
            column_key: field_name.clone(),
            question_text: if label.is_empty() {
                field_name.clone()
            } else {
                question_text.trim().to_owned()
            },
            question_xml: if label.is_empty() {
                String::new()
            } else {
                question_xml
            },
            source_file: source_file.clone(),
            display_source: display_source.clone(),
            answer_text: if answer_text == "\u{2002}".repeat(5) {
                String::new()
            } else {
                answer_text
            },
            answer_xml,
            tables: Vec::new(),
        });
        previous_end = end.range().end;
    }
    Ok(records)
}

fn extract_content_controls(
    xml: &str,
    document: &Document<'_>,
    path: &Path,
    config: &Config,
) -> Result<Vec<Record>> {
    let controls: Vec<_> = document
        .descendants()
        .filter(|node| is_w(*node, "sdt"))
        .collect();
    let mut previous_end = 0usize;
    let mut records = Vec::new();
    let source_file = path.to_string_lossy().to_string();
    let display_source = path
        .file_name()
        .unwrap_or_default()
        .to_string_lossy()
        .into_owned();

    for control in controls {
        let properties = child(control, "sdtPr");
        let tag = properties.and_then(|p| child_value(p, "tag")).unwrap_or("");
        let title = properties
            .and_then(|p| child_value(p, "alias"))
            .unwrap_or("");
        let key = match config.extract.priority {
            Priority::Tag => {
                if tag.is_empty() {
                    title
                } else {
                    tag
                }
            }
            Priority::Title => {
                if title.is_empty() {
                    tag
                } else {
                    title
                }
            }
        };
        let key = if key.is_empty() {
            "未命名控件"
        } else {
            key
        }
        .to_owned();
        let (mut question_xml, question_text) =
            fragment_for_range(xml, document, previous_end, control.range().start);
        let Some(content) = child(control, "sdtContent") else {
            previous_end = previous_end.max(control.range().end);
            continue;
        };
        let answer_xml = content_fragment(xml, content);
        // Word's clipboard expands a selection adjoining a placeholder control.
        // Keep that observable VBA behavior without retaining an active control.
        if properties.is_some_and(|p| child(p, "showingPlcHdr").is_some())
            && !question_xml.is_empty()
        {
            if control.ancestors().any(|p| is_w(p, "p")) {
                if let Some(end) = question_xml.rfind("</w:p>") {
                    let placeholder = content.children().map(|n| raw(xml, n)).collect::<String>();
                    question_xml.insert_str(end, &placeholder);
                }
            } else {
                question_xml.push_str(&answer_xml);
            }
        }
        let text = display_text(content);
        let answer_text = text.strip_suffix('\r').unwrap_or(&text).to_owned();
        let tables = extract_tables(content);
        records.push(Record {
            group_key: if tag.is_empty() && title.is_empty() {
                "未命名题目".to_owned()
            } else {
                key.clone()
            },
            column_key: key.clone(),
            question_text: if question_text.trim().is_empty() {
                key
            } else {
                question_text.trim().to_owned()
            },
            question_xml: if has_visible_prompt(&question_text) {
                question_xml
            } else {
                String::new()
            },
            source_file: source_file.clone(),
            display_source: display_source.clone(),
            answer_text,
            answer_xml,
            tables,
        });
        previous_end = previous_end.max(control.range().end);
    }
    Ok(records)
}

fn fragment_for_range(
    xml: &str,
    document: &Document<'_>,
    start: usize,
    end: usize,
) -> (String, String) {
    if end <= start {
        return (String::new(), String::new());
    }
    let mut output = String::new();
    let mut visible = String::new();
    for paragraph in document
        .descendants()
        .filter(|node| is_w(*node, "p") || is_w(*node, "tbl"))
    {
        if paragraph.ancestors().skip(1).any(|ancestor| {
            is_w(ancestor, "tbl") && ancestor.range().start >= start && ancestor.range().end <= end
        }) {
            continue;
        }
        if is_w(paragraph, "tbl") {
            if paragraph.range().start >= start && paragraph.range().end <= end {
                output.push_str(raw(xml, paragraph));
                visible.push_str(&display_text(paragraph));
            }
            continue;
        }
        let runs: Vec<_> = paragraph
            .children()
            .filter(|node| {
                node.is_element()
                    && !is_w(*node, "pPr")
                    && node.range().start >= start
                    && node.range().end <= end
            })
            .collect();
        let includes_break = paragraph.range().end > start && paragraph.range().end <= end;
        if runs.is_empty() && !includes_break {
            continue;
        }
        output.push_str("<w:p>");
        if let Some(properties) = child(paragraph, "pPr") {
            output.push_str(raw(xml, properties));
        }
        for run in runs {
            output.push_str(raw(xml, run));
            visible.push_str(&display_text(run));
        }
        output.push_str("</w:p>");
        if includes_break {
            visible.push('\r');
        }
    }
    (output, visible)
}

fn content_fragment(xml: &str, content: Node<'_, '_>) -> String {
    let children: Vec<_> = content.children().collect();
    if children
        .iter()
        .any(|node| is_w(*node, "p") || is_w(*node, "tbl"))
    {
        children
            .into_iter()
            .map(|node| {
                if is_w(node, "sdt") {
                    content_fragment_from_control(xml, node)
                } else {
                    raw(xml, node).to_owned()
                }
            })
            .collect()
    } else {
        let parent_paragraph = content.ancestors().find(|node| is_w(*node, "p"));
        let mut paragraph = String::from("<w:p>");
        if let Some(ppr) = parent_paragraph.and_then(|p| child(p, "pPr")) {
            paragraph.push_str(raw(xml, ppr));
        }
        for node in children {
            if is_w(node, "sdt") {
                paragraph.push_str(&content_fragment_from_control(xml, node));
            } else {
                paragraph.push_str(raw(xml, node));
            }
        }
        paragraph.push_str("</w:p>");
        paragraph
    }
}

fn content_fragment_from_control(xml: &str, control: Node<'_, '_>) -> String {
    child(control, "sdtContent")
        .map(|content| content_fragment(xml, content))
        .unwrap_or_default()
}

fn display_text(node: Node<'_, '_>) -> String {
    if is_w(node, "drawing") {
        return if node
            .children()
            .any(|n| n.is_element() && n.tag_name().name() == "inline")
        {
            "\u{1}".into()
        } else {
            String::new()
        };
    }
    if is_w(node, "object") || is_w(node, "pict") {
        return "\u{1}".into();
    }
    if is_w(node, "t") || is_w(node, "delText") {
        return node.text().unwrap_or_default().to_owned();
    }
    if is_w(node, "tab") {
        return "\t".to_owned();
    }
    if is_w(node, "br") || is_w(node, "cr") {
        return match node.attribute((W_NS, "type")) {
            Some("page") => "\u{c}",
            Some("column") => "\u{e}",
            _ => "\u{b}",
        }
        .into();
    }
    let mut text = String::new();
    for child in node.children() {
        text.push_str(&display_text(child));
    }
    if is_w(node, "p") {
        text.push('\r');
    } else if is_w(node, "tr") {
        text.push_str("\r\u{7}");
    } else if is_w(node, "tc") {
        text.push('\u{7}');
    }
    text
}

fn extract_tables(content: Node<'_, '_>) -> Vec<TableData> {
    content
        .descendants()
        .filter(|node| is_w(*node, "tbl"))
        .enumerate()
        .map(|(index, table)| {
            let rows = table
                .children()
                .filter(|node| is_w(*node, "tr"))
                .map(|row| {
                    row.children()
                        .filter(|node| is_w(*node, "tc"))
                        .map(|cell| {
                            display_text(cell)
                                .trim_matches(['\n', '\r', '\u{7}', '\t', ' '])
                                .to_owned()
                        })
                        .collect()
                })
                .collect();
            TableData {
                index: index + 1,
                rows,
            }
        })
        .collect()
}

fn ancestor_field_char<'a, 'input>(node: Node<'a, 'input>, kind: &str) -> Option<Node<'a, 'input>> {
    node.ancestors()
        .find(|ancestor| is_w(*ancestor, "fldChar") && field_char_type(*ancestor) == Some(kind))
}

fn field_char_type<'a, 'input>(node: Node<'a, 'input>) -> Option<&'a str> {
    node.attribute((W_NS, "fldCharType"))
}

fn child<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<Node<'a, 'input>> {
    node.children().find(|child| is_w(*child, name))
}

fn child_value<'a, 'input>(node: Node<'a, 'input>, name: &str) -> Option<&'a str> {
    child(node, name)?.attribute((W_NS, "val"))
}

fn is_w(node: Node<'_, '_>, name: &str) -> bool {
    node.is_element() && node.tag_name().namespace() == Some(W_NS) && node.tag_name().name() == name
}

fn raw<'a>(xml: &'a str, node: Node<'_, 'a>) -> &'a str {
    &xml[node.range()]
}

fn clean_prompt(text: &str) -> String {
    text.replace('\r', "")
        .replace('\n', "")
        .replace(':', "")
        .replace('：', "")
        .trim_matches(' ')
        .to_owned()
}

fn has_visible_prompt(text: &str) -> bool {
    !text
        .replace(['\r', '\n', '\t', '\u{7}'], "")
        .replace(['\u{a0}', '\u{3000}'], " ")
        .trim_matches(' ')
        .is_empty()
}

fn mapped_name(path: &Path, config: &Config) -> String {
    let filename = path
        .file_name()
        .and_then(|name| name.to_str())
        .unwrap_or_default();
    for (key, value) in &config.workbook_mapping {
        if filename.to_lowercase().contains(&key.to_lowercase()) {
            return value.clone();
        }
    }
    let Some(mapping_path) = config.security.mapping_csv.as_ref() else {
        return filename.to_owned();
    };
    let mapping_path = PathBuf::from(mapping_path);
    let Ok(text) = fs::read_to_string(mapping_path) else {
        return filename.to_owned();
    };
    for line in text
        .lines()
        .filter(|line| !line.trim().is_empty() && !line.trim_start().starts_with('#'))
    {
        if let Some((needle, replacement)) = line.split_once(',') {
            let needle = needle.trim().trim_matches('"');
            let replacement = replacement.trim().trim_matches('"');
            if !needle.is_empty() && filename.to_lowercase().contains(&needle.to_lowercase()) {
                return replacement.to_owned();
            }
        }
    }
    filename.to_owned()
}

pub fn write_excel(
    report: &Report,
    output: &Path,
    form_kind: FormKind,
    config: &Config,
) -> Result<()> {
    ensure!(
        output
            .extension()
            .is_some_and(|ext| ext.eq_ignore_ascii_case("xlsx")),
        "Rust Excel output must use .xlsx"
    );
    let mut workbook = Workbook::new();
    let worksheet_name = if form_kind == FormKind::Legacy {
        "output"
    } else {
        "output_cc"
    };
    let worksheet = workbook.add_worksheet();
    worksheet.set_name(worksheet_name)?;
    let header_format = Format::new().set_bold();

    match config.extract.layout {
        OutputLayout::Wide => write_wide_sheet(worksheet, report, form_kind, &header_format)?,
        OutputLayout::Long => write_long_sheet(worksheet, report, &header_format)?,
    }
    if config.extract.table_mode == TableMode::SeparateSheet
        && report.records.iter().any(|r| !r.tables.is_empty())
    {
        let tables = workbook.add_worksheet();
        tables.set_name("Tables")?;
        write_tables_sheet(tables, report, &header_format)?;
    }
    workbook
        .save(output)
        .with_context(|| format!("Could not save {}", output.display()))
}

fn write_wide_sheet(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    report: &Report,
    form_kind: FormKind,
    header: &Format,
) -> Result<()> {
    if form_kind == FormKind::Legacy {
        worksheet.write_string_with_format(0, 0, "源文件名 (映射后)", header)?;
        for (row, source) in report.sources.iter().enumerate() {
            worksheet.write_string(row as u32 + 1, 0, &source.display)?;
            for (col, record) in report
                .records
                .iter()
                .filter(|r| r.source_file == source.file)
                .enumerate()
            {
                if row == 0 {
                    worksheet.write_string_with_format(
                        0,
                        col as u16 + 1,
                        &record.column_key,
                        header,
                    )?;
                }
                write_vba_value(
                    worksheet,
                    row as u32 + 1,
                    col as u16 + 1,
                    &record.answer_text,
                )?;
            }
        }
        return Ok(());
    }
    let mut columns = Vec::new();
    let mut seen = HashSet::new();
    for record in &report.records {
        if seen.insert(record.column_key.clone()) {
            columns.push(record.column_key.clone());
        }
    }
    let mut sources: Vec<String> = report.sources.iter().map(|s| s.display.clone()).collect();
    let mut row_for_source: HashMap<String, usize> = report
        .sources
        .iter()
        .enumerate()
        .map(|(i, s)| (s.file.clone(), i))
        .collect();
    let mut cells: HashMap<(usize, String), String> = HashMap::new();
    for record in &report.records {
        let row = *row_for_source
            .entry(record.source_file.clone())
            .or_insert_with(|| {
                sources.push(record.display_source.clone());
                sources.len() - 1
            });
        cells.insert((row, record.column_key.clone()), record.answer_text.clone());
    }
    worksheet.write_string_with_format(0, 0, "源文件名", header)?;
    for (index, column) in columns.iter().enumerate() {
        worksheet.write_string_with_format(0, (index + 1) as u16, column, header)?;
    }
    for (row_index, source) in sources.iter().enumerate() {
        let row = (row_index + 1) as u32;
        worksheet.write_string(row, 0, source)?;
        for (column_index, column) in columns.iter().enumerate() {
            if let Some(value) = cells.get(&(row_index, column.clone())) {
                write_vba_value(worksheet, row, (column_index + 1) as u16, value)?;
            }
        }
    }
    worksheet.set_column_width(0, 28)?;
    for col in 1..=columns.len() as u16 {
        worksheet.set_column_width(col, 24)?;
    }
    Ok(())
}

fn write_vba_value(
    sheet: &mut rust_xlsxwriter::Worksheet,
    row: u32,
    col: u16,
    text: &str,
) -> Result<()> {
    if text.is_empty() {
        return Ok(());
    }
    if let Ok(number) = text.trim().parse::<f64>() {
        if number.is_finite() {
            sheet.write_number(row, col, number)?;
            return Ok(());
        }
    }
    if text.starts_with('=') {
        sheet.write_formula(row, col, text)?;
    } else {
        sheet.write_string(row, col, text)?;
    }
    Ok(())
}

fn write_long_sheet(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    report: &Report,
    header: &Format,
) -> Result<()> {
    for (col, value) in ["源文件名", "题目键", "题目", "答案"].iter().enumerate() {
        worksheet.write_string_with_format(0, col as u16, *value, header)?;
    }
    for (index, record) in report.records.iter().enumerate() {
        let row = index as u32 + 1;
        worksheet.write_string(row, 0, &record.display_source)?;
        worksheet.write_string(row, 1, &record.column_key)?;
        worksheet.write_string(row, 2, &record.question_text)?;
        worksheet.write_string(row, 3, &record.answer_text)?;
    }
    worksheet.set_column_width(0, 28)?;
    worksheet.set_column_width(1, 28)?;
    worksheet.set_column_width(2, 40)?;
    worksheet.set_column_width(3, 60)?;
    Ok(())
}

fn write_tables_sheet(
    worksheet: &mut rust_xlsxwriter::Worksheet,
    report: &Report,
    header: &Format,
) -> Result<()> {
    for (col, value) in ["源文件名", "题目键", "表格序号", "行", "列", "单元格内容"]
        .iter()
        .enumerate()
    {
        worksheet.write_string_with_format(0, col as u16, *value, header)?;
    }
    let mut row = 1u32;
    for record in &report.records {
        for table in &record.tables {
            for (row_index, cells) in table.rows.iter().enumerate() {
                for (column_index, value) in cells.iter().enumerate() {
                    worksheet.write_string(row, 0, &record.display_source)?;
                    worksheet.write_string(row, 1, &record.column_key)?;
                    worksheet.write_number(row, 2, table.index as f64)?;
                    worksheet.write_number(row, 3, (row_index + 1) as f64)?;
                    worksheet.write_number(row, 4, (column_index + 1) as f64)?;
                    worksheet.write_string(row, 5, value)?;
                    row += 1;
                }
            }
        }
    }
    worksheet.set_column_width(0, 28)?;
    worksheet.set_column_width(1, 28)?;
    worksheet.set_column_width(5, 44)?;
    Ok(())
}

pub fn write_word(report: &Report, output: &Path) -> Result<()> {
    ensure!(
        output
            .extension()
            .is_some_and(|ext| ext.eq_ignore_ascii_case("docx")),
        "Rust Word output must use .docx"
    );
    ensure!(
        !report.sources.is_empty(),
        "Cannot create a Word summary without input documents"
    );

    let mut prepared = crate::fidelity::Prepared::new(report)?;
    let body = prepared.sanitize_body(&make_word_body(&prepared.report))?;
    write_minimal_docx(
        output,
        &body,
        &prepared.report.namespaces,
        &prepared.ignorable_prefixes(),
    )?;
    prepared.finish(output)
}

fn make_word_body(report: &Report) -> String {
    if report.form_kind == Some(FormKind::ContentControl) {
        let mut body = String::new();
        let mut seen = HashSet::new();
        for record in &report.records {
            if seen.insert(record.group_key.to_lowercase()) {
                if record.question_xml.is_empty() {
                    body.push_str(&plain_paragraph(&record.group_key, true));
                } else {
                    body.push_str(&trim_prompt_space(&record.question_xml));
                }
            }
            body.push_str(&plain_paragraph(&record.display_source, false));
            if record.answer_xml.is_empty() {
                body.push_str(&plain_paragraph(&record.answer_text, false));
            } else {
                body.push_str(&record.answer_xml);
            }
        }
        body.push_str("<w:p/>");
        return body;
    }
    let mut groups: Vec<(String, &Record, Vec<&Record>)> = Vec::new();
    let mut indices: HashMap<String, usize> = HashMap::new();
    for record in &report.records {
        if let Some(index) = indices.get(&record.group_key).copied() {
            groups[index].2.push(record);
        } else {
            indices.insert(record.group_key.clone(), groups.len());
            groups.push((record.group_key.clone(), record, vec![record]));
        }
    }

    let mut body = String::new();
    for (key, heading, records) in groups {
        if heading.question_xml.is_empty() {
            body.push_str(&plain_paragraph(&heading.question_text, true));
        } else {
            body.push_str(&heading.question_xml);
            if heading.question_xml.trim_end().ends_with("</w:tbl>") {
                body.push_str("<w:p/>");
            }
        }
        if body.is_empty() {
            body.push_str(&plain_paragraph(&key, true));
        }
        for record in records {
            body.push_str(&inline_answer(record));
        }
        body.push_str("<w:p/>");
    }
    body.push_str("<w:p/>");
    body
}

fn trim_prompt_space(xml: &str) -> String {
    if let Some(end) = xml.rfind("</w:t>") {
        return format!("{}{}", xml[..end].trim_end_matches(' '), &xml[end..]);
    }
    xml.to_owned()
}

fn inline_answer(record: &Record) -> String {
    let prefix = format!(
        "<w:r><w:t xml:space=\"preserve\">【{}】: </w:t></w:r>",
        escape_xml(&record.display_source)
    );
    let suffix = "<w:r><w:t xml:space=\"preserve\">; </w:t></w:r>";
    let fragment = if record.answer_xml.is_empty() {
        plain_paragraph(&record.answer_text, false)
    } else {
        record.answer_xml.clone()
    };
    let Some(open) = fragment.find("<w:p") else {
        return plain_paragraph(
            &format!("【{}】: {}; ", record.display_source, record.answer_text),
            false,
        );
    };
    let mut insert = open + fragment[open..].find('>').unwrap_or(4) + 1;
    if fragment[insert..].starts_with("<w:pPr") {
        if let Some(end) = fragment[insert..].find("</w:pPr>") {
            insert += end + "</w:pPr>".len();
        } else if let Some(end) = fragment[insert..].find("/>") {
            insert += end + 2;
        }
    }
    let close = fragment.rfind("</w:p>").unwrap_or(fragment.len());
    format!(
        "{}{}{}{}{}",
        &fragment[..insert],
        prefix,
        &fragment[insert..close],
        suffix,
        &fragment[close..]
    )
}

fn plain_paragraph(text: &str, bold: bool) -> String {
    let bold = if bold { "<w:b/>" } else { "" };
    format!(
        "<w:p><w:r><w:rPr>{bold}</w:rPr><w:t xml:space=\"preserve\">{}</w:t></w:r></w:p>",
        escape_xml(text)
    )
}

fn escape_xml(text: &str) -> String {
    text.replace('&', "&amp;")
        .replace('<', "&lt;")
        .replace('>', "&gt;")
}

pub(crate) fn write_minimal_docx(
    output: &Path,
    body: &str,
    namespaces: &BTreeMap<String, String>,
    ignorable: &str,
) -> Result<()> {
    let mut namespace_declarations = format!(" xmlns:w=\"{W_NS}\" xmlns:r=\"{R_NS}\"");
    for (prefix, uri) in namespaces {
        if prefix.is_empty() || prefix == "xml" || prefix == "w" || prefix == "r" {
            continue;
        }
        namespace_declarations.push_str(&format!(" xmlns:{prefix}=\"{}\"", escape_xml(uri)));
    }
    let compatibility = if ignorable.is_empty() {
        String::new()
    } else {
        format!(" mc:Ignorable=\"{}\"", escape_xml(ignorable))
    };
    let document_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:document{namespace_declarations}{compatibility}><w:body>{body}<w:sectPr><w:pgSz w:w=\"11906\" w:h=\"16838\"/><w:pgMar w:top=\"1440\" w:right=\"1440\" w:bottom=\"1440\" w:left=\"1440\" w:header=\"720\" w:footer=\"720\" w:gutter=\"0\"/></w:sectPr></w:body></w:document>"
    );
    let content_types = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/><Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/><Override PartName="/docProps/core.xml" ContentType="application/vnd.openxmlformats-package.core-properties+xml"/><Override PartName="/docProps/app.xml" ContentType="application/vnd.openxmlformats-officedocument.extended-properties+xml"/></Types>"#;
    let root_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/><Relationship Id="rId2" Type="http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Target="docProps/core.xml"/><Relationship Id="rId3" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Target="docProps/app.xml"/></Relationships>"#;
    let doc_rels = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>"#;
    let styles_xml = format!(
        "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?><w:styles xmlns:w=\"{W_NS}\"><w:docDefaults><w:rPrDefault><w:rPr><w:rFonts w:ascii=\"Arial\" w:hAnsi=\"Arial\" w:eastAsia=\"Microsoft YaHei\"/><w:sz w:val=\"22\"/></w:rPr></w:rPrDefault><w:pPrDefault><w:pPr/></w:pPrDefault></w:docDefaults><w:style w:type=\"paragraph\" w:default=\"1\" w:styleId=\"Normal\"><w:name w:val=\"Normal\"/><w:qFormat/></w:style></w:styles>"
    );
    let core = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties" xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"><dc:title>vbaWord summary</dc:title></cp:coreProperties>"#;
    let app = r#"<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"><Application>vbaWord CLI</Application></Properties>"#;

    let file =
        File::create(output).with_context(|| format!("Could not create {}", output.display()))?;
    let mut zip = ZipWriter::new(file);
    let options = SimpleFileOptions::default().compression_method(zip::CompressionMethod::Deflated);
    for (name, contents) in [
        ("[Content_Types].xml", content_types),
        ("_rels/.rels", root_rels),
        ("word/_rels/document.xml.rels", doc_rels),
        ("word/document.xml", &document_xml),
        ("word/styles.xml", &styles_xml),
        ("docProps/core.xml", core),
        ("docProps/app.xml", app),
    ] {
        zip.start_file(name, options)?;
        use std::io::Write;
        zip.write_all(contents.as_bytes())?;
    }
    zip.finish()?;
    Ok(())
}
