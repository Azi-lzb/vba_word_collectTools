use std::collections::BTreeMap;

use serde::{Deserialize, Serialize};

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
pub enum FormKind {
    Legacy,
    ContentControl,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
pub enum OutputKind {
    Excel,
    Word,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum Priority {
    Tag,
    Title,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "lowercase")]
pub enum OutputLayout {
    Wide,
    Long,
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
pub enum TableMode {
    Flatten,
    SeparateSheet,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct Report {
    #[serde(default)]
    pub sources: Vec<Source>,
    pub form_kind: Option<FormKind>,
    pub records: Vec<Record>,
    pub template_path: Option<String>,
    pub namespaces: BTreeMap<String, String>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct Source {
    pub file: String,
    pub display: String,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct Record {
    /// Key used to group Word output sections.
    pub group_key: String,
    /// Key used to build Excel columns.
    pub column_key: String,
    /// Visible question text, when available in the source document.
    pub question_text: String,
    /// WordprocessingML question fragment. Contains paragraph/run formatting.
    pub question_xml: String,
    pub source_file: String,
    pub display_source: String,
    /// Plain text used for Excel cells and fallback output.
    pub answer_text: String,
    /// WordprocessingML block fragment. Paragraphs and tables remain structured.
    pub answer_xml: String,
    pub tables: Vec<TableData>,
}

#[derive(Clone, Debug, Default, Deserialize, Serialize)]
pub struct TableData {
    pub index: usize,
    pub rows: Vec<Vec<String>>,
}
