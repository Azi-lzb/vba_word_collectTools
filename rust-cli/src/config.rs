use std::{
    fs,
    path::{Path, PathBuf},
};

use anyhow::{Context, Result};
use calamine::{Reader, open_workbook_auto};
use serde::{Deserialize, Serialize};

use crate::model::{OutputLayout, Priority, TableMode};

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(default)]
pub struct Config {
    pub backend: BackendConfig,
    pub extract: ExtractConfig,
    pub security: SecurityConfig,
    #[serde(skip)]
    pub workbook_mapping: Vec<(String, String)>,
    #[serde(skip)]
    pub workbook_password: String,
}

impl Default for Config {
    fn default() -> Self {
        Self {
            backend: BackendConfig::default(),
            extract: ExtractConfig::default(),
            security: SecurityConfig::default(),
            workbook_mapping: Vec::new(),
            workbook_password: String::new(),
        }
    }
}

impl Config {
    pub fn load_optional(path: &Path) -> Result<Self> {
        let mut config: Self = if path.exists() {
            let text = fs::read_to_string(path)
                .with_context(|| format!("Could not read config {}", path.display()))?;
            toml::from_str(&text)
                .with_context(|| format!("Could not parse config {}", path.display()))?
        } else {
            Self::default()
        };
        let workbook = find_config().or_else(find_workbook);
        if let Some(workbook) = workbook {
            let standalone = workbook
                .file_name()
                .is_some_and(|name| name.eq_ignore_ascii_case("config.xlsx"));
            let mut book = open_workbook_auto(&workbook)
                .with_context(|| format!("无法读取配置工作簿 {}", workbook.display()))?;
            let sheet = book
                .worksheet_range("mapping")
                .context("配置工作簿缺少 mapping 表")?;
            let cell = |row, col| {
                sheet
                    .get_value((row, col))
                    .map(|v| v.to_string())
                    .unwrap_or_default()
            };
            config.workbook_password = cell(0, 4);
            let mut priority = cell(0, 5);
            if let Ok(settings) = book.worksheet_range("config") {
                if standalone || config.workbook_password.is_empty() {
                    config.workbook_password = settings
                        .get_value((1, 1))
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                }
                if standalone || priority.is_empty() {
                    priority = settings
                        .get_value((2, 1))
                        .map(|v| v.to_string())
                        .unwrap_or_default();
                }
            }
            if standalone || !path.exists() {
                config.extract.priority = if priority.eq_ignore_ascii_case("TITLE") {
                    Priority::Title
                } else {
                    Priority::Tag
                };
            }
            for row in sheet.rows().skip(1) {
                let key = row.first().map(|v| v.to_string()).unwrap_or_default();
                let value = row.get(1).map(|v| v.to_string()).unwrap_or_default();
                if !key.trim().is_empty() {
                    config.workbook_mapping.push((key, value));
                }
            }
        }
        if let Some(mapping) = config.security.mapping_csv.as_mut() {
            let mapping_path = Path::new(mapping);
            if mapping_path.is_relative() {
                let base = path
                    .parent()
                    .filter(|parent| !parent.as_os_str().is_empty())
                    .unwrap_or(Path::new("."));
                *mapping = base.join(mapping_path).to_string_lossy().to_string();
            }
        }
        Ok(config)
    }
}

pub(crate) fn find_workbook() -> Option<PathBuf> {
    let current = std::env::current_dir().ok()?.join("vbaForWordV3.xls");
    if current.is_file() {
        return Some(current);
    }
    std::env::current_exe().ok().and_then(|exe| {
        exe.ancestors()
            .skip(1)
            .map(|p| p.join("vbaForWordV3.xls"))
            .find(|p| p.is_file())
    })
}

pub(crate) fn find_config() -> Option<PathBuf> {
    if let Some(path) = std::env::current_exe()
        .ok()
        .and_then(|p| p.parent().map(|p| p.join("config.xlsx")))
        .filter(|p| p.is_file())
    {
        return Some(path);
    }
    let current = std::env::current_dir().ok()?.join("config.xlsx");
    if current.is_file() {
        return Some(current);
    }
    std::env::current_exe().ok().and_then(|exe| {
        exe.ancestors()
            .skip(1)
            .map(|p| p.join("config.xlsx"))
            .find(|p| p.is_file())
    })
}

pub(crate) fn ensure_config() -> Result<PathBuf> {
    if let Some(path) = find_config() {
        return Ok(path);
    }
    let path = find_workbook()
        .and_then(|p| p.parent().map(|p| p.join("config.xlsx")))
        .or_else(|| {
            std::env::current_exe()
                .ok()
                .and_then(|p| p.parent().map(|p| p.join("config.xlsx")))
        })
        .context("无法定位配置目录")?;
    #[cfg(target_os = "linux")]
    let path = if path.parent().is_some_and(|p| {
        ["/usr/bin", "/usr/local/bin", "/bin", "/sbin", "/usr/sbin"]
            .iter()
            .any(|s| p == Path::new(s))
    }) {
        std::env::current_dir()?.join("config.xlsx")
    } else {
        path
    };
    create_workbook_config(&path)?;
    Ok(path)
}

pub fn create_workbook_config(path: &Path) -> Result<()> {
    use rust_xlsxwriter::{Format, Workbook};
    anyhow::ensure!(!path.exists(), "配置已存在，不覆盖：{}", path.display());
    let mut source = find_workbook()
        .map(|p| open_workbook_auto(p))
        .transpose()
        .context("无法读取原 VBA 配置")?;
    let mut output = Workbook::new();
    let header = Format::new().set_bold();
    for name in ["config", "mapping"] {
        let sheet = output.add_worksheet();
        sheet.set_name(name)?;
        if let Some(range) = source
            .as_mut()
            .and_then(|book| book.worksheet_range(name).ok())
        {
            if let Some((row_start, col_start)) = range.start() {
                for (r, row) in range.rows().enumerate() {
                    for (c, cell) in row.iter().enumerate() {
                        let value = cell.to_string();
                        if !value.is_empty() {
                            sheet.write_string(
                                row_start + r as u32,
                                col_start as u16 + c as u16,
                                &value,
                            )?;
                        }
                    }
                }
            }
        } else if name == "config" {
            for (r, values) in [
                ["配置项", "值", "说明"],
                ["文档保护密码", "", "受保护 Word 文档的密码，未设置时留空"],
                ["题目识别优先级", "TAG", "TAG 或 TITLE"],
            ]
            .iter()
            .enumerate()
            {
                for (c, value) in values.iter().enumerate() {
                    sheet.write_string(r as u32, c as u16, *value)?;
                }
            }
        } else {
            sheet.write_string(0, 0, "源文件名关键词")?;
            sheet.write_string(0, 1, "映射名称")?;
        }
        sheet.set_row_format(0, &header)?;
        sheet.set_column_width(0, 26)?;
        sheet.set_column_width(1, 30)?;
        sheet.set_column_width(2, 65)?;
        sheet.set_freeze_panes(1, 0)?;
    }
    let mut file = std::fs::OpenOptions::new()
        .write(true)
        .create_new(true)
        .open(path)
        .with_context(|| format!("无法创建配置 {}", path.display()))?;
    output.save_to_writer(&mut file)?;
    Ok(())
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(default)]
pub struct BackendConfig {
    pub auto_excel: ExcelPreference,
    pub python_executable: String,
}

impl Default for BackendConfig {
    fn default() -> Self {
        Self {
            auto_excel: ExcelPreference::Rust,
            python_executable: "python".to_owned(),
        }
    }
}

#[derive(Clone, Copy, Debug, Deserialize, Serialize, PartialEq, Eq)]
#[serde(rename_all = "kebab-case")]
pub enum ExcelPreference {
    Rust,
    Python,
    Com,
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(default)]
pub struct ExtractConfig {
    pub priority: Priority,
    pub layout: OutputLayout,
    pub table_mode: TableMode,
    pub preserve_format: bool,
}

impl Default for ExtractConfig {
    fn default() -> Self {
        Self {
            priority: Priority::Tag,
            layout: OutputLayout::Wide,
            table_mode: TableMode::SeparateSheet,
            preserve_format: true,
        }
    }
}

#[derive(Clone, Debug, Deserialize, Serialize)]
#[serde(default)]
pub struct SecurityConfig {
    pub password_env: Option<String>,
    pub mapping_csv: Option<String>,
}

impl Default for SecurityConfig {
    fn default() -> Self {
        Self {
            password_env: Some("VBAWORD_DOC_PASSWORD".to_owned()),
            mapping_csv: None,
        }
    }
}

pub fn write_example(path: &Path) -> Result<()> {
    if path.exists() {
        anyhow::bail!("Refusing to overwrite existing config {}", path.display());
    }
    let text = r#"[backend]
# 自动路由固定优先 Rust；下列解释器仅供 openpyxl/COM 回退使用。
python_executable = "python"

[extract]
priority = "tag"                 # 内容控件：tag 或 title
layout = "wide"                   # wide 或 long
table_mode = "separate-sheet"    # separate-sheet 或 flatten
preserve_format = true

[security]
# 密码从该环境变量读取，不明文保存于配置文件。
password_env = "VBAWORD_DOC_PASSWORD"
mapping_csv = "mapping.csv"      # 可选：文件名关键词,映射名称
"#;
    fs::write(path, text).with_context(|| format!("Could not write config {}", path.display()))
}
