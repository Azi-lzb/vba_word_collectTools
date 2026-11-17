use std::{
    io::Write,
    path::{Path, PathBuf},
    process::{Command, Stdio},
};

use anyhow::{Context, Result, bail, ensure};
use serde::{Deserialize, Serialize};

use crate::{
    config::Config,
    model::{FormKind, OutputKind, Report},
};

const WORKER: &str = include_str!("../python/worker.py");

#[derive(Debug, Deserialize)]
pub struct ProbeInfo {
    pub executable: String,
    pub openpyxl: bool,
    pub pywin32: bool,
    pub word_com: bool,
    pub excel_com: bool,
}

#[derive(Serialize)]
struct ComJob<'a> {
    form_kind: FormKind,
    output_kind: OutputKind,
    inputs: Vec<String>,
    output: String,
    priority: crate::model::Priority,
    layout: crate::model::OutputLayout,
    table_mode: crate::model::TableMode,
    password: String,
    mapping_csv: Option<&'a str>,
    workbook_mapping: &'a [(String, String)],
}

#[derive(Serialize)]
struct ExcelJob<'a> {
    report: &'a Report,
    form_kind: FormKind,
    layout: crate::model::OutputLayout,
    table_mode: crate::model::TableMode,
    output: String,
}

pub fn probe() -> Result<ProbeInfo> {
    let output = run_worker("probe", &[])?;
    serde_json::from_slice(&output).context("Could not read Python backend probe result")
}

pub fn probe_openpyxl(python: &str) -> Result<bool> {
    probe_available(python, "probe-openpyxl")
}

pub fn probe_word_com(python: &str) -> Result<bool> {
    probe_available(python, "probe-word-com")
}

pub fn probe_excel_com(python: &str) -> Result<bool> {
    probe_available(python, "probe-excel-com")
}

fn probe_available(python: &str, operation: &str) -> Result<bool> {
    let Ok(output) = run_worker_with_python(python, operation, &[]) else {
        return Ok(false);
    };
    Ok(serde_json::from_slice::<serde_json::Value>(&output)
        .ok()
        .is_some_and(|value| value["available"] == true))
}

pub fn run_com(
    inputs: &[PathBuf],
    output: &Path,
    form_kind: FormKind,
    output_kind: OutputKind,
    config: &Config,
) -> Result<()> {
    let password = config
        .security
        .password_env
        .as_deref()
        .and_then(std::env::var_os)
        .map(|value| value.to_string_lossy().to_string())
        .unwrap_or_else(|| config.workbook_password.clone());
    let job = ComJob {
        form_kind,
        output_kind,
        inputs: inputs
            .iter()
            .map(|path| path.to_string_lossy().to_string())
            .collect(),
        output: output.to_string_lossy().to_string(),
        priority: config.extract.priority,
        layout: config.extract.layout,
        table_mode: config.extract.table_mode,
        password,
        mapping_csv: config.security.mapping_csv.as_deref(),
        workbook_mapping: &config.workbook_mapping,
    };
    let payload = serde_json::to_vec(&job)?;
    let _ = run_worker_with_python(&config.backend.python_executable, "com-extract", &payload)?;
    Ok(())
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
        "Python/openpyxl output must use .xlsx"
    );
    let job = ExcelJob {
        report,
        form_kind,
        layout: config.extract.layout,
        table_mode: config.extract.table_mode,
        output: output.to_string_lossy().to_string(),
    };
    let payload = serde_json::to_vec(&job)?;
    let _ = run_worker_with_python(&config.backend.python_executable, "write-excel", &payload)?;
    Ok(())
}

fn run_worker(operation: &str, input: &[u8]) -> Result<Vec<u8>> {
    run_worker_with_python("python", operation, input)
}

fn run_worker_with_python(python: &str, operation: &str, input: &[u8]) -> Result<Vec<u8>> {
    let mut script = tempfile::Builder::new()
        .prefix("vbaword-worker-")
        .suffix(".py")
        .tempfile()
        .context("Could not stage the Python adapter")?;
    script.write_all(WORKER.as_bytes())?;
    script.flush()?;

    let mut child = Command::new(python)
        .arg(script.path())
        .arg(operation)
        .stdin(Stdio::piped())
        .stdout(Stdio::piped())
        .stderr(Stdio::piped())
        .spawn()
        .with_context(|| format!("Could not start Python interpreter '{python}'"))?;
    if !input.is_empty() {
        child
            .stdin
            .take()
            .context("Python worker stdin was unavailable")?
            .write_all(input)?;
    }
    let output = child.wait_with_output()?;
    if !output.status.success() {
        let message = String::from_utf8_lossy(&output.stderr);
        bail!("Python adapter '{operation}' failed: {}", message.trim());
    }
    Ok(output.stdout)
}
