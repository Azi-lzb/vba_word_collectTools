mod config;
mod fidelity;
mod hyperlinks;
mod menu;
mod model;
mod native;
mod opc;
mod python;
mod samples;
mod xml;

use std::fs;
use std::path::{Path, PathBuf};

use anyhow::{Context, Result, bail, ensure};
use clap::{Args, Parser, Subcommand, ValueEnum};
use config::Config;
use model::{FormKind, OutputKind};

#[derive(Parser)]
#[command(
    name = "vbaword",
    version,
    about = "Extract Word questionnaire forms to Excel or Word"
)]
struct Cli {
    #[command(subcommand)]
    command: Command,
}

#[derive(Subcommand)]
enum Command {
    Extract(ExtractArgs),
    InitConfig {
        #[arg(short, long, default_value = "vbaword.toml")]
        output: PathBuf,
    },
    Probe,
    InitWorkbookConfig {
        #[arg(long, default_value = "config.xlsx")]
        output: PathBuf,
        #[arg(long)]
        source: Option<PathBuf>,
        #[arg(long)]
        clear_password: bool,
    },
    /// Generate example Word forms without Office.
    GenerateExamples {
        #[arg(long, value_enum, default_value = "both")]
        kind: samples::SampleKind,
        #[arg(long)]
        output: PathBuf,
    },
}

#[derive(Args)]
struct ExtractArgs {
    #[arg(long, value_enum)]
    kind: FormKindArg,
    #[arg(long, value_enum)]
    to: OutputKindArg,
    #[arg(long, value_enum, default_value_t = BackendArg::Auto)]
    backend: BackendArg,
    #[arg(long, required = true, num_args = 1..)]
    input: Vec<PathBuf>,
    #[arg(long)]
    output: PathBuf,
    #[arg(long)]
    config: Option<PathBuf>,
    #[arg(long, value_enum)]
    priority: Option<PriorityArg>,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
enum FormKindArg {
    Legacy,
    ContentControl,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
enum OutputKindArg {
    Excel,
    Word,
}

#[derive(Clone, Copy, Debug, ValueEnum, PartialEq, Eq)]
enum BackendArg {
    Auto,
    Com,
    Python,
    Rust,
}

#[derive(Clone, Copy, Debug, ValueEnum)]
enum PriorityArg {
    Tag,
    Title,
}

fn main() -> Result<()> {
    if std::env::args_os().len() == 1 {
        return menu::run();
    }
    match Cli::parse().command {
        Command::InitWorkbookConfig {
            output,
            source,
            clear_password,
        } => config::create_workbook_config_from(&output, source.as_deref(), clear_password),
        Command::GenerateExamples { kind, output } => {
            samples::generate(kind, &output).map(|files| {
                for path in files {
                    println!("{}", path.display());
                }
            })
        }
        Command::InitConfig { output } => {
            config::write_example(&output)?;
            println!("Created {}", output.display());
            Ok(())
        }
        Command::Probe => {
            println!("Rust OOXML: true");
            match python::probe() {
                Ok(info) => {
                    println!("Python: {}", info.executable);
                    println!("openpyxl: {}", info.openpyxl);
                    println!("pywin32: {}", info.pywin32);
                    println!("Word COM: {}", info.word_com);
                    println!("Excel COM: {}", info.excel_com);
                }
                Err(error) => println!("Optional Python/COM backends unavailable: {error:#}"),
            }
            Ok(())
        }
        Command::Extract(args) => run_extract(args),
    }
}

fn run_extract(args: ExtractArgs) -> Result<()> {
    let config_path = args
        .config
        .clone()
        .unwrap_or_else(|| PathBuf::from("vbaword.toml"));
    let mut config = Config::load_optional(&config_path)?;
    if let Some(priority) = args.priority {
        config.extract.priority = match priority {
            PriorityArg::Tag => model::Priority::Tag,
            PriorityArg::Title => model::Priority::Title,
        };
    }

    let inputs = expand_inputs(&args.input)?;
    if args.output.exists() {
        let output = fs::canonicalize(&args.output)?;
        ensure!(
            !inputs
                .iter()
                .any(|input| fs::canonicalize(input).is_ok_and(|path| path == output)),
            "Output must not overwrite an input document"
        );
    }

    let form_kind = match args.kind {
        FormKindArg::Legacy => FormKind::Legacy,
        FormKindArg::ContentControl => FormKind::ContentControl,
    };
    let output_kind = match args.to {
        OutputKindArg::Excel => OutputKind::Excel,
        OutputKindArg::Word => OutputKind::Word,
    };
    if args.backend == BackendArg::Auto {
        return run_auto_extract(&inputs, &args.output, form_kind, output_kind, &config);
    }
    let backend = route_backend(args.backend, output_kind, &inputs, &args.output, &config)?;
    println!("Backend: {backend:?}");
    if backend == BackendArg::Com {
        return python::run_com(&inputs, &args.output, form_kind, output_kind, &config);
    }
    if backend == BackendArg::Python && output_kind != OutputKind::Excel {
        bail!(
            "The Python backend writes Excel output only. Use --backend com or --backend rust for Word output."
        );
    }

    let report = native::extract(&inputs, form_kind, &config).with_context(|| {
        format!("Could not extract {form_kind:?} fields with the Rust OOXML reader")
    })?;
    match (backend, output_kind) {
        (BackendArg::Rust, OutputKind::Excel) => {
            native::write_excel(&report, &args.output, form_kind, &config)
        }
        (BackendArg::Rust, OutputKind::Word) => native::write_word(&report, &args.output),
        (BackendArg::Python, OutputKind::Excel) => {
            python::write_excel(&report, &args.output, form_kind, &config)
        }
        _ => bail!("Unsupported backend route"),
    }
}

fn run_auto_extract(
    inputs: &[PathBuf],
    output: &Path,
    kind: FormKind,
    target: OutputKind,
    config: &Config,
) -> Result<()> {
    validate_output_extension(output, target)?;
    if inputs.iter().any(|path| extension_is(path, "doc"))
        || !["docx", "xlsx"].iter().any(|ext| extension_is(output, ext))
    {
        println!("Backend: Com (legacy format)");
        return python::run_com(inputs, output, kind, target, config)
            .context("旧式输入或输出格式需要 Word/Excel COM");
    }
    println!("Backend: Rust");
    let report = match native::extract(inputs, kind, config) {
        Ok(report) => report,
        Err(error) => {
            eprintln!("Rust 读取失败，尝试 COM：{error:#}");
            println!("Backend: Com (fallback)");
            return python::run_com(inputs, output, kind, target, config)
                .with_context(|| format!("自动提取失败；Rust 读取错误：{error:#}"));
        }
    };
    let native = match target {
        OutputKind::Excel => native::write_excel(&report, output, kind, config),
        OutputKind::Word => native::write_word(&report, output),
    };
    let python = if target == OutputKind::Excel {
        Some(|| python::write_excel(&report, output, kind, config))
    } else {
        None
    };
    fallback_writers(native, python, || {
        python::run_com(inputs, output, kind, target, config)
    })
}

fn fallback_writers<P, C>(native: Result<()>, python: Option<P>, com: C) -> Result<()>
where
    P: FnOnce() -> Result<()>,
    C: FnOnce() -> Result<()>,
{
    let error = match native {
        Ok(()) => return Ok(()),
        Err(error) => error,
    };
    let mut errors = format!("Rust: {error:#}");
    if let Some(python) = python {
        println!("Backend: Python (fallback)");
        match python() {
            Ok(()) => return Ok(()),
            Err(error) => errors.push_str(&format!("; openpyxl: {error:#}")),
        }
    }
    println!("Backend: Com (fallback)");
    com().with_context(|| format!("自动提取的全部路径均失败；{errors}"))
}

#[cfg(test)]
mod auto_route_tests {
    use super::*;
    use std::cell::RefCell;
    #[test]
    fn fallback_order_and_success_short_circuit() {
        let calls = RefCell::new(Vec::new());
        fallback_writers(
            Ok(()),
            Some(|| {
                calls.borrow_mut().push("python");
                Ok(())
            }),
            || {
                calls.borrow_mut().push("com");
                Ok(())
            },
        )
        .unwrap();
        assert!(calls.borrow().is_empty());
        fallback_writers(
            Err(anyhow::anyhow!("native failed")),
            Some(|| {
                calls.borrow_mut().push("python");
                Ok(())
            }),
            || {
                calls.borrow_mut().push("com");
                Ok(())
            },
        )
        .unwrap();
        assert_eq!(*calls.borrow(), ["python"]);
        calls.borrow_mut().clear();
        fallback_writers(
            Err(anyhow::anyhow!("native failed")),
            Some(|| {
                calls.borrow_mut().push("python");
                bail!("python failed")
            }),
            || {
                calls.borrow_mut().push("com");
                Ok(())
            },
        )
        .unwrap();
        assert_eq!(*calls.borrow(), ["python", "com"]);
        calls.borrow_mut().clear();
        fallback_writers(
            Err(anyhow::anyhow!("native failed")),
            None::<fn() -> Result<()>>,
            || {
                calls.borrow_mut().push("com");
                Ok(())
            },
        )
        .unwrap();
        assert_eq!(*calls.borrow(), ["com"]);
    }
}

fn route_backend(
    requested: BackendArg,
    output_kind: OutputKind,
    inputs: &[PathBuf],
    output: &Path,
    config: &Config,
) -> Result<BackendArg> {
    validate_output_extension(output, output_kind)?;
    if requested != BackendArg::Auto {
        validate_backend(requested, output_kind, inputs, output, config)?;
        return Ok(requested);
    }

    if inputs.iter().any(|path| extension_is(path, "doc"))
        || !["docx", "xlsx"].iter().any(|ext| extension_is(output, ext))
    {
        return Ok(BackendArg::Com);
    }

    Ok(BackendArg::Rust)
}

fn validate_backend(
    backend: BackendArg,
    output_kind: OutputKind,
    inputs: &[PathBuf],
    output: &Path,
    config: &Config,
) -> Result<()> {
    if backend == BackendArg::Rust {
        if inputs.iter().any(|path| extension_is(path, "doc")) {
            bail!(
                "Pure Rust currently supports .docx and .docm input, not legacy binary .doc. Use --backend com for .doc."
            );
        }
        if output_kind == OutputKind::Excel && !extension_is(output, "xlsx") {
            bail!("Pure Rust Excel output must use .xlsx.");
        }
        if output_kind == OutputKind::Word && !extension_is(output, "docx") {
            bail!("Pure Rust Word output must use .docx.");
        }
    }
    if backend == BackendArg::Python {
        if output_kind != OutputKind::Excel {
            bail!("The Python backend writes Excel output only.");
        }
        if !extension_is(output, "xlsx") {
            bail!("The Python/openpyxl backend writes .xlsx files only.");
        }
        if inputs.iter().any(|path| extension_is(path, "doc")) {
            bail!(
                "The Python backend does not read legacy binary .doc. Use --backend com or convert the file to .docx."
            );
        }
        if !python::probe_openpyxl(&config.backend.python_executable)? {
            bail!(
                "openpyxl is unavailable in Python interpreter '{}'.",
                config.backend.python_executable
            );
        }
    }
    if backend == BackendArg::Com {
        if !com_available(output_kind, &config.backend.python_executable)? {
            bail!(
                "The COM backend requires Microsoft Word; Excel output also requires Microsoft Excel, and Python with pywin32 must be available."
            );
        }
    }
    Ok(())
}

fn com_available(output: OutputKind, python: &str) -> Result<bool> {
    if !python::probe_word_com(python)? {
        return Ok(false);
    }
    if output == OutputKind::Excel {
        return python::probe_excel_com(python);
    }
    Ok(true)
}

fn validate_output_extension(path: &Path, output_kind: OutputKind) -> Result<()> {
    let extension = path
        .extension()
        .and_then(|ext| ext.to_str())
        .unwrap_or_default();
    let supported = ["xlsx", "xls", "xlsm", "docx", "doc", "docm"]
        .iter()
        .any(|expected| extension.eq_ignore_ascii_case(expected));
    if !supported {
        bail!(
            "Unsupported output extension '.{extension}'. Use .xlsx/.xls/.xlsm for Excel or .docx/.doc/.docm for Word."
        );
    }
    let extension_is_excel = ["xlsx", "xls", "xlsm"]
        .iter()
        .any(|ext| extension.eq_ignore_ascii_case(ext));
    ensure!(
        (output_kind == OutputKind::Excel) == extension_is_excel,
        "Output extension '.{extension}' does not match the selected output kind"
    );
    Ok(())
}

fn expand_inputs(inputs: &[PathBuf]) -> Result<Vec<PathBuf>> {
    let mut files = Vec::new();
    for input in inputs {
        if input.is_file() {
            ensure!(
                is_word_file(input),
                "Unsupported input file: {}",
                input.display()
            );
            files.push(input.clone());
            continue;
        }
        ensure!(input.is_dir(), "Input does not exist: {}", input.display());
        let directory_start = files.len();
        let mut pending = vec![input.clone()];
        while let Some(directory) = pending.pop() {
            for entry in fs::read_dir(&directory)
                .with_context(|| format!("Could not read {}", directory.display()))?
            {
                let path = entry?.path();
                if path.is_dir() {
                    pending.push(path);
                } else if !path
                    .file_name()
                    .and_then(|name| name.to_str())
                    .unwrap_or_default()
                    .starts_with("~$")
                    && is_word_file(&path)
                {
                    files.push(path);
                }
            }
        }
        files[directory_start..].sort();
    }
    let mut seen = std::collections::HashSet::new();
    files.retain(|p| seen.insert(p.clone()));
    ensure!(
        !files.is_empty(),
        "No .doc, .docx, or .docm inputs were found"
    );
    Ok(files)
}

fn is_word_file(path: &Path) -> bool {
    ["doc", "docx", "docm"]
        .iter()
        .any(|ext| extension_is(path, ext))
}

fn extension_is(path: &Path, expected: &str) -> bool {
    path.extension()
        .and_then(|extension| extension.to_str())
        .is_some_and(|extension| extension.eq_ignore_ascii_case(expected))
}
