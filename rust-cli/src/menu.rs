use crate::{BackendArg, ExtractArgs, FormKindArg, OutputKindArg, is_word_file, run_extract};
use anyhow::{Context, Result};
use std::{
    fs::{self, OpenOptions},
    io::{self, Write},
    path::{Path, PathBuf},
};

const OUTPUT_NAMES: [&str; 4] = [
    "旧式窗体汇总",
    "旧式窗体Word汇总",
    "内容控件汇总",
    "内容控件Word汇总",
];

fn output_directory() -> Result<PathBuf> {
    let exe = std::env::current_exe().context("无法定位程序所在目录")?;
    let parent = exe.parent().context("无法定位程序所在目录")?;
    // System package directories are not normally writable by regular users.
    #[cfg(target_os = "linux")]
    if ["/usr/bin", "/usr/local/bin", "/bin", "/sbin", "/usr/sbin"]
        .iter()
        .any(|system_dir| parent == Path::new(system_dir))
    {
        return Ok(std::env::current_dir()?.join("output"));
    }
    Ok(parent.join("output"))
}

fn reserve_output_in(dir: &Path, kind: FormKindArg, to: OutputKindArg) -> Result<PathBuf> {
    fs::create_dir_all(dir).with_context(|| format!("无法创建输出目录 {}", dir.display()))?;
    let name = match (kind, to) {
        (FormKindArg::Legacy, OutputKindArg::Excel) => OUTPUT_NAMES[0],
        (FormKindArg::Legacy, OutputKindArg::Word) => OUTPUT_NAMES[1],
        (FormKindArg::ContentControl, OutputKindArg::Excel) => OUTPUT_NAMES[2],
        (FormKindArg::ContentControl, OutputKindArg::Word) => OUTPUT_NAMES[3],
    };
    let extension = if matches!(to, OutputKindArg::Excel) {
        "xlsx"
    } else {
        "docx"
    };
    for number in 1u64.. {
        let suffix = if number == 1 {
            String::new()
        } else {
            format!("_{number}")
        };
        let path = dir.join(format!("{name}{suffix}.{extension}"));
        match OpenOptions::new().write(true).create_new(true).open(&path) {
            Ok(_) => return Ok(path),
            Err(error) if error.kind() == io::ErrorKind::AlreadyExists => continue,
            Err(error) => {
                return Err(error).with_context(|| format!("无法创建输出文件 {}", path.display()));
            }
        }
    }
    unreachable!()
}

fn latest_output_in(dir: &Path) -> Result<Option<PathBuf>> {
    if !dir.exists() {
        return Ok(None);
    }
    let mut latest = None;
    for entry in fs::read_dir(dir).with_context(|| format!("无法读取输出目录 {}", dir.display()))?
    {
        let entry = entry?;
        let path = entry.path();
        let Some(name) = path.file_name().and_then(|s| s.to_str()) else {
            continue;
        };
        let Some(ext) = path.extension().and_then(|s| s.to_str()) else {
            continue;
        };
        if name.starts_with("~$")
            || !["xlsx", "xls", "xlsm", "docx", "doc", "docm"]
                .iter()
                .any(|candidate| ext.eq_ignore_ascii_case(candidate))
        {
            continue;
        }
        let metadata = entry.metadata()?;
        if !metadata.is_file() || metadata.len() == 0 {
            continue;
        }
        let modified = metadata
            .modified()
            .unwrap_or(std::time::SystemTime::UNIX_EPOCH);
        if latest
            .as_ref()
            .is_none_or(|(time, previous): &(std::time::SystemTime, PathBuf)| {
                (modified, &path) > (*time, previous)
            })
        {
            latest = Some((modified, path));
        }
    }
    Ok(latest.map(|(_, path)| path))
}

#[cfg(windows)]
fn open_path(path: &Path) -> Result<()> {
    use std::{os::windows::ffi::OsStrExt, ptr};
    #[link(name = "shell32")]
    unsafe extern "system" {
        fn ShellExecuteW(
            hwnd: *mut std::ffi::c_void,
            operation: *const u16,
            file: *const u16,
            parameters: *const u16,
            directory: *const u16,
            show: i32,
        ) -> isize;
    }
    #[link(name = "ole32")]
    unsafe extern "system" {
        fn CoInitializeEx(reserved: *mut std::ffi::c_void, mode: u32) -> i32;
        fn CoUninitialize();
    }
    let wide: Vec<u16> = path.as_os_str().encode_wide().chain(Some(0)).collect();
    let verb: Vec<u16> = "open".encode_utf16().chain(Some(0)).collect();
    // ShellExecute opens directories in Explorer and files in the user's associated application.
    let initialized = unsafe { CoInitializeEx(ptr::null_mut(), 2) } >= 0;
    let result = unsafe {
        ShellExecuteW(
            ptr::null_mut(),
            verb.as_ptr(),
            wide.as_ptr(),
            ptr::null(),
            ptr::null(),
            1,
        )
    };
    if initialized {
        unsafe { CoUninitialize() };
    }
    anyhow::ensure!(
        result > 32,
        "无法打开 {}（Windows 错误代码 {result}）",
        path.display()
    );
    Ok(())
}

#[cfg(not(windows))]
fn open_path(path: &Path) -> Result<()> {
    #[cfg(target_os = "macos")]
    let launcher = "open";
    #[cfg(not(target_os = "macos"))]
    let launcher = "xdg-open";
    std::process::Command::new(launcher)
        .arg(path)
        .spawn()
        .with_context(|| format!("无法打开 {}", path.display()))?;
    Ok(())
}

fn ask(prompt: &str) -> Result<Option<String>> {
    print!("{prompt}");
    io::stdout().flush()?;
    let mut line = String::new();
    if io::stdin().read_line(&mut line)? == 0 {
        return Ok(None);
    }
    Ok(Some(line.trim().to_owned()))
}

#[cfg(windows)]
fn dialog(mode: &str) -> Result<Vec<PathBuf>> {
    use std::os::windows::process::CommandExt;
    let script = match mode {
        "files" => {
            "$d=New-Object System.Windows.Forms.OpenFileDialog; $d.Multiselect=$true; $d.Filter='Word (*.doc;*.docx;*.docm)|*.doc;*.docx;*.docm'; if($d.ShowDialog() -eq 'OK'){ $d.FileNames | ForEach-Object { [Console]::WriteLine($_) } }"
        }
        _ => {
            "$d=New-Object System.Windows.Forms.FolderBrowserDialog; if($d.ShowDialog() -eq 'OK'){ [Console]::WriteLine($d.SelectedPath) }"
        }
    };
    let command = format!(
        "Add-Type -AssemblyName System.Windows.Forms; [Console]::OutputEncoding=[Text.UTF8Encoding]::new(); {script}"
    );
    let out = std::process::Command::new("powershell.exe")
        .args(["-NoProfile", "-STA", "-Command", &command])
        .creation_flags(0x08000000)
        .output()?;
    anyhow::ensure!(out.status.success(), "文件选择窗口启动失败");
    Ok(String::from_utf8(out.stdout)?
        .lines()
        .filter(|s| !s.is_empty())
        .map(PathBuf::from)
        .collect())
}
#[cfg(not(windows))]
fn dialog(_: &str) -> Result<Vec<PathBuf>> {
    Ok(Vec::new())
}

fn select_action() -> Result<Option<String>> {
    loop {
        println!("\nWord 提取工具\n  1 文件提取\n  2 示例文档生成\n  3 常用操作\n  0 退出");
        let Some(group) = ask("请选择功能> ")? else {
            return Ok(None);
        };
        match group.as_str() {
            "0" => return Ok(None),
            "2" => return Ok(Some("5".into())),
            "1" | "3" => loop {
                if group == "1" {
                    println!(
                        "\n文件提取\n  1 旧式窗体 → Excel\n  2 旧式窗体 → Word\n  3 内容控件 → Excel\n  4 内容控件 → Word\n  0 返回主菜单"
                    );
                } else {
                    println!(
                        "\n常用操作\n  1 打开输出目录\n  2 打开最新输出文件\n  3 打开配置\n  0 返回主菜单"
                    );
                }
                let Some(item) = ask("请选择子功能> ")? else {
                    return Ok(None);
                };
                if item == "0" {
                    break;
                }
                if group == "1" && ["1", "2", "3", "4"].contains(&item.as_str()) {
                    return Ok(Some(item));
                }
                if group == "3" {
                    match item.as_str() {
                        "1" => return Ok(Some("6".into())),
                        "2" => return Ok(Some("7".into())),
                        "3" => return Ok(Some("8".into())),
                        _ => {}
                    }
                }
                println!("选项无效，请重新选择。");
            },
            _ => println!("请输入 0 至 3。"),
        }
    }
}

pub fn run() -> Result<()> {
    loop {
        let Some(choice) = select_action()? else {
            return Ok(());
        };
        if choice == "8" {
            match crate::config::ensure_config().and_then(|path| {
                open_path(&path)?;
                Ok(path)
            }) {
                Ok(path) => println!(
                    "配置文件：{}。修改并保存后，下次提取自动生效。",
                    path.display()
                ),
                Err(error) => println!("打开配置失败：{error:#}"),
            }
            continue;
        }
        if choice == "6" || choice == "7" {
            let result = (|| -> Result<()> {
                let dir = output_directory()?;
                if choice == "6" {
                    fs::create_dir_all(&dir)?;
                    open_path(&dir)?;
                    println!("输出目录：{}", dir.display());
                } else if let Some(file) = latest_output_in(&dir)? {
                    open_path(&file)?;
                    println!("已打开：{}", file.display());
                } else {
                    println!("输出目录中暂无 Word 或 Excel 结果文件：{}", dir.display());
                }
                Ok(())
            })();
            if let Err(error) = result {
                println!("打开失败：{error:#}");
            }
            continue;
        }
        if choice == "5" {
            println!(
                "\n  1 旧式窗体示例（3 份）\n  2 内容控件示例（3 份）\n  3 两种都生成（6 份）\n  0 返回"
            );
            let Some(kind) = ask("选择示例类型> ")? else {
                return Ok(());
            };
            let kind = match kind.as_str() {
                "1" => crate::samples::SampleKind::Legacy,
                "2" => crate::samples::SampleKind::ContentControl,
                "3" => crate::samples::SampleKind::Both,
                "0" => continue,
                _ => {
                    println!("选项无效。");
                    continue;
                }
            };
            let path = crate::config::find_workbook()
                .and_then(|p| p.parent().map(|p| p.to_path_buf()))
                .or_else(|| {
                    std::env::current_exe()
                        .ok()
                        .and_then(|p| p.parent().map(|p| p.to_path_buf()))
                })
                .unwrap_or(std::env::current_dir()?);
            println!("示例保存位置：{}", path.display());
            match crate::samples::generate(kind, &path) {
                Ok(files) => {
                    println!("已生成 {} 份示例：", files.len());
                    for file in files {
                        println!("  {}", file.display());
                    }
                }
                Err(e) => println!("生成失败：{e:#}"),
            }
            if ask("按回车返回主菜单，输入 0 退出> ")?.as_deref() == Some("0") {
                return Ok(());
            }
            continue;
        }
        let (kind, to) = match choice.as_str() {
            "0" => return Ok(()),
            "1" => (FormKindArg::Legacy, OutputKindArg::Excel),
            "2" => (FormKindArg::Legacy, OutputKindArg::Word),
            "3" => (FormKindArg::ContentControl, OutputKindArg::Excel),
            "4" => (FormKindArg::ContentControl, OutputKindArg::Word),
            _ => {
                println!("请输入 0 至 8。");
                continue;
            }
        };
        let backend = BackendArg::Auto;
        println!(
            "\n  1 选择文件（可多选）\n  2 选择文件夹（仅当前层）\n  3 选择文件夹（递归子目录）\n  4 手动输入路径（多个路径用分号分隔）\n  0 返回"
        );
        let Some(mode) = ask("请选择输入方式> ")? else {
            return Ok(());
        };
        if mode == "0" {
            continue;
        }
        if !["1", "2", "3", "4"].contains(&mode.as_str()) {
            println!("输入方式无效。");
            continue;
        }
        let mut paths = if mode == "4" {
            Vec::new()
        } else {
            match dialog(if mode == "1" { "files" } else { "folder" }) {
                Ok(p) => p,
                Err(e) => {
                    println!("{e:#}");
                    Vec::new()
                }
            }
        };
        if paths.is_empty() {
            println!("未选择路径，可手动输入；输入 0 返回。");
            let Some(raw) = ask("路径> ")? else {
                return Ok(());
            };
            if raw == "0" || raw.is_empty() {
                continue;
            }
            paths = raw
                .split(';')
                .map(|p| PathBuf::from(p.trim().trim_matches('"')))
                .collect();
        }
        if mode == "2" {
            let mut files = Vec::new();
            for path in paths {
                match std::fs::read_dir(&path) {
                    Ok(entries) => {
                        for entry in entries.flatten() {
                            let p = entry.path();
                            if p.is_file() && is_word_file(&p) {
                                files.push(p);
                            }
                        }
                    }
                    Err(e) => println!("无法读取 {}: {e}", path.display()),
                }
            }
            files.sort();
            paths = files;
        }
        if paths.is_empty() {
            println!("没有找到 Word 文件。");
            continue;
        }
        // Expand before reserving the output so a recursive input folder cannot
        // accidentally include the newly created placeholder document.
        let auto_output = match output_directory() {
            Ok(path) => path,
            Err(error) => {
                println!("无法定位输出目录：{error:#}");
                continue;
            }
        };
        let explicit_files: std::collections::HashSet<_> = paths
            .iter()
            .filter(|p| p.is_file())
            .filter_map(|p| fs::canonicalize(p).ok())
            .collect();
        let excluded_dir = fs::canonicalize(&auto_output).ok();
        let paths = match crate::expand_inputs(&paths) {
            Ok(mut paths) => {
                paths.retain(|path| {
                    let Ok(resolved) = fs::canonicalize(path) else {
                        return true;
                    };
                    explicit_files.contains(&resolved)
                        || excluded_dir
                            .as_ref()
                            .is_none_or(|dir| !resolved.starts_with(dir))
                });
                if paths.is_empty() {
                    println!("没有找到 Word 文件（自动输出目录已跳过）。");
                    continue;
                }
                paths
            }
            Err(e) => {
                println!("读取输入失败：{e:#}");
                continue;
            }
        };
        let output = match reserve_output_in(&auto_output, kind, to) {
            Ok(path) => path,
            Err(e) => {
                println!("创建输出失败：{e:#}");
                continue;
            }
        };
        println!("正在提取，输出：{}", output.display());
        let started = std::time::Instant::now();
        match run_extract(ExtractArgs {
            kind,
            to,
            backend,
            input: paths,
            output: output.clone(),
            config: None,
            priority: None,
        }) {
            Ok(()) => println!(
                "完成：{}，本次执行耗时 {:.2} 秒",
                output.display(),
                started.elapsed().as_secs_f64()
            ),
            Err(e) => {
                // This path was created exclusively by this run.
                if let Err(cleanup) = fs::remove_file(&output) {
                    if cleanup.kind() != io::ErrorKind::NotFound {
                        eprintln!("无法清理未完成输出 {}：{cleanup}", output.display());
                    }
                }
                println!(
                    "提取失败（耗时 {:.2} 秒）：{e:#}",
                    started.elapsed().as_secs_f64()
                );
            }
        }
        if ask("按回车返回主菜单，输入 0 退出> ")?.as_deref() == Some("0") {
            return Ok(());
        }
    }
}

#[cfg(test)]
mod tests {
    use super::*;
    #[test]
    fn unique_output_never_overwrites_and_latest_ignores_unrelated_files() {
        let temp = tempfile::tempdir().unwrap();
        let dir = temp.path();
        let first = reserve_output_in(dir, FormKindArg::Legacy, OutputKindArg::Excel).unwrap();
        fs::write(&first, b"first").unwrap();
        let second = reserve_output_in(dir, FormKindArg::Legacy, OutputKindArg::Excel).unwrap();
        assert_ne!(first, second);
        assert_eq!(fs::read(&first).unwrap(), b"first");
        fs::write(&second, b"second").unwrap();
        fs::write(dir.join("~$临时文件.xlsx"), b"office lock").unwrap();
        fs::write(dir.join("无关.txt"), b"unrelated").unwrap();
        assert_eq!(latest_output_in(dir).unwrap(), Some(second));
    }
}
