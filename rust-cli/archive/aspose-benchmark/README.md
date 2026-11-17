# Aspose.NET experimental comparison

This optional backend uses Aspose.Words for .NET and Aspose.Cells for .NET 26.10.0. It is explicitly experimental and is not used by `auto`. It targets Windows/.NET 8 and supports the four form/output combinations with `.doc`, `.docx` or `.docm` input and `.xlsx`/`.docx` output.

**It is not VBA-equivalent or a replacement for the native writer.** Word output is text-only: source formatting, tables, pictures and hyperlinks are not preserved. Excel currently supports the basic wide layout only. TITLE priority and workbook filename mapping are rejected by the Rust adapter; some edge cases (nested controls, placeholders, duplicates, numeric values, question grouping and protected inputs) have not been validated against VBA. Use Rust or COM for actual work.

The normal unlicensed evaluation inserts an evaluation notice into Word and an `Evaluation Warning` worksheet into Excel; Aspose.Words also limits document size. These notices are left intact. If you have your own appropriate licenses, point `ASPOSE_WORDS_LICENSE` and `ASPOSE_CELLS_LICENSE` at the license files; they are not included in the project or binaries. See Aspose's [Words licensing](https://docs.aspose.com/words/net/licensing/) and [Cells evaluation limits](https://docs.aspose.com/cells/net/evaluation-version-limitations/).

Build with a .NET 8 SDK and restore packages from NuGet:

```powershell
dotnet build rust-cli/aspose-benchmark/aspose-benchmark.csproj -c Release
```

The Windows executable needs a .NET 8 runtime and its worker files. The worker is not included in the standalone Rust distribution; the CLI looks for `aspose-benchmark/bin/Release/net8.0/aspose-benchmark.exe` in the working directory or an executable ancestor. Alternatively set `ASPOSE_BENCHMARK_EXE` to the worker executable's absolute path. Then choose backend **5 Aspose 实验版** in the menu, or:

```powershell
rust-cli/target/release/vbaword-cli.exe extract --backend aspose --kind legacy --to excel --input ./forms --output ./aspose-test.xlsx
```

Run `python rust-cli/aspose-benchmark/benchmark.py --runs 5` and `python rust-cli/aspose-benchmark/benchmark.py --runs 3 --copies 10` after building the Rust release and worker. Each measurement launches a new CLI process (and an Aspose worker process for Aspose), includes start, parse and save, performs one warmup, and reports individual wall-clock times. The 30-file corpus is copies of the same three synthetic documents; it is not a diverse real-world workload. The generated `benchmark-report*.json` contain output data, paths, differences, checksums and timing samples. Documents and VBA files in the project are not altered.

Since the Aspose Word prototype is text-only and the Rust Word writer preserves formatting, the Word timings are not equivalent-work comparisons. Excel data-cell equality is reported separately from whole-file equality, which is affected by evaluation sheets. Trial-mode performance may differ from licensed mode. All measured processes are fresh; results do not describe a long-running resident Aspose service.
