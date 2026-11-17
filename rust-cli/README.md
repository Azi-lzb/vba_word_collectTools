# vbaword Rust CLI

A standalone command-line extractor for Word legacy form fields and content controls. The project lives beside the existing VBA implementation and does not edit or import the workbook or VBA modules.

Build the executable from this directory with `cargo build --release`; on Windows it is written to `target\release\vbaword-cli.exe`.

Run without arguments to open the interactive menu. On Windows, double-click the executable to choose 1 File extraction and one of the four operations, then select files, a folder (current level or recursive), or entering paths manually. Extraction starts immediately and writes a uniquely named file to `output` beside the executable; existing files are not overwritten. Common operations (main menu 3) provides output-directory, latest-file and configuration shortcuts; its first two options open the output directory or its most recently modified Word/Excel file (excluding Office lock files). On Linux, use the manual path option. For system installations such as `/usr/bin`, output goes to `output` in the current directory. Existing command-line arguments, including `--output`, remain available.

Automatic routing uses Rust first for DOCX/DOCM inputs and DOCX/XLSX outputs. Excel writer failures try openpyxl and then COM; Word failures try COM. Reader failures go directly to COM because the openpyxl adapter shares the Rust reader. Binary DOC input and older output formats require COM. Legacy auto_excel preferences no longer change this order.

Main-menu option 2 creates example Word documents: three legacy form documents, three content-control documents, or both. Examples are saved beside the discovered VBA workbook, or beside the executable if no workbook is found; existing examples are never overwritten. Generated documents are saved, protected for forms with a blank password, and saved again. Examples can also be generated without Office using `vbaword-cli generate-examples --kind both --output ./examples`; `--kind legacy` and `--kind content-control` select one set.

## Four extraction operations

The two options combine into the four supported operations:

```powershell
vbaword-cli extract --kind legacy --to excel --input .\forms --output .\summary.xlsx
vbaword-cli extract --kind legacy --to word --input .\forms --output .\summary.docx
vbaword-cli extract --kind content-control --to excel --input .\forms --output .\summary.xlsx
vbaword-cli extract --kind content-control --to word --input .\forms --output .\summary.docx
```

`--input` accepts one or more `.doc`, `.docx`, `.docm` files or folders. Folder input is recursive. Office temporary files beginning with `~$` are skipped.

## Backends and format limits


| Backend | Inputs | Excel output | Word output | Needs Office / Python |
| --- | --- | --- | --- | --- |
| `rust` | `.docx`, `.docm` | `.xlsx` | `.docx` | Neither |
| `python` | `.docx`, `.docm` | `.xlsx` | — | Python and `openpyxl` |
| `com` | `.doc`, `.docx`, `.docm` | `.xls`, `.xlsx`, `.xlsm` | `.doc`, `.docx`, `.docm` | Python, `pywin32`, Word; Excel too for Excel output |
| `auto` | Chooses a backend from format and config | Chooses a backend from format and config | Chooses a backend from format and config | Depends on the selected route |

Examples of explicit backend selection:

```powershell
vbaword-cli extract --kind content-control --to excel --backend rust --input .\forms --output .\summary.xlsx
vbaword-cli extract --kind legacy --to word --backend com --input .\forms --output .\summary.docx
```

The COM Word writer copies source question and answer ranges. The pure Rust Word writer now imports the source OOXML resource graph and preserves embedded, floating and VML pictures, crop rectangles, external hyperlinks, internal bookmarks, and hyperlink field instructions. Resource paths, relationship IDs, bookmark names, drawing IDs, styles and numbering definitions are isolated between source documents.

Paragraph/character/table styles retain their inheritance and linked-style references. Source document defaults are carried through a private base style; source theme fonts and colors are materialized so two documents with different themes do not overwrite one another. Table grids, horizontal/vertical merges, nested tables, fixed widths, row heights, repeated headers, borders and shading remain structured Word content. Font-table resources and referenced package parts are copied as needed.

If a hyperlink points to a source bookmark outside the extracted content, the summary links back to that bookmark in the original document. Keep that source file accessible for those links. Links to bookmarks included in the summary are remapped within the summary.

Relative external hyperlinks (including HYPERLINK fields) are resolved against the source document folder or its HyperlinkBase before import. External-document bookmark names remain unchanged.

Excel output contains cell values. `Tables` is an optional normalized data sheet, not a conversion of Word page layout into Excel cells. `extract.table_mode = "flatten"` omits that additional sheet.

Pure Rust does not read the old binary `.doc` format. Auto mode sends `.doc` input to COM; without Word, convert it to `.docx` first. Rust and openpyxl write `.xlsx`, not the older `.xls` format.

## Configuration and routing

Main menu **3 → 3** opens the independent `config.xlsx`. It contains the VBA-style `config` and `mapping` sheets, initially migrated as values from the original workbook. In this standalone file, `config!B2` supplies the document password and `config!B3` supplies TAG/TITLE; `mapping!A:B` supplies legacy filename mappings from row 2. The original workbook is not modified. Save edits before starting the next extraction. An executable-local `config.xlsx` takes precedence, followed by the working directory and executable ancestors. The ZIP/tar packages include a configuration workbook; repackaging never overwrites an existing `dist/config.xlsx`. A standalone EXE can create a new configuration through **3 → 3**. `init-workbook-config --output config.xlsx` also creates it without opening Excel.

If no standalone configuration is found, the CLI reads `vbaForWordV3.xls` without modifying it from the working directory or an executable ancestor. This legacy fallback reads mapping A/B from row 2 and E1/F1 for password/priority, falling back to config B2/B3. Standalone `config.xlsx` instead uses config B2/B3 directly. Filename mapping follows VBA's legacy-form outputs; content-control outputs keep source filenames. `VBAWORD_DOC_PASSWORD` overrides the workbook password, and `--priority` overrides the workbook priority. Existing TOML priority is used only in the legacy fallback. Defaults apply if no workbook is available.

Create an example config:

```powershell
vbaword-cli init-config --output .\vbaword.toml
```

Auto mode always tries the native Rust reader/writer first for OOXML input and `.xlsx`/`.docx` output. Legacy `backend.auto_excel` settings are accepted for backward compatibility but do not change this order. `.doc` input/output, `.docm` output, `.xls`, and `.xlsm` output require COM. COM extraction needs Word; COM Excel output also needs Excel.

If native Excel writing fails, auto tries openpyxl, followed by COM; native Word failures fall back to COM. A reader failure skips openpyxl because it uses the same Rust reader. Explicit CLI backend selection reports its error without switching. The interactive menu always uses auto. `probe` and native extraction work even when Python is absent. Aspose has been removed from the application; its earlier experiment is archived under `archive/`.

Other options:

- `extract.priority`: choose content control `tag` or `title` as the field key.
- `extract.layout`: `wide` (one row per source) or `long` (one row per answer).
- `extract.table_mode`: `separate-sheet` or `flatten` for Excel output.
- `security.password_env`: environment variable holding a Word editing-protection password. The password is not written to the config.
- `security.mapping_csv`: optional UTF-8 file-name substring mapping, one `keyword,name` pair per line. Relative paths are resolved from the config file.

The COM adapter re-protects a document only after `Unprotect` succeeds. Source documents are opened for reading and closed without saving.

## VBA compatibility

The default wide output follows the existing workbook algorithms:

- Legacy Excel columns follow field position, with headings from the first selected document. Empty source documents still receive a row.
- Content-control Excel columns use the exact Tag/Title text, including trailing spaces; repeated keys in one document overwrite earlier values.
- Legacy Word output groups by the VBA-normalized question text and uses the same source labels and separators.
- Content-control Word output keeps document order and emits each heading only on its first occurrence, ignoring key case for heading deduplication.
- Empty form placeholders, checkbox/dropdown values, Word paragraph/cell markers and numeric Excel coercion are handled explicitly. Word result-range output follows VBA even when a checkbox/dropdown has no cached visible result.

Explicit input files keep their given order; recursive folder input is sorted. `long` layout and the additional `Tables` sheet are extensions to the original VBA output.

## Reproducible verification

Run `python verification/run_all.py` on Windows with Word, Excel, pywin32, openpyxl, python-docx, Pillow, lxml and PyMuPDF installed. Office is used by this verification harness as an independent reference and renderer; `--backend rust` itself needs neither Office nor Python.

The harness runs the workbook VBA on disposable copies with file-picker/message-box adapters, compares actual output cells and Word content, checks OPC relationship closure and ID uniqueness, and verifies effective formatting and PDF rendering. Reports, actual VBA outputs and PDFs are under `verification/evidence/`; see [verification/README.md](verification/README.md).

The corpus includes real repository samples and protected boundary cases, same-name images/styles/numbering from different documents, floating/cropped/VML images, hyperlink fields, omitted bookmark targets and complex tables. Theme tint/shade comparisons permit one RGB channel unit (1/255) of Word rendering rounding. This is evidence for the covered features, not a claim of pixel-identical pagination across different fonts or Office versions.

Known VBA failures are recorded separately: the original nested-control sample can raise error 4608 or 4605, and the rich block-control fixture can raise error 4605. Such cases are never reported as VBA parity passes; native fidelity is checked independently against the source documents. When rich-control VBA output succeeds, its raw comparison is retained: VBA clipboard insertion can add empty paragraphs and expose picture markers as `/`. The rich-content comparison explicitly normalizes those spacing/object-marker differences and permits the intentional repair of links to omitted source bookmarks. It is reported separately from exact output equality.

Inspect local backend availability with:

```powershell
vbaword-cli probe
```

## GUI decision

Slint is not needed for this batch-oriented tool. The CLI is easier to automate, script, and deploy, and it keeps the no-COM Rust route small. The extraction model and backends are separate modules so a Slint front end can be added later if interactive file selection, preview, or job history becomes a real need.
