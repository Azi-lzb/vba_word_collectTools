# VBA parity and OOXML fidelity verification

The original workbook and exported VBA remain immutable. `inspect_baseline.py`
records SHA256 values and reads the workbook VBA through Excel without saving.
`run_vba_baseline.py` copies the workbook and executes that workbook's original
algorithms with UI-only adapters for file selection, message boxes and the Word
application instance. The adapted source is retained for audit. An outer handler
records VBA runtime errors; extraction statements are not replaced with Python.

`compare_parity.py` runs the native CLI on the same ordered input paths, then
uses read-only Word and Excel COM inspection to compare actual output content.
Excel comparisons ignore unused blank cells that only have formatting. Word
comparisons retain paragraph boundaries, filenames, punctuation and table text.
Images and hyperlinks are recorded as additional structural evidence.

The rich block-control case additionally retains `raw_equal` and an explicit
`checks` map. Its content comparison normalizes VBA clipboard blank paragraphs,
trailing spaces and picture-object markers; the only allowed hyperlink address
changes are an originally local bookmark repaired to its own source file, or a
relative external link resolved against the source folder. This
case must be reported separately from exact equality. Source bookmark targets,
relative external links and external-document anchors are checked independently.

Current authoritative cases:

- `legacy_excel`: all seven legacy samples, including differing field names and counts.
- `legacy_word`: the same seven samples, grouped by source question text.
- `cc_excel`: all three content-control samples, including duplicate tags, trailing spaces and a table.
- `cc_word_valid`: samples 2 and 3, preserving document order and first-seen heading behavior.
- `cc_word`: the original VBA can fail on sample 1's nested controls with runtime error 4608 or 4605. A failed run has no complete VBA output and remains a baseline failure, not a parity pass.
- `edge_*`: protected empty sources, checkbox/dropdown defaults, placeholder text, formula/numeric values, duplicate keys, trailing spaces and control characters.
- `rich_legacy_word`: complete tables, pictures and hyperlinks in source questions.
- `rich_cc_word`: same-name resources across two sources, style inheritance, two themes, list numbering, floating/cropped/VML images, fields and local/external hyperlinks, merged/nested tables and repeated headers. VBA may fail with 4605; a successful run is compared with the documented normalization above.

Example runs from `rust-cli`:

```powershell
python verification/inspect_baseline.py
python verification/run_vba_baseline.py
python verification/run_vba_baseline.py --task cc_word --pattern CC_Sample_[23].docx --case cc_word_valid
python verification/compare_parity.py
python verification/run_all.py
cargo build --release --offline
python verification/verify_native_independence.py --release
python verification/summarize_results.py
```

Office calls need an interactive Windows session with Word/Excel COM access.
Generated workbooks, documents, adapted module text, hashes and comparison JSON
are kept under `verification/evidence/`. Reports must not mark a baseline error
or an unexecuted fidelity check as successful.

`verify_opc.py` checks recursive package relationships, identifiers and references.
`inspect_word_fidelity.py` opens DOCX without repair, checks effective formatting,
all hyperlink targets, and exports PDF without saving the sources.
`verify_fidelity_structures.py` compares image bytes, six rendered image placements,
cropping, table grids/merges/nesting/header geometry, and rendered borders/fills.
Theme color rendering permits a maximum difference of 1/255 per RGB channel.
`verify_native_independence.py` runs four Rust and four automatic operations with
an empty PATH and nonexistent Python, and checks `probe`.

The aggregate report is `RESULTS.md`; underlying evidence is retained in
`evidence/verification-summary.json`, `evidence/parity-*.json`,
`evidence/opc-final-report.json`, `evidence/fidelity-structures.json` and
`evidence/rendered/`. Generated evidence and fixtures are excluded from source
control by `.gitignore`; no Git commands are needed to run verification.
