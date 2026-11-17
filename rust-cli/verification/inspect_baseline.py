"""Read the authoritative workbook and exports without saving either."""
import hashlib
import json
from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]
OUT = Path(__file__).resolve().parent / "evidence"

def snapshot():
    files = sorted((ROOT / "VBA_Export").rglob("*")) + [ROOT / "vbaForWordV3.xls"]
    return {str(p.relative_to(ROOT)): hashlib.sha256(p.read_bytes()).hexdigest()
            for p in files if p.is_file()}

def main():
    OUT.mkdir(exist_ok=True)
    before = snapshot()
    manifest = OUT / "original-sha256.json"
    if manifest.exists():
        assert json.loads(manifest.read_text()) == before, "Original files changed"
    else:
        manifest.write_text(json.dumps(before, indent=2), encoding="utf-8")
    import pythoncom
    import win32com.client
    pythoncom.CoInitialize()
    excel = None
    book = None
    report = {}
    try:
        excel = win32com.client.DispatchEx("Excel.Application")
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.AutomationSecurity = 3
        book = excel.Workbooks.Open(str(ROOT / "vbaForWordV3.xls"), ReadOnly=True, UpdateLinks=0)
        report["sheets"] = [book.Worksheets(i).Name for i in range(1, book.Worksheets.Count + 1)]
        try:
            components = {}
            for component in book.VBProject.VBComponents:
                module = component.CodeModule
                components[component.Name] = module.Lines(1, module.CountOfLines) if module.CountOfLines else ""
            report["workbook_modules"] = components
            report["vbproject_access"] = True
        except Exception as error:
            report["vbproject_access"] = False
            report["vbproject_error"] = str(error)
        report["mapping"] = [[book.Worksheets("mapping").Cells(r,c).Value for c in range(1,7)]
                             for r in range(1, min(20, book.Worksheets("mapping").UsedRange.Rows.Count)+1)]
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        if excel is not None:
            excel.Quit()
        pythoncom.CoUninitialize()
    assert before == snapshot(), "Original files changed during read"
    (OUT / "workbook-inspection.json").write_text(json.dumps(report, ensure_ascii=True, indent=2), encoding="utf-8")
    print(json.dumps({"vbproject_access": report.get("vbproject_access"), "sheets": report.get("sheets"), "originals_unchanged": True}))

if __name__ == "__main__":
    main()
