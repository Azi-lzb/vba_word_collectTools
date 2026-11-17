"""Execute workbook VBA on disposable copies; replace only interactive UI plumbing.

The authoritative extraction code is read from the workbook, not reimplemented
in Python. Every transformed module is retained in evidence for audit.
"""
import hashlib
import json
import re
import shutil
from pathlib import Path
from inspect_baseline import ROOT, OUT, snapshot

TASKS = {
    "legacy_excel": ("BatchSummarizeWordForms", "Sample_Data_*.docx"),
    "legacy_word": ("SummarizeToNewWordDoc", "Sample_Data_*.docx"),
    "cc_excel": ("SummarizeContentControlsToExcel", "CC_Sample_*.docx"),
    "cc_word": ("SummarizeCCWithTablesToWord", "CC_Sample_*.docx"),
}

def literal(value):
    return '"' + str(value).replace('"', '""') + '"'

def replace_code(component, text):
    module = component.CodeModule
    if module.CountOfLines:
        module.DeleteLines(1, module.CountOfLines)
    module.AddFromString(text)

def run(task, macro, pattern, case=None, input_dir=None):
    import win32com.client
    directory = OUT / (case or task)
    directory.mkdir(parents=True, exist_ok=True)
    copy = directory / "baseline.xls"
    shutil.copy2(ROOT / "vbaForWordV3.xls", copy)
    excel = win32com.client.DispatchEx("Excel.Application")
    book = None
    result = {"task":task, "macro":macro}
    try:
        excel.Visible = False
        excel.DisplayAlerts = False
        excel.EnableEvents = False
        excel.AutomationSecurity = 1
        book = excel.Workbooks.Open(str(copy), UpdateLinks=0)
        project = book.VBProject
        inspection = json.loads((OUT / "workbook-inspection.json").read_text())
        for name in ("extractWord", "extractToWord", "extractCCtoExcel", "extractCCtoWordWithTables"):
            original = inspection["workbook_modules"][name].replace('\r\r\n','\n').replace('\r\n','\n').replace('\r','\n')
            adapted = original.replace("As FileDialog", "As Object")
            adapted = re.sub(r'Set fd = Application.FileDialog\([^\n]+\)', 'Set fd = New BaselineDialog', adapted)
            adapted = adapted.replace('Set wdApp = GetObject(, "Word.Application")', 'Set wdApp = baselineApp')
            adapted = re.sub(r'MsgBox [^\n]+', 'BaselineNotice', adapted)
            replace_code(project.VBComponents(name), adapted)
            (directory / (name + ".vba.txt")).write_text(adapted,encoding="utf-8")
        filters = project.VBComponents.Add(2)
        filters.Name = "BaselineFilters"
        replace_code(filters, 'Public Sub Clear()\nEnd Sub\nPublic Sub Add(ByVal description As String, ByVal extensions As String, Optional ByVal index As Variant)\nEnd Sub')
        dialog = project.VBComponents.Add(2)
        dialog.Name = "BaselineDialog"
        replace_code(dialog, '''Public Title As String
Public Property Get Filters() As BaselineFilters
Set Filters = New BaselineFilters
End Property
Public Property Get SelectedItems() As Collection
Set SelectedItems = baselinePaths
End Property
Public Function Show() As Long
Show = -1
End Function''')
        source_inputs = sorted((Path(input_dir).resolve() if input_dir else ROOT / "SampleForTest").glob(pattern))
        input_copies=directory/'inputs';input_copies.mkdir(exist_ok=True)
        inputs=[]
        for original in source_inputs:
            target=input_copies/original.name;shutil.copy2(original,target);inputs.append(target)
        result['source_inputs']=[str(path) for path in source_inputs]
        before_inputs={str(path):hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs}
        runner = project.VBComponents.Add(1)
        runner.Name = "BaselineRunner"
        code = 'Public baselinePaths As Collection\nPublic baselineApp As Object\nPublic Sub BaselineNotice()\nEnd Sub\nPublic Sub BaselineRun()\nOn Error GoTo Failed\nSet baselinePaths = New Collection\n'
        code += '\n'.join('baselinePaths.Add ' + literal(p) for p in inputs)
        code += '\nSet baselineApp = CreateObject("Word.Application")\nbaselineApp.DisplayAlerts = 0\nCall '+macro+'\n'
        if task.endswith('word'):
            code += 'Dim d As Object, n As Long\nn = 0\nFor Each d In baselineApp.Documents\nn = n + 1\nd.SaveAs2 FileName:='+literal(directory)+' & "\\word-" & n & ".docx", FileFormat:=12\nNext d\n'
        code += 'GoTo Cleanup\nFailed:\nThisWorkbook.Worksheets("config").Range("Z1").Value = Err.Number\nThisWorkbook.Worksheets("config").Range("Z2").Value = Err.Description\nCleanup:\nOn Error Resume Next\nbaselineApp.Quit SaveChanges:=0\nEnd Sub'
        replace_code(runner, code)
        (directory / "runner.vba.txt").write_text(code, encoding="utf-8")
        result['inputs'] = [str(p) for p in inputs]
        book.Worksheets('config').Range('Z1:Z3').ClearContents()
        excel.Run("'"+book.Name+"'!BaselineRun")
        error_number = book.Worksheets('config').Range('Z1').Value
        if error_number:
            raise RuntimeError('VBA runtime '+str(error_number)+': '+str(book.Worksheets('config').Range('Z2').Value))
        result['status'] = 'passed'
        if task.endswith('excel'):
            sheet = book.Worksheets('output' if task.startswith('legacy') else 'output_cc')
            result['cells'] = sheet.UsedRange.Value
        book.Save()
        assert before_inputs=={str(path):hashlib.sha256(path.read_bytes()).hexdigest() for path in inputs},'VBA input copies changed'
    except Exception as error:
        result['status'] = 'failed'
        result['error'] = str(error)
    finally:
        if book is not None:
            book.Close(SaveChanges=False)
        excel.Quit()
    (directory / 'result.json').write_text(json.dumps(result,ensure_ascii=True,indent=2,default=str),encoding='utf-8')
    print(json.dumps({k:v for k,v in result.items() if k not in ('cells','inputs')}),flush=True)
    return result

def main():
    import argparse
    parser = argparse.ArgumentParser()
    parser.add_argument('--task', choices=TASKS)
    parser.add_argument('--pattern')
    parser.add_argument('--case')
    parser.add_argument('--input-dir')
    parser.add_argument('--suite',choices=['edge'])
    args = parser.parse_args()
    import pythoncom
    before = json.loads((OUT / 'original-sha256.json').read_text())
    assert snapshot() == before
    pythoncom.CoInitialize()
    try:
        selected = {args.task:TASKS[args.task]} if args.task else TASKS
        results = [run(task, details[0], args.pattern or (('Edge_Legacy_*.docx' if task.startswith('legacy') else 'Edge_CC_*.docx') if args.suite=='edge' else details[1]), args.case or ('edge_'+task if args.suite else None), args.input_dir or (Path(__file__).parent/'fixtures' if args.suite else None)) for task,details in selected.items()]
    finally:
        pythoncom.CoUninitialize()
        assert snapshot() == before, 'Original VBA/workbook changed'
    (OUT / ((args.case or args.task or args.suite or 'vba-baseline')+'-summary.json')).write_text(json.dumps(results,ensure_ascii=True,indent=2,default=str),encoding='utf-8')

if __name__ == '__main__':
    main()
