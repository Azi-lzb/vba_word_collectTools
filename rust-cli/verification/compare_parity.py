"""Compare actual Rust output against previously executed workbook VBA."""
import argparse
import difflib
import json
import subprocess
from pathlib import Path
from inspect_baseline import ROOT, OUT, snapshot

CLI = ROOT / 'rust-cli' / 'target' / 'debug' / 'vbaword-cli.exe'

def trim_grid(rows):
    rows = [list(row) for row in rows]
    while rows and all(v is None for v in rows[-1]):
        rows.pop()
    width = max((i+1 for row in rows for i,v in enumerate(row) if v is not None),default=0)
    return [row[:width] + [None]*max(0,width-len(row)) for row in rows]

def word_signature(word, path):
    doc = word.Documents.Open(FileName=str(path), ReadOnly=True, Visible=False, AddToRecentFiles=False)
    try:
        return {'text':doc.Content.Text,
                'tables':[{'rows':t.Rows.Count,'columns':t.Columns.Count,'text':t.Range.Text}
                          for t in doc.Tables],
                'images':doc.InlineShapes.Count + doc.Shapes.Count,
                'hyperlinks':[h.Address for h in doc.Hyperlinks]}
    finally:
        doc.Close(SaveChanges=0)


def rich_equivalence(expected, actual, inputs):
    # VBA's clipboard path adds blank paragraphs, trims text differently and exposes
    # VML object markers as '/' in Word.Range.Text. Retain both raw signatures.
    def content(text):
        return [('\x01' if line == '/' else line.rstrip())
                for line in text.split('\r') if line.strip()]
    checks={'content_ignoring_clipboard_spacing_and_object_marker':content(expected['text'])==content(actual['text']),
            'table_structure_and_text':expected['tables']==actual['tables'],
            'image_count':expected['images']==actual['images'],
            'hyperlink_count':len(expected['hyperlinks'])==len(actual['hyperlinks'])}
    allowed={str(Path(path)).casefold() for path in inputs}
    def addresses_match(left,right):
        if left==right or (not left and right.casefold() in allowed):return True
        if left and not Path(left).is_absolute() and ':' not in left:
            return any(str((Path(path).parent/left).resolve()).casefold()==right.casefold() for path in inputs)
        return False
    # Only previously local links may become original-file links. Their precise
    # anchors and all included targets are separately checked by the fidelity suite.
    checks['hyperlink_addresses']=len(expected['hyperlinks'])==len(actual['hyperlinks']) and all(
        addresses_match(left,right)
        for left,right in zip(expected['hyperlinks'],actual['hyperlinks']))
    return checks

def main():
    import pythoncom
    import win32com.client
    parser=argparse.ArgumentParser()
    parser.add_argument('--case',action='append')
    parser.add_argument('--backend',choices=['rust','com','python'],default='rust')
    parser.add_argument('--output')
    args=parser.parse_args()
    before=snapshot()
    assert before == json.loads((OUT/'original-sha256.json').read_text())
    pythoncom.CoInitialize()
    excel=word=None
    results=[]
    try:
        excel=win32com.client.DispatchEx('Excel.Application')
        excel.Visible=False
        excel.DisplayAlerts=False
        excel.AutomationSecurity=3
        word=win32com.client.DispatchEx('Word.Application')
        word.Visible=False
        word.DisplayAlerts=0
        for case in args.case or ['legacy_excel','legacy_word','cc_excel','cc_word_valid']:
            directory=OUT/case
            baseline=json.loads((directory/'result.json').read_text())
            if baseline['status'] != 'passed':
                results.append({'case':case,'status':'baseline-failed'})
                continue
            task=baseline['task']
            is_excel=task.endswith('excel')
            if args.backend=='python' and not is_excel:
                continue
            output=directory/(args.backend+('.xlsx' if is_excel else '.docx'))
            command=[str(CLI),'extract','--backend',args.backend,'--kind','legacy' if task.startswith('legacy') else 'content-control',
                     '--to','excel' if is_excel else 'word','--input',*baseline['inputs'],'--output',str(output)]
            subprocess.run(command,check=True)
            item={'case':case,'command':command}
            if is_excel:
                book=excel.Workbooks.Open(str(output),ReadOnly=True,UpdateLinks=0)
                try:
                    sheet=book.Worksheets('output' if task.startswith('legacy') else 'output_cc')
                    actual=trim_grid(sheet.UsedRange.Value)
                finally:
                    book.Close(SaveChanges=False)
                expected=trim_grid(baseline['cells'])
                item.update(expected=expected,actual=actual,passed=expected==actual)
            else:
                expected=word_signature(word,directory/'word-1.docx')
                actual=word_signature(word,output)
                item.update(expected=expected,actual=actual,passed=expected==actual)
                if case=='rich_cc_word':
                    checks=rich_equivalence(expected,actual,baseline['inputs'])
                    item.update(raw_equal=expected==actual,comparison_scope='content, tables, images; explicit clipboard spacing/object-marker normalization, source-bookmark repair and relative-link rebasing',checks=checks,passed=all(checks.values()))
                (directory/'text-diff.txt').write_text(''.join(difflib.unified_diff(expected['text'].splitlines(True),actual['text'].splitlines(True),fromfile='vba',tofile='rust')),encoding='utf-8')
            results.append(item)
            print(json.dumps({'case':case,'passed':item['passed']}),flush=True)
    finally:
        if word is not None: word.Quit()
        if excel is not None: excel.Quit()
        pythoncom.CoUninitialize()
        assert snapshot()==before
    report=Path(args.output) if args.output else OUT/('parity-report'+('' if args.backend=='rust' else '-'+args.backend)+'.json')
    report.write_text(json.dumps(results,ensure_ascii=True,indent=2),encoding='utf-8')
    assert all(item.get('passed',False) for item in results),'VBA parity comparison failed'

if __name__=='__main__':
    main()
