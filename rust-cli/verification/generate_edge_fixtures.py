"""Word-authored protected form fixtures, saved before and after protection."""
from pathlib import Path
import json
import pythoncom
import win32com.client

DEST=Path(__file__).resolve().parent/'fixtures'

def append(doc,text):
    end=doc.Content.End-1
    doc.Range(end,end).InsertAfter(text)

def save(doc,name,protection):
    path=DEST/name
    doc.SaveAs2(FileName=str(path),FileFormat=12)
    doc.Protect(Type=protection,NoReset=True,Password='')
    doc.Save()
    doc.Close(SaveChanges=0)

def main():
    DEST.mkdir(exist_ok=True)
    pythoncom.CoInitialize();word=None
    evidence={}
    try:
        word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0
        doc=word.Documents.Add(Visible=False);append(doc,'No legacy fields.');save(doc,'Edge_Legacy_0.docx',2)
        for index in (1,2):
            doc=word.Documents.Add(Visible=False)
            for name,kind in [('Text',70),('Check',71),('Choice',83),('Formula',70)]:
                append(doc,name+': ');end=doc.Content.End-1
                field=doc.FormFields.Add(Range=doc.Range(end,end),Type=kind);field.Name=name
                if kind==70:field.Result=('=1+2' if index==1 else '1e3') if name=='Formula' else (' 001 ' if index==1 else '')
                if kind==71:field.CheckBox.Value=index==1
                if kind==83:
                    field.DropDown.ListEntries.Add(Name='One');field.DropDown.ListEntries.Add(Name='Two')
                    field.DropDown.Value=2 if index==1 else 1
                append(doc,'\r')
            evidence['legacy'+str(index)]=[{'name':f.Name,'result':f.Result} for f in doc.FormFields]
            save(doc,f'Edge_Legacy_{index}.docx',2)
        doc=word.Documents.Add(Visible=False);append(doc,'No content controls.');save(doc,'Edge_CC_0.docx',3)
        for index in (1,2):
            doc=word.Documents.Add(Visible=False)
            values=[('Same','Title Same','First'),('Same','Title Same','Last'),('same','lowercase','lower'),('',' title trailing ','line1\vline2\tvalue'),('','','')]
            for number,(tag,title,value) in enumerate(values):
                append(doc,'Question '+str(number)+': ');end=doc.Content.End-1
                control=doc.ContentControls.Add(Type=1,Range=doc.Range(end,end));control.Tag=tag;control.Title=title;control.MultiLine=True
                control.Range.Text=value+str(index) if value else ''
                append(doc,'\r')
            evidence['cc'+str(index)]=[{'tag':c.Tag,'title':c.Title,'text':c.Range.Text} for c in doc.ContentControls]
            save(doc,f'Edge_CC_{index}.docx',3)
    finally:
        if word is not None:word.Quit(SaveChanges=0)
        pythoncom.CoUninitialize()
    (DEST/'edge-expected.json').write_text(json.dumps(evidence,ensure_ascii=True,indent=2),encoding='utf-8')
    print('Generated protected checkbox/dropdown/empty/duplicate/control-character fixtures')

if __name__=='__main__':main()
