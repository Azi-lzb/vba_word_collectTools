"""Read effective Word formatting and export PDFs without saving input DOCX."""
import json
from pathlib import Path
import pythoncom
import win32com.client

HERE=Path(__file__).resolve().parent
OUT=HERE/'evidence'/'rendered'

def find(doc,text,start=0):
    rng=doc.Range(start,doc.Content.End)
    ok=rng.Find.Execute(FindText=text,MatchCase=True,Wrap=0)
    if not ok:return None
    return rng

def properties(rng):
    if rng is None:return None
    return {'font':rng.Font.Name,'size':rng.Font.Size,'color':rng.Font.Color,
            'bold':rng.Font.Bold,'italic':rng.Font.Italic,'space_after':rng.ParagraphFormat.SpaceAfter}

def inspect(word,path):
    doc=word.Documents.Open(FileName=str(path),ReadOnly=True,Visible=False,AddToRecentFiles=False,OpenAndRepair=False,ConfirmConversions=False)
    try:
        info={'path':str(path),'text':doc.Content.Text,'pages':doc.ComputeStatistics(2),
              'inline_images':[{'width':shape.Width,'height':shape.Height,'type':shape.Type} for shape in doc.InlineShapes],
              'floating_images':[{'width':shape.Width,'height':shape.Height,'type':shape.Type,'left':shape.Left,'top':shape.Top} for shape in doc.Shapes],
              'links':[{'address':h.Address,'anchor':h.SubAddress,'text':h.TextToDisplay} for h in doc.Hyperlinks],
              'markers':{}}
        for index in (1,2):
            marker=find(doc,'Source '+str(index)+' styled answer')
            if marker is None:continue
            info['markers']['source'+str(index)]=properties(marker)
            info['markers']['inherited'+str(index)]=properties(find(doc,'Inherited '+str(index)))
            theme=find(doc,'Theme text',marker.End)
            info['markers']['theme'+str(index)]=properties(theme)
            first=find(doc,'First list item',marker.End)
            second=find(doc,'Second list item',marker.End)
            info['markers']['list'+str(index)]=[{'label':r.ListFormat.ListString,'value':r.ListFormat.ListValue} if r is not None else None for r in (first,second)]
        info['tables']=[]
        for table in doc.Tables:
            data={'text':table.Range.Text,'nesting':table.NestingLevel}
            try:data['border_color']=table.Borders.Item(-1).Color
            except Exception as error:data['border_error']=str(error)
            info['tables'].append(data)
        pdf=OUT/(path.stem+'-'+path.parent.name+'.pdf')
        doc.ExportAsFixedFormat(OutputFileName=str(pdf),ExportFormat=17,OpenAfterExport=False)
        info['pdf']=str(pdf)
        return info
    finally:doc.Close(SaveChanges=0)

def main():
    OUT.mkdir(parents=True,exist_ok=True)
    paths=[HERE/'fixtures'/f'Rich_CC_{i}.docx' for i in (1,2)]
    paths += [HERE/'fixtures'/f'Rich_Legacy_{i}.docx' for i in (1,2)]
    paths += [HERE/'evidence'/'rich-cc-rust.docx',HERE/'evidence'/'rich-legacy-rust.docx']
    pythoncom.CoInitialize();word=None
    results=[]
    try:
        word=win32com.client.DispatchEx('Word.Application');word.Visible=False;word.DisplayAlerts=0
        for path in paths:
            results.append(inspect(word,path));print('Inspected '+path.name,flush=True)
    finally:
        if word is not None:word.Quit()
        pythoncom.CoUninitialize()
    (OUT/'word-effective-properties.json').write_text(json.dumps(results,ensure_ascii=True,indent=2),encoding='utf-8')
    output=next(r for r in results if Path(r['path']).name=='rich-cc-rust.docx')
    checks=[]
    for index in (1,2):
        source=next(r for r in results if Path(r['path']).name==f'Rich_CC_{index}.docx')
        for marker in ('source','theme','list','inherited'):
            key=marker+str(index)
            expected=source['markers'].get(key)
            actual=output['markers'].get(key)
            if marker=='theme':
                # Word returns an encoded theme token in Font.Color, not resolved RGB.
                expected={k:v for k,v in expected.items() if k!='color'}
                actual={k:v for k,v in actual.items() if k!='color'}
            checks.append({'check':key,'expected':expected,'actual':actual,'passed':expected==actual})
        import pymupdf
        def colors(pdf):
            document=pymupdf.open(pdf)
            return [span['color'] for page in document for block in page.get_text('dict')['blocks'] for line in block.get('lines',[]) for span in line['spans'] if 'Theme text' in span['text']]
        expected=colors(source['pdf'])[0]
        actual=colors(output['pdf'])[index-1]
        difference=max(abs((expected>>shift&255)-(actual>>shift&255)) for shift in (0,8,16))
        checks.append({'check':'rendered_theme_rgb'+str(index),'expected':hex(expected),'actual':hex(actual),'max_channel_difference':difference,'tolerance':1,'passed':difference<=1})
    expected_count=sum(len(source['inline_images'])+len(source['floating_images']) for source in results[:2])
    checks.append({'check':'com_accessible_images_including_vml','expected':expected_count,'actual':len(output['inline_images'])+len(output['floating_images']),'passed':len(output['inline_images'])+len(output['floating_images'])==expected_count,'note':'Word may omit floating images inside controls from Document.Shapes; PDF placement checks independently require all six images.'})
    for category in ('inline_images','floating_images'):
        expected=[image for result in results[:2] for image in result[category]]
        actual=output[category]
        checks.append({'check':category+'_geometry','expected':expected,'actual':actual,'passed':expected==actual})
    expected_links=sum(len(source['links']) for source in results[:2])
    checks.append({'check':'all_hyperlinks_including_fields','expected':expected_links,'actual':len(output['links']),'passed':len(output['links'])==expected_links})
    fallback=[link for link in output['links'] if link['anchor']=='OutsideTarget' and link['address']]
    checks.append({'check':'omitted_targets_link_back_to_source','expected':4,'actual':len(fallback),'passed':len(fallback)==4})
    relative=[link for link in output['links'] if 'relative' in link['text']]
    expected_target=str((HERE/'fixtures'/'linked-target.docx').resolve()).casefold()
    checks.append({'check':'relative_external_links_rebased_without_remapping_external_bookmarks','expected':4,'actual':relative,
                   'passed':len(relative)==4 and all(link['address'].casefold()==expected_target and link['anchor']=='SharedBookmark' for link in relative)})
    (OUT/'fidelity-checks.json').write_text(json.dumps(checks,ensure_ascii=True,indent=2),encoding='utf-8')
    for check in checks:print(json.dumps(check),flush=True)
    assert all(check['passed'] for check in checks),'Effective Word fidelity check failed'

if __name__=='__main__':main()
