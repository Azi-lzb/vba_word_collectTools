"""Validate relationship closure, ID isolation and preserved rich structures."""
import argparse
import hashlib
import json
import posixpath
import zipfile
from pathlib import Path
from urllib.parse import unquote
from lxml import etree

W='http://schemas.openxmlformats.org/wordprocessingml/2006/main'
R='http://schemas.openxmlformats.org/officeDocument/2006/relationships'
REL='http://schemas.openxmlformats.org/package/2006/relationships'
WP='http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing'
NS={'w':W,'r':R,'wp':WP}

def owner(rels):
    if rels=='_rels/.rels':return ''
    directory,name=rels.rsplit('/_rels/',1)
    return directory+'/'+name[:-5]

def inspect(path):
    errors=[]
    with zipfile.ZipFile(path) as z:
        names=z.namelist()
        if len(names)!=len(set(names)):errors.append('duplicate ZIP entries')
        parts={name:z.read(name) for name in names}
    roots={}
    for name,data in parts.items():
        if name.endswith(('.xml','.rels')):
            try:roots[name]=etree.fromstring(data)
            except etree.XMLSyntaxError as e:errors.append(f'{name}: {e}')
    relations={}
    for name,root in roots.items():
        if not name.endswith('.rels'):continue
        part=owner(name)
        ids={}
        for rel in root:
            ident=rel.get('Id')
            if ident in ids:errors.append(f'duplicate relationship {part}#{ident}')
            target=rel.get('Target')
            external=rel.get('TargetMode')=='External'
            resolved=target if external else posixpath.normpath(posixpath.join(posixpath.dirname(part),unquote(target).split('#')[0]))
            resolved=resolved.lstrip('/')
            if not external and resolved not in parts:errors.append(f'missing target {part}#{ident}: {resolved}')
            ids[ident]={'type':rel.get('Type'),'target':target,'resolved':resolved,'external':external}
        relations[part]=ids
    for name,root in roots.items():
        if name.endswith('.rels'):continue
        for element in root.iter():
            for attr,value in element.attrib.items():
                if attr in ('{'+R+'}id','{'+R+'}embed','{'+R+'}link') and value not in relations.get(name,{}):
                    errors.append(f'unresolved XML relationship {name}#{value}')
    document=roots['word/document.xml']
    styles=roots.get('word/styles.xml')
    style_ids=[e.get('{'+W+'}styleId') for e in styles.findall('w:style',NS)] if styles is not None else []
    if len(style_ids)!=len(set(style_ids)):errors.append('duplicate style IDs')
    for element in document.xpath('//w:pStyle|//w:rStyle|//w:tblStyle',namespaces=NS):
        if element.get('{'+W+'}val') not in style_ids:errors.append('undefined style '+str(element.get('{'+W+'}val')))
    num=roots.get('word/numbering.xml')
    nums=num.findall('w:num',NS) if num is not None else []
    abstracts=num.findall('w:abstractNum',NS) if num is not None else []
    num_ids=[e.get('{'+W+'}numId') for e in nums]
    abstract_ids=[e.get('{'+W+'}abstractNumId') for e in abstracts]
    if len(num_ids)!=len(set(num_ids)):errors.append('duplicate numbering IDs')
    if len(abstract_ids)!=len(set(abstract_ids)):errors.append('duplicate abstract numbering IDs')
    for element in document.findall('.//w:numId',NS):
        value=element.get('{'+W+'}val')
        if value!='0' and value not in num_ids:errors.append('undefined numId '+value)
    for element in nums:
        target=element.find('w:abstractNumId',NS)
        if target is None or target.get('{'+W+'}val') not in abstract_ids:errors.append('undefined abstract numbering')
    bookmarks=document.findall('.//w:bookmarkStart',NS)
    bookmark_names=[e.get('{'+W+'}name') for e in bookmarks]
    bookmark_ids=[e.get('{'+W+'}id') for e in bookmarks]
    if len(bookmark_ids)!=len(set(bookmark_ids)):errors.append('duplicate bookmark IDs')
    if len(bookmark_names)!=len(set(bookmark_names)):errors.append('duplicate bookmark names')
    for e in document.findall('.//w:hyperlink',NS):
        anchor=e.get('{'+W+'}anchor')
        if anchor and e.get('{'+R+'}id') is None and anchor not in bookmark_names and anchor not in ('_top','_bottom','_startOfDoc','_endOfDoc'):errors.append('missing bookmark '+anchor)
    end_ids=[e.get('{'+W+'}id') for e in document.findall('.//w:bookmarkEnd',NS)]
    if sorted(bookmark_ids)!=sorted(end_ids):errors.append('unpaired bookmark markers')
    vml_ids=[e.get('id') for e in document.iter() if etree.QName(e).namespace=='urn:schemas-microsoft-com:vml' and e.get('id')]
    if len(vml_ids)!=len(set(vml_ids)):errors.append('duplicate VML IDs')
    drawing_ids=[e.get('id') for e in document.findall('.//wp:docPr',NS)]
    if len(drawing_ids)!=len(set(drawing_ids)):errors.append('duplicate drawing IDs')
    image_targets={r['resolved'] for r in relations.get('word/document.xml',{}).values() if r['type'].endswith('/image') and not r['external']}
    return {'path':str(path),'errors':errors,'images':{p:hashlib.sha256(parts[p]).hexdigest() for p in sorted(image_targets)},
            'external_hyperlinks':[r['target'] for r in relations.get('word/document.xml',{}).values() if r['type'].endswith('/hyperlink')],
            'tables':len(document.findall('.//w:tbl',NS)),'horizontal_merges':len(document.findall('.//w:gridSpan',NS)),
            'vertical_merges':len(document.findall('.//w:vMerge',NS)),'bookmarks':bookmark_names,'drawings':len(drawing_ids),
            'styles':len(style_ids),'numbering':len(num_ids),
            'floating_images':len(document.findall('.//wp:anchor',NS)),
            'cropped_images':len(document.findall('.//{http://schemas.openxmlformats.org/drawingml/2006/main}srcRect'))}

def main():
    parser=argparse.ArgumentParser();parser.add_argument('documents',nargs='+');parser.add_argument('--output',required=True)
    args=parser.parse_args();reports=[inspect(Path(p)) for p in args.documents]
    Path(args.output).write_text(json.dumps(reports,ensure_ascii=True,indent=2),encoding='utf-8')
    for report in reports:print(json.dumps(report,ensure_ascii=True))
    assert not any(r['errors'] for r in reports),'OPC integrity validation failed'

if __name__=='__main__':main()
