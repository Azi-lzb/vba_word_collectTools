"""Compare source table geometry/crops and rendered colors with the Rust summary."""
import hashlib
import json
import zipfile
from pathlib import Path
import pymupdf
from lxml import etree
from verify_opc import inspect, NS, W

HERE=Path(__file__).resolve().parent
A='http://schemas.openxmlformats.org/drawingml/2006/main'

def document(path):
    with zipfile.ZipFile(path) as z:return etree.fromstring(z.read('word/document.xml'))

def attrs(element):
    return None if element is None else {etree.QName(k).localname:v for k,v in element.attrib.items()}

def geometry(root):
    tables=[]
    for table in root.findall('.//w:tbl',NS):
        rows=[]
        for row in table.findall('w:tr',NS):
            cells=[]
            for cell in row.findall('w:tc',NS):
                cells.append({name:attrs(cell.find('w:tcPr/w:'+name,NS)) for name in ('tcW','gridSpan','vMerge','vAlign','textDirection','noWrap')})
            rows.append({'cells':cells,'height':attrs(row.find('w:trPr/w:trHeight',NS)),'header':row.find('w:trPr/w:tblHeader',NS) is not None})
        tables.append({'grid':[attrs(col) for col in table.findall('w:tblGrid/w:gridCol',NS)],
                       'width':attrs(table.find('w:tblPr/w:tblW',NS)),'layout':attrs(table.find('w:tblPr/w:tblLayout',NS)), 'rows':rows})
    return tables

def rendered_colors(path):
    doc=pymupdf.open(path)
    colors=set()
    for page in doc:
        for shape in page.get_drawings():
            for key in ('color','fill'):
                color=shape.get(key)
                if color is not None and len(color)==3:colors.add(tuple(round(c*255) for c in color))
    return colors

def image_placements(path):
    document=pymupdf.open(path);rectangles=[]
    for page in document:
        for xref in {image[0] for image in page.get_images()}:
            rectangles.extend((round(rect.width,3),round(rect.height,3)) for rect in page.get_image_rects(xref))
    return sorted(rectangles)

def main():
    sources=[HERE/'fixtures'/f'Rich_CC_{i}.docx' for i in (1,2)]
    output=HERE/'evidence'/'rich-cc-rust.docx'
    expected_tables=[table for path in sources for table in geometry(document(path))]
    actual_tables=geometry(document(output))
    expected_crops=[attrs(e) for path in sources for e in document(path).findall('.//{'+A+'}srcRect')]
    actual_crops=[attrs(e) for e in document(output).findall('.//{'+A+'}srcRect')]
    graph=inspect(output)
    expected_images={hashlib.sha256((HERE/'fixtures'/f'picture-{i}.png').read_bytes()).hexdigest() for i in (1,2)}
    checks=[{'check':'relationship_closure_and_unique_ids','passed':not graph['errors'],'errors':graph['errors']},
            {'check':'unchanged_source_image_bytes','passed':set(graph['images'].values())==expected_images},
            {'check':'table_geometry_merges_headers_and_nested_tables','passed':expected_tables==actual_tables,'expected':expected_tables,'actual':actual_tables},
            {'check':'picture_crop_rectangles','passed':expected_crops==actual_crops,'expected':expected_crops,'actual':actual_crops}]
    rendered=HERE/'evidence'/'rendered'
    expected_rectangles=sorted(rect for source in sources for rect in image_placements(rendered/(source.stem+'-fixtures.pdf')))
    actual_rectangles=image_placements(rendered/'rich-cc-rust-evidence.pdf')
    checks.append({'check':'six_rendered_images_including_floating_and_vml','expected':expected_rectangles,'actual':actual_rectangles,'passed':len(expected_rectangles)==6 and expected_rectangles==actual_rectangles})
    expected_colors=set().union(*(rendered_colors(rendered/(p.stem+'-fixtures.pdf')) for p in sources))
    actual_colors=rendered_colors(rendered/'rich-cc-rust-evidence.pdf')
    unmatched=[color for color in expected_colors if not any(max(abs(a-b) for a,b in zip(color,candidate))<=1 for candidate in actual_colors)]
    checks.append({'check':'rendered_table_border_and_fill_colors','expected':sorted(expected_colors),'actual':sorted(actual_colors),'tolerance_per_channel':1,'unmatched':unmatched,'passed':not unmatched})
    # All answer text nodes, including nested table cells, must remain present.
    expected_text=[e.text for p in sources for c in document(p).findall('.//w:sdtContent',NS) for e in c.findall('.//w:t',NS) if e.text]
    actual_text=[e.text for e in document(output).findall('.//w:t',NS) if e.text]
    from collections import Counter
    missing=Counter(expected_text)-Counter(actual_text)
    checks.append({'check':'all_source_answer_text_preserved','missing':dict(missing),'passed':not missing})
    (HERE/'evidence'/'fidelity-structures.json').write_text(json.dumps(checks,ensure_ascii=True,indent=2),encoding='utf-8')
    for check in checks:print(json.dumps({'check':check['check'],'passed':check['passed']}))
    assert all(check['passed'] for check in checks),'Fidelity structure/render validation failed'

if __name__=='__main__':main()
