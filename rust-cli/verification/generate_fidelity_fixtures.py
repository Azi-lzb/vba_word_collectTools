"""Deterministic collision fixtures; never modifies repository sample documents."""
from pathlib import Path
from PIL import Image, ImageDraw
from docx import Document
from docx.enum.style import WD_STYLE_TYPE
from docx.shared import Inches, Pt, RGBColor
from docx.oxml import OxmlElement
from docx.oxml.ns import qn
from docx.opc.constants import RELATIONSHIP_TYPE as RT
import zipfile
from lxml import etree

DEST=Path(__file__).resolve().parent/'fixtures'

def node(tag, **attrs):
    result=OxmlElement(tag)
    for name,value in attrs.items(): result.set(qn('w:'+name),str(value))
    return result

def link(paragraph,address,text,anchor=None):
    element=OxmlElement('w:hyperlink')
    if anchor: element.set(qn('w:anchor'),anchor)
    if address: element.set(qn('r:id'),paragraph.part.relate_to(address,RT.HYPERLINK,is_external=True))
    run=node('w:r');properties=node('w:rPr');properties.append(node('w:color',themeColor='hyperlink'))
    run.append(properties); t=node('w:t');t.text=text;run.append(t);element.append(run);paragraph._p.append(element)

def field_link(paragraph,target,text,split=False):
    if not split:
        field=node('w:fldSimple',instr=' HYPERLINK \\l "'+target+'" ');run=node('w:r');t=node('w:t');t.text=text;run.append(t);field.append(run);paragraph._p.append(field);return
    for kind in ('begin',):
        r=node('w:r');r.append(node('w:fldChar',fldCharType=kind));paragraph._p.append(r)
    code=' HYPERLINK \\l "'+target+'" '
    for value in (code[:21],code[21:]):
        r=node('w:r');t=node('w:instrText');t.text=value;r.append(t);paragraph._p.append(r)
    r=node('w:r');r.append(node('w:fldChar',fldCharType='separate'));paragraph._p.append(r)
    paragraph.add_run(text)
    r=node('w:r');r.append(node('w:fldChar',fldCharType='end'));paragraph._p.append(r)

def vml_picture(paragraph,picture):
    V='urn:schemas-microsoft-com:vml';O='urn:schemas-microsoft-com:office:office'
    rid,_=paragraph.part.get_or_add_image(str(picture))
    pict=node('w:pict')
    shape_type=etree.Element('{'+V+'}shapetype',nsmap={'v':V,'o':O})
    for name,value in {'id':'_x0000_t75','coordsize':'21600,21600','path':'m0,0l21600,0,21600,21600,0,21600xe','filled':'f','stroked':'f','{'+O+'}spt':'75'}.items():shape_type.set(name,value)
    pict.append(shape_type)
    shape=etree.Element('{'+V+'}shape');shape.set('id','SameImage');shape.set('type','#_x0000_t75');shape.set('style','width:36pt;height:18pt');shape.set('{'+O+'}spid','_x0000_s1025')
    data=etree.SubElement(shape,'{'+V+'}imagedata');data.set(qn('r:id'),rid);data.set('{'+O+'}title','Legacy VML picture');pict.append(shape)
    run=node('w:r');run.append(pict);paragraph._p.append(run)

def styles(doc,index):
    base=doc.styles.add_style('SharedBase',WD_STYLE_TYPE.PARAGRAPH)
    base.paragraph_format.space_after=Pt(25 if index==1 else 31)
    inherited=doc.styles.add_style('InheritedStyle',WD_STYLE_TYPE.PARAGRAPH);inherited.base_style=base
    style=doc.styles.add_style('SharedStyle',WD_STYLE_TYPE.PARAGRAPH)
    style.font.name='Arial' if index==1 else 'Courier New'
    style.font.size=Pt(12 if index==1 else 15)
    style.font.color.rgb=RGBColor.from_string('AA1122' if index==1 else '117733')
    style.paragraph_format.space_after=Pt(7 if index==1 else 13)
    table=doc.styles.add_style('SharedTable',WD_STYLE_TYPE.TABLE)
    properties=node('w:tblPr');properties.append(node('w:shd',val='clear',themeFill='accent1',themeFillTint='CC'))
    borders=node('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'): borders.append(node('w:'+edge,val='single',sz='12',themeColor='accent1'))
    properties.append(borders);table.element.append(properties)

def themed(run):
    properties=run._r.get_or_add_rPr()
    properties.append(node('w:color',themeColor='accent1',themeTint='99'))
    properties.append(node('w:rFonts',asciiTheme='minorHAnsi',hAnsiTheme='minorHAnsi',eastAsiaTheme='minorEastAsia'))

def floating(shape):
    inline=shape._inline
    anchor=OxmlElement('wp:anchor')
    for name,value in {'distT':'0','distB':'0','distL':'0','distR':'0','simplePos':'0','relativeHeight':'251658240','behindDoc':'0','locked':'0','layoutInCell':'1','allowOverlap':'1'}.items():anchor.set(name,value)
    position=OxmlElement('wp:simplePos');position.set('x','0');position.set('y','0');anchor.append(position)
    for direction,relative in [('H','column'),('V','paragraph')]:
        position=OxmlElement('wp:position'+direction);position.set('relativeFrom',relative);offset=OxmlElement('wp:posOffset');offset.text='0';position.append(offset);anchor.append(position)
    for child in list(inline):
        if child.tag in (qn('wp:extent'),qn('wp:effectExtent')):anchor.append(child)
    anchor.append(OxmlElement('wp:wrapTopAndBottom'))
    for child in list(inline):anchor.append(child)
    inline.getparent().replace(inline,anchor)

def change_theme(path,index):
    with zipfile.ZipFile(path) as z: entries={name:z.read(name) for name in z.namelist()}
    theme=etree.fromstring(entries['word/theme/theme1.xml'])
    ns={'a':'http://schemas.openxmlformats.org/drawingml/2006/main'}
    theme.find('.//a:clrScheme/a:accent1/a:srgbClr',ns).set('val','993322' if index==1 else '2244AA')
    theme.find('.//a:fontScheme/a:minorFont/a:latin',ns).set('typeface','Arial' if index==1 else 'Courier New')
    entries['word/theme/theme1.xml']=etree.tostring(theme,xml_declaration=True,encoding='UTF-8',standalone=True)
    with zipfile.ZipFile(path,'w',zipfile.ZIP_DEFLATED) as z:
        for name,data in entries.items(): z.writestr(name,data)

def main():
    DEST.mkdir(exist_ok=True)
    target=Document();p=target.add_paragraph('External document target')
    p._p.insert(0,node('w:bookmarkStart',id='7',name='SharedBookmark'));p._p.append(node('w:bookmarkEnd',id='7'))
    target.save(DEST/'linked-target.docx')
    for index in (1,2):
        image=Image.new('RGB',(160,80),'white');draw=ImageDraw.Draw(image)
        draw.rectangle((4,4,155,75),fill='red' if index==1 else 'blue');draw.text((12,28),'SOURCE '+str(index),fill='white')
        picture=DEST/('picture-'+str(index)+'.png');image.save(picture)
        doc=Document();styles(doc,index)
        question=doc.add_paragraph('Rich question: ',style='SharedStyle');question.runs[0].bold=True
        answer=doc.add_paragraph('Source '+str(index)+' styled answer',style='SharedStyle')
        doc.add_paragraph('Inherited '+str(index),style='InheritedStyle')
        themed(answer.add_run(' Theme text'))
        paragraph=doc.add_paragraph(style='SharedStyle');floating(paragraph.add_run().add_picture(str(picture),width=Inches(1.3)))
        link(paragraph,'https://example.com/source/'+str(index)+'?a=1&b=2',' external link ')
        paragraph._p.append(node('w:bookmarkStart',id='7',name='SharedBookmark'))
        paragraph.add_run('Bookmark target');paragraph._p.append(node('w:bookmarkEnd',id='7'))
        link(paragraph,'',' internal link ',anchor='SharedBookmark')
        field_link(paragraph,'SharedBookmark',' simple field link ')
        field_link(paragraph,'SharedBookmark',' split field link ',split=True)
        link(paragraph,'',' omitted bookmark link ',anchor='OutsideTarget')
        field_link(paragraph,'OutsideTarget',' omitted field link ')
        link(paragraph,'linked-target.docx',' relative document link ',anchor='SharedBookmark')
        external=node('w:fldSimple',instr=' HYPERLINK "linked-target.docx" \\l "SharedBookmark" ')
        run=node('w:r');text=node('w:t');text.text=' relative field link ';run.append(text);external.append(run);paragraph._p.append(external)
        vml_picture(doc.add_paragraph(style='SharedStyle'),picture)
        table=doc.add_table(rows=3,cols=3);table.style='SharedTable';table.autofit=False
        table.cell(0,0).merge(table.cell(0,1)).text='Horizontal merge '+str(index)
        table.cell(1,0).merge(table.cell(2,0)).text='Vertical merge'
        for row,col in ((0,2),(1,1),(1,2),(2,1),(2,2)):table.cell(row,col).text=f'Cell {index}/{row}/{col}'
        table.rows[0]._tr.get_or_add_trPr().append(node('w:tblHeader'))
        table.rows[1]._tr.get_or_add_trPr().append(node('w:trHeight',val='720',hRule='atLeast'))
        table.cell(1,2)._tc.get_or_add_tcPr().append(node('w:shd',val='clear',themeFill='accent1',themeFillTint='99'))
        nested=table.cell(1,1).add_table(rows=1,cols=2);nested.cell(0,0).text='Nested left';nested.cell(0,1).text='Nested right'
        cropped=table.cell(2,2).paragraphs[0].add_run().add_picture(str(picture),width=Inches(.55))
        fill=cropped._inline.find('.//'+qn('pic:blipFill'));rectangle=OxmlElement('a:srcRect');rectangle.set('l','10000');rectangle.set('r','10000');fill.insert(1,rectangle)
        numbering=doc.part.numbering_part.element
        abstract=node('w:abstractNum',abstractNumId='42');lvl=node('w:lvl',ilvl='0');lvl.append(node('w:start',val='1'));lvl.append(node('w:numFmt',val='decimal' if index==1 else 'upperRoman'));lvl.append(node('w:lvlText',val='%1.'));abstract.append(lvl);numbering.append(abstract)
        num=node('w:num',numId='42');num.append(node('w:abstractNumId',val='42'));numbering.append(num)
        for label in ('First list item','Second list item'):
            p=doc.add_paragraph(label,style='SharedStyle');pr=p._p.get_or_add_pPr();np=node('w:numPr');np.append(node('w:ilvl',val='0'));np.append(node('w:numId',val='42'));pr.append(np)
        body=doc._element.body
        control=node('w:sdt');properties=node('w:sdtPr');properties.append(node('w:tag',val='Rich'));properties.append(node('w:alias',val='Rich title'));properties.append(node('w:id',val=str(index)));control.append(properties)
        content=node('w:sdtContent')
        for child in list(body):
            if child is question._p or child.tag==qn('w:sectPr'):continue
            body.remove(child);content.append(child)
        control.append(content);body.insert(1,control)
        outside=doc.add_paragraph('Outside extracted control')
        outside._p.insert(0,node('w:bookmarkStart',id='99',name='OutsideTarget'));outside._p.append(node('w:bookmarkEnd',id='99'))
        path=DEST/f'Rich_CC_{index}.docx';doc.save(path)
        change_theme(path,index)
        # A separate legacy fixture carries images/links in the question range.
        legacy=Document();styles(legacy,index)
        p=legacy.add_paragraph('Legacy rich question ',style='SharedStyle');p.add_run().add_picture(str(picture),width=Inches(.8));link(p,'https://example.com/legacy/'+str(index),' link ')
        question_table=legacy.add_table(rows=1,cols=2);question_table.style='SharedTable';question_table.cell(0,0).text='Question table left';question_table.cell(0,1).text='Question table right'
        field=legacy.add_paragraph(style='SharedStyle')
        begin=node('w:fldChar',fldCharType='begin');data=node('w:ffData');data.append(node('w:name',val='RichLegacy'));data.append(node('w:enabled'));data.append(node('w:textInput'));begin.append(data)
        r=node('w:r');r.append(begin);field._p.append(r)
        r=node('w:r');t=node('w:instrText');t.text=' FORMTEXT ';r.append(t);field._p.append(r)
        r=node('w:r');r.append(node('w:fldChar',fldCharType='separate'));field._p.append(r)
        field.add_run('Legacy answer '+str(index)).italic=True
        r=node('w:r');r.append(node('w:fldChar',fldCharType='end'));field._p.append(r)
        path=DEST/f'Rich_Legacy_{index}.docx';legacy.save(path);change_theme(path,index)
    print('Generated four collision fixtures in '+str(DEST))

if __name__=='__main__':main()
