"""Validate shipped executables, Linux installation and cross-platform output."""
import hashlib
import argparse
import json
import os
from pathlib import Path
import subprocess
import sys
import tempfile
import zipfile
from datetime import datetime, timezone

ROOT=Path(__file__).resolve().parents[1]
sys.path.insert(0,str(ROOT/'verification'))
from inspect_baseline import snapshot
from verify_opc import inspect


def run(command, **kwargs):
    result=subprocess.run([str(arg) for arg in command],capture_output=True,text=True,encoding='utf-8',errors='replace',**kwargs)
    assert result.returncode==0,result.stdout+'\n'+result.stderr
    return result.stdout+result.stderr


def signature(path):
    if path.suffix=='.xlsx':
        from openpyxl import load_workbook
        book=load_workbook(path,data_only=False)
        try:return {sheet.title:list(sheet.values) for sheet in book}
        finally:book.close()
    from lxml import etree
    with zipfile.ZipFile(path) as z:
        doc=etree.fromstring(z.read('word/document.xml'))
    graph=inspect(path)
    assert not graph['errors'],graph['errors']
    W='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    return {'text':[element.text for element in doc.iter(W+'t')],
            'table_text':[[e.text for e in table.iter(W+'t')] for table in doc.iter(W+'tbl')],
            'image_hashes':sorted(graph['images'].values()),
            'tables':graph['tables'],'horizontal_merges':graph['horizontal_merges'],
            'vertical_merges':graph['vertical_merges'],'drawings':graph['drawings'],
            'floating_images':graph['floating_images'],'cropped_images':graph['cropped_images']}


def main():
    parser=argparse.ArgumentParser()
    parser.add_argument('--current-baseline',action='store_true',help='Check the current originals for changes without replacing the historical manifest')
    args=parser.parse_args()
    dist=ROOT/'dist';output=ROOT/'target'/'package-check';output.mkdir(exist_ok=True)
    originals=snapshot()
    historical=json.loads((ROOT/'verification/evidence/original-sha256.json').read_text())
    if not args.current_baseline:assert originals==historical,'Original files differ from the historical manifest; use --current-baseline to preserve and verify the current files'
    fixtures=list((ROOT/'verification/fixtures').glob('*.docx'))
    hashes=lambda:{str(p):hashlib.sha256(p.read_bytes()).hexdigest() for p in fixtures}
    before=hashes();report={'passed':False,'containers':[],'comparisons':[],'historical_manifest_matches':originals==historical,'original_sha256_at_start':originals}
    manifest={line.split('  ',1)[1]:line.split('  ',1)[0] for line in (dist/'SHA256SUMS.txt').read_text(encoding='utf-8').splitlines()}
    for name,digest in manifest.items():assert hashlib.sha256((dist/name).read_bytes()).hexdigest()==digest
    report['artifacts']=manifest
    ref=output/'windows';ref.mkdir(exist_ok=True)
    try:
        with tempfile.TemporaryDirectory(prefix='vbaword-package-') as temp:
            with zipfile.ZipFile(dist/'Word汇总工具-0.1.0-windows-x86_64.zip') as archive:
                assert set(archive.namelist())=={'Word汇总工具.exe','config.xlsx'},archive.namelist()
                archive.extractall(temp)
            cli=Path(temp)/'Word汇总工具.exe'
            assert cli.read_bytes()==(dist/'Word汇总工具.exe').read_bytes()
            run([cli,'--version'])
            for kind,pattern in [('legacy','Rich_Legacy'),('content-control','Rich_CC')]:
                for to,extension in [('excel','xlsx'),('word','docx')]:
                    run([cli,'extract','--backend','rust','--kind',kind,'--to',to,'--input',
                         ROOT/'verification/fixtures'/f'{pattern}_1.docx',ROOT/'verification/fixtures'/f'{pattern}_2.docx',
                         '--output',ref/f'{kind}-{to}.{extension}'],env=dict(os.environ,PATH=''))
        for version in ('20.04','22.04'):
            directory=output/('ubuntu-'+version);directory.mkdir(exist_ok=True)
            command=['docker','run','--rm','-v',str(ROOT.parent)+':/project:ro','-v',str(directory)+':/out',
                     'ubuntu:'+version,'sh','/project/rust-cli/packaging/verify-linux.sh']
            log=run(command)
            (directory/'verification.log').write_text(log,encoding='utf-8')
            report['containers'].append({'image':'ubuntu:'+version,'install_extract_remove_passed':True,'log':str(directory/'verification.log')})
            for route in ('tar','deb'):
                for expected in sorted(ref.iterdir()):
                    actual=directory/route/expected.name
                    same=signature(expected)==signature(actual)
                    report['comparisons'].append({'image':'ubuntu:'+version,'package':route,'file':expected.name,'passed':same})
                    assert same,str(actual)
            assert not inspect(directory/'auto.docx')['errors']
        report['passed']=True
    finally:
        report['originals_unchanged']=snapshot()==originals
        report['fixtures_unchanged']=hashes()==before
        report['completed_utc']=datetime.now(timezone.utc).isoformat()
        (dist/'package-verification.json').write_text(json.dumps(report,indent=2),encoding='utf-8')
        assert report['originals_unchanged'] and report['fixtures_unchanged']
    print(json.dumps({'passed':report['passed'],'windows_operations':4,'linux_operations':18,'cross_platform_comparisons':len(report['comparisons']),'originals_unchanged':True}))


if __name__=='__main__':main()
