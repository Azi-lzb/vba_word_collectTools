"""Reproducible end-to-end timings, including process startup, on generated samples."""
import argparse
import hashlib
import json
import os
import platform
import shutil
from pathlib import Path
import statistics
import subprocess
import tempfile
import time
import zipfile
from xml.etree import ElementTree as ET

ROOT=Path(__file__).resolve().parents[1]


def digest(path):return hashlib.sha256(path.read_bytes()).hexdigest()


def signature(path):
    if path.suffix=='.xlsx':
        import openpyxl
        book=openpyxl.load_workbook(path,data_only=True)
        try:
            return {'sheets':book.sheetnames,'data':{sheet.title:[list(row) for row in sheet.values] for sheet in book}}
        finally:book.close()
    with zipfile.ZipFile(path) as package:
        root=ET.fromstring(package.read('word/document.xml'))
        w='{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
        return {'text':[node.text or '' for node in root.iter(w+'t')],
                'tables':len(list(root.iter(w+'tbl'))),
                'images':len([name for name in package.namelist() if '/media/' in name])}


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--runs',type=int,default=5);parser.add_argument('--copies',type=int,default=1);args=parser.parse_args()
    if args.runs<1 or args.copies<1:parser.error('--runs and --copies must be positive')
    rust=ROOT/'target/release/vbaword-cli.exe'
    aspose=ROOT/'aspose-benchmark/bin/Release/net8.0/aspose-benchmark.exe'
    assert rust.is_file() and aspose.is_file(),'Build both release executables first'
    workspace=Path(tempfile.mkdtemp(prefix='aspose-comparison-',dir=ROOT/'target'))
    samples=workspace/'samples'
    subprocess.run([str(rust),'generate-examples','--output',str(samples)],check=True,capture_output=True)
    corpus={kind:sorted(samples.glob(pattern)) for kind,pattern in [('legacy','旧式*.docx'),('content-control','内容*.docx')]}
    if args.copies>1:
        for kind, originals in corpus.items():
            destination=workspace/('copies-'+kind);destination.mkdir()
            expanded=[]
            for index in range(args.copies):
                for source in originals:
                    target=destination/(f'{index:03d}-'+source.name);shutil.copyfile(source,target);expanded.append(target)
            corpus[kind]=expanded
    sample_hashes={str(path):digest(path) for paths in corpus.values() for path in paths}
    originals=[ROOT.parent/'vbaForWordV3.xls',*sorted((ROOT.parent/'VBA_Export').rglob('*.bas')),*sorted((ROOT.parent/'VBA_Export').rglob('*.cls'))]
    original_hashes={str(path):digest(path) for path in originals}
    report={'measurement':'end-to-end subprocess wall time, includes startup, parse and save; one warmup followed by measured runs',
            'warning':'Aspose Word output is text-only; Aspose trial adds evaluation notices and may limit output. Timings do not compare equivalent fidelity.',
            'runs':args.runs,'platform':platform.platform(),'processor':platform.processor(),
            'rust_executable_sha256':digest(rust),'aspose_executable_sha256':digest(aspose),
            'aspose_words_version':'26.10.0','aspose_cells_version':'26.10.0',
            'aspose_words_license_env_present':bool(os.environ.get('ASPOSE_WORDS_LICENSE')),
            'aspose_cells_license_env_present':bool(os.environ.get('ASPOSE_CELLS_LICENSE')),
            'sample_hashes':sample_hashes,'results':[]}
    for kind,files in corpus.items():
        assert len(files)==3*args.copies
        for target in ('excel','word'):
            ext='xlsx' if target=='excel' else 'docx';outputs={};timings={}
            for engine in ('rust','aspose'):
                elapsed=[]
                for attempt in range(args.runs+1):
                    output=workspace/f'{kind}-{target}-{engine}-{attempt}.{ext}'
                    command=[str(rust),'extract','--backend',engine,'--kind',kind,'--to',target,'--input',*map(str,files),'--output',str(output)]
                    start=time.perf_counter();result=subprocess.run(command,capture_output=True,text=True,encoding='utf8',errors='replace');duration=time.perf_counter()-start
                    if result.returncode!=0:raise RuntimeError(f'{engine} {kind}/{target}: {result.stderr}')
                    if attempt:elapsed.append(duration)
                    outputs[engine]=str(output)
                timings[engine]={'seconds':elapsed,'median_seconds':statistics.median(elapsed),'mean_seconds':statistics.mean(elapsed)}
            rust_sig=signature(Path(outputs['rust']));aspose_sig=signature(Path(outputs['aspose']))
            main_sheet='output' if kind=='legacy' else 'output_cc'
            data_equal=(rust_sig['data'].get(main_sheet)==aspose_sig['data'].get(main_sheet)) if target=='excel' else None
            report['results'].append({'kind':kind,'to':target,'input_count':len(files),'timings':timings,'main_sheet_values_equal':data_equal,
                'rust_over_aspose_median':timings['rust']['median_seconds']/timings['aspose']['median_seconds'],
                'output_exact_signature_equal':rust_sig==aspose_sig,'rust_signature':rust_sig,'aspose_signature':aspose_sig,'outputs':outputs})
            print(json.dumps({'kind':kind,'to':target,'rust_median':timings['rust']['median_seconds'],'aspose_median':timings['aspose']['median_seconds'],'signature_equal':rust_sig==aspose_sig}),flush=True)
    report['inputs_unchanged']=all(digest(Path(path))==hash for path,hash in sample_hashes.items())
    report['originals_unchanged']=all(digest(Path(path))==hash for path,hash in original_hashes.items())
    assert report['inputs_unchanged'] and report['originals_unchanged']
    target=ROOT/'aspose-benchmark'/('benchmark-report.json' if args.copies==1 else f'benchmark-report-{args.copies*3}-files.json')
    target.write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print('Report: '+str(target),flush=True)


if __name__=='__main__':main()
