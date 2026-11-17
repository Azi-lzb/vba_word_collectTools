"""Time the four backend routes on the same protected generated samples."""
import hashlib
import json
import os
from pathlib import Path
import statistics
import subprocess
import tempfile
import time

ROOT=Path(__file__).resolve().parents[1]
exe=ROOT/'target/release/vbaword-cli.exe'
out=ROOT/'verification/evidence'


def main():
    work=Path(tempfile.mkdtemp(prefix='routes-',dir=ROOT/'target'))
    corpus=work/'examples'
    subprocess.run([str(exe),'generate-examples','--output',str(corpus)],check=True,capture_output=True)
    inputs={kind:sorted(corpus.glob(pattern)) for kind,pattern in [('legacy','旧式*.docx'),('content-control','内容*.docx')]}
    originals=[ROOT.parent/'vbaForWordV3.xls',*sorted((ROOT.parent/'VBA_Export').rglob('*.bas')),*sorted((ROOT.parent/'VBA_Export').rglob('*.cls'))]
    sha=lambda path:hashlib.sha256(path.read_bytes()).hexdigest()
    before={str(p):sha(p) for p in originals}
    source={str(p):sha(p) for files in inputs.values() for p in files}
    results=[]
    for kind,files in inputs.items():
        for to in ('excel','word'):
            timings={}
            for backend in (['com','python','rust','aspose'] if to=='excel' else ['com','rust','aspose']):
                elapsed=[]; ext='xlsx' if to=='excel' else 'docx'
                for attempt in range(4):
                    output=work/f'{kind}-{to}-{backend}-{attempt}.{ext}'
                    command=[str(exe),'extract','--kind',kind,'--to',to,'--backend',backend,'--input',*map(str,files),'--output',str(output)]
                    start=time.perf_counter()
                    process=subprocess.run(command,capture_output=True,text=True,encoding='utf8',errors='replace')
                    elapsed_seconds=time.perf_counter()-start
                    if process.returncode:raise RuntimeError(f'{kind}/{to}/{backend}: {process.stderr}')
                    if attempt:elapsed.append(elapsed_seconds)
                timings[backend]={'seconds':elapsed,'median_seconds':statistics.median(elapsed)}
                print(json.dumps({'kind':kind,'to':to,'backend':backend,'median_seconds':timings[backend]['median_seconds']}),flush=True)
            results.append({'kind':kind,'to':to,'timings':timings,'input_count':len(files)})
    assert before=={str(p):sha(p) for p in originals} and source=={str(p):sha(p) for files in inputs.values() for p in files}
    report={'description':'Windows, three protected synthetic Word documents, one warmup plus three new-process runs; end-to-end startup, parse and save. Aspose trial Word is text-only and not equivalent fidelity.',
            'runs':3,'results':results,'originals_unchanged':True,'inputs_unchanged':True,
            'aspose_license_words':bool(os.environ.get('ASPOSE_WORDS_LICENSE')),'aspose_license_cells':bool(os.environ.get('ASPOSE_CELLS_LICENSE'))}
    (out/'routing-benchmark.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    (ROOT/'verification'/'route-reference.json').write_text(json.dumps(report,ensure_ascii=False,indent=2),encoding='utf8')
    print('Saved routing benchmark',flush=True)


if __name__=='__main__':main()
