"""Reproduce VBA parity and native OOXML fidelity evidence on Windows Office."""
import hashlib
import argparse
from datetime import datetime
import json
import os
from pathlib import Path
import subprocess
import sys
import time
from inspect_baseline import ROOT, OUT, snapshot

PROJECT=ROOT/'rust-cli'

def hashes(paths):
    return {str(path.relative_to(ROOT)):hashlib.sha256(path.read_bytes()).hexdigest() for path in paths}

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--resume',action='store_true',help='Reuse successful commands from the last run after validating original and baseline input hashes')
    args=parser.parse_args()
    sys.stdout.reconfigure(encoding='utf-8',errors='replace')
    OUT.mkdir(exist_ok=True)
    original=snapshot()
    assert original==json.loads((OUT/'original-sha256.json').read_text()),'Original workbook/VBA changed'
    samples=sorted((ROOT/'SampleForTest').glob('*.docx'))
    sample_hashes=hashes(samples)
    previous=None
    if args.resume:
        previous=json.loads((OUT/'verification-summary.json').read_text())
        assert previous['original_sha256']==original and previous['sample_sha256']==sample_hashes
        implementation=list((PROJECT/'src').glob('*.rs'))+[PROJECT/'python'/'worker.py',PROJECT/'Cargo.toml',PROJECT/'Cargo.lock']
        cutoff=datetime.fromisoformat(previous['started_utc'].replace('Z','+00:00')).timestamp()
        assert all(path.stat().st_mtime<=cutoff for path in implementation),'Implementation changed; run full verification'
        if previous.get('fixture_hashes'):
            assert hashes(sorted((PROJECT/'verification'/'fixtures').glob('*.docx')))==previous['fixture_hashes']
        # Input copies used by the actual VBA run are immutable reference inputs.
        for result_path in OUT.glob('*/result.json'):
            baseline=json.loads(result_path.read_text())
            for source,copy in zip(baseline.get('source_inputs',[]),baseline.get('inputs',[])):
                assert Path(source).read_bytes()==Path(copy).read_bytes(),f'Baseline input changed: {source}'
    report={'started_utc':time.strftime('%Y-%m-%dT%H:%M:%SZ',time.gmtime()),'steps':[],'passed':False}
    env=dict(os.environ,PYTHONIOENCODING='utf-8')
    def run(*command):
        key=[str(c) for c in command]
        reused=next((step for step in previous['steps'] if step['command']==key and step['exit_code']==0),None) if previous else None
        if reused:
            report['steps'].append(dict(reused,reused=True))
            print('REUSE '+json.dumps(key,ensure_ascii=True),flush=True)
            return
        started=time.monotonic();print('RUN '+json.dumps(key,ensure_ascii=True),flush=True)
        with subprocess.Popen([str(c) for c in command],cwd=PROJECT,env=env,stdout=subprocess.PIPE,stderr=subprocess.STDOUT,text=True,encoding='utf-8',errors='replace',bufsize=1) as process:
            for line in process.stdout:
                print(line,end='',flush=True);log.write(line);log.flush()
            code=process.wait()
        report['steps'].append({'command':[str(c) for c in command],'exit_code':code,'seconds':round(time.monotonic()-started,2)})
        if code:raise RuntimeError('Verification command failed: '+str(command))
    def python(script,*args):run(sys.executable,'verification/'+script,*args)
    with (OUT/'verification-run.log').open('a' if previous else 'w',encoding='utf-8') as log:
        try:
            run('cargo','fmt','--check')
            run('cargo','test','--offline')
            run('cargo','build','--offline')
            python('inspect_baseline.py')
            python('generate_fidelity_fixtures.py')
            python('generate_edge_fixtures.py')
            python('verify_native_independence.py')
            fixtures=sorted((PROJECT/'verification'/'fixtures').glob('*.docx'))
            fixture_hashes=hashes(fixtures)
            report['fixture_hashes']=fixture_hashes
            python('run_vba_baseline.py')
            failure=json.loads((OUT/'cc_word'/'result.json').read_text())
            if failure['status']=='failed':
                assert any(code in failure['error'] for code in ('4605','4608')),failure
                report['expected_vba_failures']=[{'case':'cc_word','error':failure['error']}]
            else:
                report['expected_vba_failures']=[]
            python('run_vba_baseline.py','--task','cc_word','--pattern','CC_Sample_[23].docx','--case','cc_word_valid')
            python('run_vba_baseline.py','--suite','edge')
            python('run_vba_baseline.py','--task','legacy_word','--input-dir','verification/fixtures','--pattern','Rich_Legacy_*.docx','--case','rich_legacy_word')
            python('run_vba_baseline.py','--task','cc_word','--input-dir','verification/fixtures','--pattern','Rich_CC_*.docx','--case','rich_cc_word')
            rich_failure=json.loads((OUT/'rich_cc_word'/'result.json').read_text())
            if rich_failure['status']=='failed':
                assert '4605' in rich_failure['error'],rich_failure
                report['expected_vba_failures'].append({'case':'rich_cc_word','error':rich_failure['error']})
            parity_reports=[]
            for backend in ('rust','python','com'):
                destination='verification/evidence/parity-final-'+backend+'.json'
                python('compare_parity.py','--backend',backend,'--output',destination)
                parity_reports.append(destination)
                destination='verification/evidence/parity-edge-'+backend+'.json'
                args=[arg for case in ('edge_legacy_excel','edge_legacy_word','edge_cc_excel','edge_cc_word') for arg in ('--case',case)]
                python('compare_parity.py','--backend',backend,*args,'--output',destination)
                parity_reports.append(destination)
            for backend in ('rust','com'):
                destination='verification/evidence/parity-rich-legacy-'+backend+'.json'
                python('compare_parity.py','--backend',backend,'--case','rich_legacy_word','--output',destination)
                parity_reports.append(destination)
            for case,baseline in [('cc_word',failure),('rich_cc_word',rich_failure)]:
                if baseline['status']=='passed':
                    destination='verification/evidence/parity-'+case+'-extra.json'
                    python('compare_parity.py','--case',case,'--output',destination);parity_reports.append(destination)
            cli=PROJECT/'target'/'debug'/'vbaword-cli.exe'
            for kind,pattern,name in [('content-control','Rich_CC_*.docx','rich-cc-rust.docx'),('legacy','Rich_Legacy_*.docx','rich-legacy-rust.docx')]:
                inputs=sorted((PROJECT/'verification'/'fixtures').glob(pattern))
                run(cli,'extract','--kind',kind,'--to','word','--backend','rust','--input',*inputs,'--output',OUT/name)
            documents=[OUT/'rich-cc-rust.docx',OUT/'rich-legacy-rust.docx']
            for case in ('legacy_word','cc_word_valid','edge_legacy_word','edge_cc_word','rich_legacy_word'):documents.append(OUT/case/'rust.docx')
            python('verify_opc.py',*map(str,documents),'--output',str(OUT/'opc-final-report.json'))
            python('inspect_word_fidelity.py')
            python('verify_fidelity_structures.py')
            assert hashes(fixtures)==fixture_hashes,'Word-authored source fixtures were modified'
            report['fixture_hashes']=fixture_hashes
            report['parity_reports']=parity_reports
            report['parity_cases_passed']=sum(len(json.loads((PROJECT/path).read_text())) for path in parity_reports)
            report['passed']=True
        except Exception as error:
            report['error']=str(error)
            raise
        finally:
            report['originals_unchanged']=snapshot()==original
            report['repository_samples_unchanged']=hashes(samples)==sample_hashes
            report['original_sha256']=original
            report['sample_sha256']=sample_hashes
            report['finished_utc']=time.strftime('%Y-%m-%dT%H:%M:%SZ',time.gmtime())
            (OUT/'verification-summary.json').write_text(json.dumps(report,ensure_ascii=True,indent=2),encoding='utf-8')
            assert report['originals_unchanged'] and report['repository_samples_unchanged'],'Original files changed'
    print(json.dumps({'passed':report['passed'],'parity_cases':report['parity_cases_passed'],'originals_unchanged':True}),flush=True)

if __name__=='__main__':main()
