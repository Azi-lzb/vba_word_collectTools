"""Run all four native operations with no PATH and an invalid Python backend."""
import hashlib
import argparse
import json
import os
from pathlib import Path
import subprocess

HERE=Path(__file__).resolve().parent
PROJECT=HERE.parent

def main():
    parser=argparse.ArgumentParser();parser.add_argument('--release',action='store_true');args=parser.parse_args()
    label='native-independence'+('-release' if args.release else '')
    directory=HERE/'evidence'/label;directory.mkdir(exist_ok=True)
    config=directory/'no-python.toml'
    config.write_text('[backend]\nauto_excel = "com"\npython_executable = '+json.dumps(str(directory/'does-not-exist.exe'))+'\n',encoding='utf-8')
    cli=PROJECT/'target'/('release' if args.release else 'debug')/'vbaword-cli.exe'
    environment=dict(os.environ,PATH='',PYTHONIOENCODING='utf-8')
    cases=[]
    for backend in ('rust','auto'):
        for kind,source in [('legacy','Edge_Legacy_1.docx'),('content-control','Rich_CC_1.docx')]:
            for output_kind,extension in [('excel','xlsx'),('word','docx')]:
                path=HERE/'fixtures'/source;before=hashlib.sha256(path.read_bytes()).hexdigest()
                output=directory/f'{backend}-{kind}.{extension}'
                command=[str(cli),'extract','--backend',backend,'--kind',kind,'--to',output_kind,'--input',str(path),'--output',str(output),'--config',str(config)]
                result=subprocess.run(command,env=environment,capture_output=True,text=True,encoding='utf-8',errors='replace')
                passed=result.returncode==0 and 'Backend: Rust' in result.stdout and output.exists()
                assert hashlib.sha256(path.read_bytes()).hexdigest()==before
                cases.append({'backend':backend,'kind':kind,'to':output_kind,'passed':passed,'stdout':result.stdout,'stderr':result.stderr})
    probe=subprocess.run([str(cli),'probe'],env=environment,capture_output=True,text=True,encoding='utf-8',errors='replace')
    checks={'executable':str(cli),'executable_sha256':hashlib.sha256(cli.read_bytes()).hexdigest(),'PATH_empty':True,'configured_python_does_not_exist':True,'cases':cases,'probe':{'passed':probe.returncode==0 and 'Rust OOXML: true' in probe.stdout,'stdout':probe.stdout}}
    checks['passed']=all(case['passed'] for case in cases) and checks['probe']['passed']
    (HERE/'evidence'/(label+'.json')).write_text(json.dumps(checks,ensure_ascii=True,indent=2),encoding='utf-8')
    print(json.dumps({'passed':checks['passed'],'operations':len(cases),'probe_passed':checks['probe']['passed']}))
    assert checks['passed'],'Native/no-Python automatic routing check failed'

if __name__=='__main__':main()
