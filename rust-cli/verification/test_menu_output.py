"""Exercise automatic menu output and open commands in an isolated directory."""
import os
from pathlib import Path
import shutil
import subprocess
import tempfile
import zipfile

ROOT=Path(__file__).resolve().parents[1]
EXE=ROOT/'target/release/vbaword-cli.exe'


def run(exe, inputs, env):
    result=subprocess.run([str(exe)],input=inputs.encode('utf8'),capture_output=True,env=env,cwd=ROOT,timeout=90)
    assert result.returncode==0,result.stderr.decode('utf8','replace')
    return result.stdout.decode('utf8','replace')


def main():
    with tempfile.TemporaryDirectory(prefix='vbaword-output-') as temp:
        root=Path(temp);exe=root/'vbaword-cli.exe';shutil.copyfile(EXE,exe)
        out=root/'output';out.mkdir()
        env=dict(os.environ)
        files=[ROOT/'verification/fixtures/Rich_Legacy_1.docx',ROOT/'verification/fixtures/Rich_CC_1.docx']
        old=out/'旧式窗体汇总.xlsx';old.write_bytes(b'previous output')
        cases=[('1',files[0],'旧式窗体汇总_2.xlsx'),('2',files[0],'旧式窗体Word汇总.docx'),
               ('3',files[1],'内容控件汇总.xlsx'),('4',files[1],'内容控件Word汇总.docx')]
        for kind,source,name in cases:
            text=run(exe,f'1\n{kind}\n4\n{source}\n0\n',env)
            assert '输出路径' not in text and '输入 y' not in text and '完成：' in text,text
            assert zipfile.is_zipfile(out/name),name
        assert old.read_bytes()==b'previous output'
        # A recursive source containing the output folder must not re-import a
        # previous (or partial) Word output as a source.
        shutil.copyfile(files[0],root/'input.docx');(out/'invalid.docx').write_bytes(b'invalid')
        text=run(exe,f'1\n2\n4\n{root}\n0\n',env)
        assert '完成：' in text and zipfile.is_zipfile(out/'旧式窗体Word汇总_2.docx'),text
    print('Four automatic outputs, duplicate protection and recursive exclusion passed')


if __name__=='__main__':main()
