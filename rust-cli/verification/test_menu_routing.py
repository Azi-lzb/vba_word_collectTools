"""Exercise route-first menu choices without changing user documents."""
import os
from pathlib import Path
import subprocess
import tempfile
import shutil

ROOT=Path(__file__).resolve().parents[1]
EXE=ROOT/'target/release/vbaword-cli.exe'


def menu(sequence, exe=EXE):
    return subprocess.run([str(exe)],input=sequence.encode('utf8'),capture_output=True,env=os.environ.copy(),check=True).stdout.decode('utf8')


def main():
    with tempfile.TemporaryDirectory(prefix='route-menu-') as tmp:
        source=ROOT/'verification/fixtures/Rich_Legacy_1.docx'
        exe=Path(tmp)/EXE.name;shutil.copyfile(EXE,exe)
        folder=Path(tmp)/'output';folder.mkdir();old=folder/'旧式窗体汇总.xlsx';old.write_bytes(b'previous')
        output=menu(f'1\n1\n4\n{source}\n0\n',exe)
        assert (folder/'旧式窗体汇总_2.xlsx').is_file() and old.read_bytes()==b'previous'
        assert 'Backend: Rust' in output and '执行路由>' not in output and 'Aspose' not in output and '本次执行耗时' in output and '输出路径' not in output and '输入 y' not in output
    navigation=menu('3\n0\n1\n9\n0\n0\n')
    assert '常用操作' in navigation and '选项无效' in navigation
    print('Nested menus, route order, automatic output and no-overwrite passed')


if __name__=='__main__':main()
