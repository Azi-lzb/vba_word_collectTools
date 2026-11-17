"""Package prebuilt Windows x64 and Linux x86_64 executables; Python stdlib only."""
import argparse
import gzip
import hashlib
import io
import json
from pathlib import Path, PurePosixPath
import shutil
import subprocess
import tarfile
import tomllib
import zipfile

ROOT = Path(__file__).resolve().parents[1]
VERSION = tomllib.loads((ROOT / 'Cargo.toml').read_text())['package']['version']


def config(platform):
    return ('[backend]\nauto_excel = "rust"\npython_executable = "' +
            ('python' if platform == 'windows' else 'python3') +
            '"\n\n[extract]\npriority = "tag"\nlayout = "wide"\n'
            'table_mode = "separate-sheet"\npreserve_format = true\n\n'
            '[security]\npassword_env = "VBAWORD_DOC_PASSWORD"\n').encode()


def guide(platform):
    cli = '.\\Word汇总工具.exe' if platform == 'windows' else './Word汇总工具'
    return f'''Word汇总工具 {VERSION} ({platform} x86_64)

Four extraction operations (replace forms with your input folder):
{cli} extract --kind legacy --to excel --backend rust --input forms --output legacy.xlsx
{cli} extract --kind legacy --to word --backend rust --input forms --output legacy.docx
{cli} extract --kind content-control --to excel --backend rust --input forms --output controls.xlsx
{cli} extract --kind content-control --to word --backend rust --input forms --output controls.docx

--backend rust: no Python or Office. DOCX/DOCM input, DOCX/XLSX output.
--backend auto: choose from formats and configuration, with Rust fallback for OOXML.
--backend com: Windows only, requires Python + pywin32 + Word (Excel for Excel output).
Old binary .doc input requires Windows COM.
The interactive menu automatically uses Rust first, then openpyxl for Excel output,
and finally COM. Legacy binary formats require COM. No route selection is needed.
Interactive output is saved to an output folder beside the executable with a
unique name. Main menu 3 contains: 1 open output folder, 2 open latest result,
3 open configuration.
For Linux system installations such as /usr/bin, output is under the current directory.
Run without arguments and choose main menu 2 to create sample documents, or run:
{cli} generate-examples --kind both --output examples
Use --config vbaword.example.toml to use the included configuration.
Linux optional Python backend: use python3 in config and install openpyxl yourself.
The Python worker is embedded in the executable; no separate worker file is needed.
Keep output files outside input folders to avoid re-importing earlier summaries.
Run {cli} --help or {cli} probe for details.

Linux tar.gz: tar -xzf Word汇总工具-{VERSION}-linux-x86_64.tar.gz
Debian/Ubuntu: sudo apt install ./Word汇总工具_{VERSION}_amd64.deb
Debian package installs /usr/bin/Word汇总工具; run Word汇总工具 without ./.
Requires Linux x86_64, glibc >= 2.28. Alpine/musl is not supported by this build.
'''.encode()


def archive(entries):
    stream = io.BytesIO()
    with tarfile.open(fileobj=stream, mode='w', format=tarfile.USTAR_FORMAT) as tar:
        directories={str(parent) for name in entries for parent in PurePosixPath(name).parents if str(parent)!='.'}
        for name in sorted(directories):
            item=tarfile.TarInfo(name+'/');item.type=tarfile.DIRTYPE;item.mode=0o755
            item.uid=item.gid=item.mtime=0;item.uname=item.gname='root';tar.addfile(item)
        for name, (data, mode) in sorted(entries.items()):
            item = tarfile.TarInfo(name)
            item.size, item.mode, item.mtime = len(data), mode, 0
            item.uid = item.gid = 0
            item.uname = item.gname = 'root'
            tar.addfile(item, io.BytesIO(data))
    return gzip.compress(stream.getvalue(), mtime=0)


def deb_file(path, control, data):
    # Debian's standard ar container: format marker, control tarball, data tarball.
    with path.open('wb') as out:
        out.write(b'!<arch>\n')
        for name, payload in [('debian-binary', b'2.0\n'), ('control.tar.gz', control), ('data.tar.gz', data)]:
            header = f'{name+"/":<16}{0:<12}{0:<6}{0:<6}{"100644":<8}{len(payload):<10}`\n'
            out.write(header.encode('ascii')); out.write(payload)
            if len(payload) % 2: out.write(b'\n')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--windows-binary', type=Path)
    parser.add_argument('--linux-binary', type=Path)
    parser.add_argument('--output', type=Path, default=ROOT/'dist')
    args = parser.parse_args()
    if not (args.windows_binary or args.linux_binary): parser.error('Provide at least one executable')
    args.output.mkdir(parents=True, exist_ok=True)
    outputs=[]
    project_config=ROOT.parent/'config.xlsx'
    installed_config=args.output/'config.xlsx'
    if project_config.is_file() and not installed_config.exists():shutil.copyfile(project_config,installed_config)
    if args.windows_binary and not installed_config.exists():
        subprocess.run([str(args.windows_binary.resolve()),'init-workbook-config','--output',str(installed_config.resolve())],check=True)
    config_entries={'config.xlsx':installed_config.read_bytes()} if installed_config.is_file() else {}
    if installed_config.is_file():outputs.append(installed_config)
    if args.windows_binary:
        binary=args.windows_binary.read_bytes(); assert binary[:2]==b'MZ', 'Expected PE executable'
        exe=args.output/'Word汇总工具.exe';shutil.copyfile(args.windows_binary,exe);outputs.append(exe)
        zip_path=args.output/f'Word汇总工具-{VERSION}-windows-x86_64.zip'
        with zipfile.ZipFile(zip_path,'w',zipfile.ZIP_DEFLATED) as zip:
            for name,content in {'Word汇总工具.exe':binary,**config_entries}.items():
                item=zipfile.ZipInfo(name,date_time=(1980,1,1,0,0,0));item.compress_type=zipfile.ZIP_DEFLATED
                zip.writestr(item,content)
        outputs.append(zip_path)
    if args.linux_binary:
        binary=args.linux_binary.read_bytes()
        assert binary[:4]==b'\x7fELF' and binary[4]==2 and binary[18:20]==b'\x3e\x00', 'Expected x86_64 ELF'
        top=f'Word汇总工具-{VERSION}-linux-x86_64'
        tgz=args.output/(top+'.tar.gz')
        tgz.write_bytes(archive({top+'/Word汇总工具':(binary,0o755),top+'/README.txt':(guide('linux'),0o644),top+'/vbaword.example.toml':(config('linux'),0o644),**{top+'/'+name:(content,0o644) for name,content in config_entries.items()}}));outputs.append(tgz)
        entries={'usr/bin/Word汇总工具':(binary,0o755),
                 'usr/share/doc/vbaword-cli/README.txt':(guide('linux'),0o644),
                 'usr/share/doc/vbaword-cli/vbaword.example.toml':(config('linux'),0o644)}
        installed=(sum(len(data) for data,mode in entries.values())+1023)//1024
        control=f'''Package: vbaword-cli
Version: {VERSION}
Section: utils
Priority: optional
Architecture: amd64
Maintainer: vbaWord project <vbaword@localhost>
Installed-Size: {installed}
Depends: libc6 (>= 2.28), libgcc-s1 | libgcc1
Description: Extract Word legacy fields and content controls
 Native Rust DOCX/DOCM extraction to DOCX and XLSX.
 Python and Office are not required for the native backend.
'''.encode()
        md5=''.join(hashlib.md5(data).hexdigest()+'  '+name+'\n' for name,(data,mode) in sorted(entries.items())).encode()
        deb=args.output/f'Word汇总工具_{VERSION}_amd64.deb'
        deb_file(deb,archive({'control':(control,0o644),'md5sums':(md5,0o644)}),archive(entries));outputs.append(deb)
    (args.output/'SHA256SUMS.txt').write_text(''.join(hashlib.sha256(path.read_bytes()).hexdigest()+'  '+path.name+'\n' for path in outputs),encoding='utf-8')
    print(json.dumps([{'file':str(path),'bytes':path.stat().st_size} for path in outputs],indent=2))


if __name__ == '__main__': main()
