"""Publish prepared, password-free release assets; never reads credentials into logs."""
import argparse
import hashlib
import io
import json
from pathlib import Path
import subprocess
import tarfile
import urllib.parse
import urllib.request
import zipfile


def clean_config(data):
    import openpyxl
    book=openpyxl.load_workbook(io.BytesIO(data),data_only=False)
    try:
        assert set(book.sheetnames)=={'config','mapping'}
        assert book['config']['B2'].value in (None,'')
        assert book['mapping']['E1'].value in (None,'')
        for row in book['config'].iter_rows():
            label=str(row[0].value or '').lower()
            if '密码' in label or 'password' in label:assert row[1].value in (None,'')
    finally:book.close()


def main():
    parser=argparse.ArgumentParser();parser.add_argument('--directory',type=Path,required=True);parser.add_argument('--tag',required=True)
    parser.add_argument('--repo',default='Azi-lzb/vba_word_collectTools');parser.add_argument('--notes',type=Path,required=True);args=parser.parse_args()
    folder=args.directory.resolve();manifest=folder/'SHA256SUMS.txt'
    assets=[]
    for line in manifest.read_text(encoding='utf8').splitlines():
        digest,name=line.split('  ',1);path=folder/name
        assert path.parent==folder and path.is_file()
        assert hashlib.sha256(path.read_bytes()).hexdigest()==digest
        assets.append(path)
    assert any(p.suffix=='.exe' for p in assets) and any(p.suffix=='.deb' for p in assets)
    for path in assets:
        if path.name=='config.xlsx':clean_config(path.read_bytes())
        elif path.suffix=='.zip':
            with zipfile.ZipFile(path) as z:
                assert set(z.namelist())=={'Word汇总工具.exe','config.xlsx'}
                clean_config(z.read('config.xlsx'))
        elif path.name.endswith('.tar.gz'):
            with tarfile.open(path) as tar:
                configurations=[m for m in tar.getmembers() if m.name.endswith('/config.xlsx')]
                assert len(configurations)==1
                clean_config(tar.extractfile(configurations[0]).read())
    assets.append(manifest)
    cred_result=subprocess.run(['git','credential','fill'],input='protocol=https\nhost=github.com\n\n',capture_output=True,text=True,check=True)
    credentials=dict(line.split('=',1) for line in cred_result.stdout.splitlines() if '=' in line)
    token=credentials['password']
    def api(url,method='GET',payload=None,raw=None,content_type='application/json'):
        body=json.dumps(payload).encode() if payload is not None else raw
        request=urllib.request.Request(url,data=body,method=method,headers={'Authorization':'Bearer '+token,'Accept':'application/vnd.github+json','Content-Type':content_type,'User-Agent':'Word-summary-release'})
        with urllib.request.urlopen(request,timeout=180) as response:return json.load(response)
    base='https://api.github.com/repos/'+args.repo
    existing=api(base+'/releases')
    assert not any(r['tag_name']==args.tag for r in existing),'Tag already has a release; choose a new tag'
    commit=subprocess.check_output(['git','rev-parse','HEAD'],text=True).strip()
    release=api(base+'/releases','POST',{'tag_name':args.tag,'target_commitish':commit,'name':'Word汇总工具 '+args.tag,'body':args.notes.read_text(encoding='utf8'),'draft':True,'prerelease':False})
    uploaded=[]
    for path in assets:
        remote_name=path.name.replace('Word汇总工具','WordSummaryTool')
        content=path.read_bytes()
        if path==manifest:content=content.decode('utf8').replace('Word汇总工具','WordSummaryTool').encode('utf8')
        digest='sha256:'+hashlib.sha256(content).hexdigest()
        url=release['upload_url'].split('{')[0]+'?'+urllib.parse.urlencode({'name':remote_name,'label':path.name})
        result=api(url,'POST',raw=content,content_type='application/octet-stream')
        assert result['size']==len(content)
        assert result['name']==remote_name
        if result.get('digest'):assert result['digest']==digest
        uploaded.append({'name':result['name'],'local_name':path.name,'size':len(content),'sha256':digest,'url':result['browser_download_url']})
        print(json.dumps({'uploaded':path.name,'bytes':len(content)}),flush=True)
    # Publication happens only after every upload and checksum check succeeds.
    published=api(base+'/releases/'+str(release['id']),'PATCH',{'draft':False,'make_latest':'true'})
    assert not published['draft']
    proof={'release_url':published['html_url'],'tag':args.tag,'commit':commit,'assets':uploaded,'password_cells_blank':True}
    (folder/'release-published.json').write_text(json.dumps(proof,ensure_ascii=False,indent=2),encoding='utf8')
    print(json.dumps({'published':proof['release_url'],'assets':len(uploaded)}),flush=True)


if __name__=='__main__':main()
