#!/usr/bin/env python3
"""Reproducible build: dotnet SDK + official Excel-DNA 1.9.0 loader."""
import argparse, pathlib, subprocess, urllib.request, zipfile, io, shutil, json, hashlib
p=argparse.ArgumentParser();p.add_argument('--dotnet',default='dotnet');args=p.parse_args()
root=pathlib.Path(__file__).resolve().parents[1];dist=root/'dist'/'搞快点-1.4.0-品牌升级_安装器增强_水电提项目_智能求和'
if dist.exists():raise SystemExit(f'Output directory already exists; refusing to overwrite: {dist}')
for proj in ['src/Outline.Addin/Outline.Addin.csproj','src/Outline.Setup/Outline.Setup.csproj']:
 subprocess.run([args.dotnet,'build',str(root/proj),'-c','Release','--nologo'],check=True)
dist.mkdir(parents=True,exist_ok=True)
url='https://api.nuget.org/v3-flatcontainer/exceldna.addin/1.9.0/exceldna.addin.1.9.0.nupkg'
raw=urllib.request.urlopen(url,timeout=90).read();z=zipfile.ZipFile(io.BytesIO(raw))
addin=root/'src/Outline.Addin/bin/Release/net48';setup=root/'src/Outline.Setup/bin/Release/net48/Outline.Setup.exe'
assert setup.exists(),setup
shutil.copy2(setup,dist/'安装与卸载.exe')
shutil.copy2(pathlib.Path(str(setup)+'.config'),dist/'安装与卸载.exe.config')
# Keep the renamed executable and its matching runtime config together.
for f in setup.parent.glob('*.dll'):shutil.copy2(f,dist/f.name)
manifest={}
for arch,loader in [('x86','ExcelDna.xll'),('x64','ExcelDna64.xll')]:
 d=dist/'payload'/arch;d.mkdir(parents=True,exist_ok=True)
 (d/'Outline.xll').write_bytes(z.read('tools/net462/'+loader))
 for name in ['Outline.Addin.dll','Outline.Core.dll','ExcelDna.Integration.dll']:
  source=addin/name;assert source.exists(),source;shutil.copy2(source,d/name)
 (d/'Outline.dna').write_text('<DnaLibrary Name="搞快点" RuntimeVersion="v4.0" Language="C#"><ExternalLibrary Path="Outline.Addin.dll" ExplicitExports="false" LoadFromBytes="true" /></DnaLibrary>',encoding='utf-8')
shutil.copy2(root/'README.md',dist/'使用说明.md')
shutil.copy2(root/'验证记录.md',dist/'验证记录.md')
shutil.copytree(root/'samples',dist/'自检工作簿')
shutil.copy2(root/'installer/README.md',dist/'安装技术说明.md')
shutil.copy2(root/'third-party/Excel-DNA-LICENSE.txt',dist/'Excel-DNA-LICENSE.txt')
for f in dist.rglob('*'):
 if f.is_file():manifest[str(f.relative_to(dist))]={'bytes':f.stat().st_size,'sha256':hashlib.sha256(f.read_bytes()).hexdigest()}
(dist/'文件校验.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),encoding='utf-8')
archive=root/'dist'/(dist.name+'.zip')
with zipfile.ZipFile(archive,'w',zipfile.ZIP_DEFLATED) as out:
 for f in dist.rglob('*'):
  if f.is_file():out.write(f,str(pathlib.Path(dist.name)/f.relative_to(dist)))
with zipfile.ZipFile(archive) as out:assert out.testzip() is None
print('BUILD PACKAGE:',archive)
