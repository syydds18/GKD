#!/usr/bin/env python3
"""Reproducible build: dotnet SDK + official Excel-DNA 1.9.0 loader."""
import argparse
import hashlib
import io
import json
import pathlib
import shutil
import subprocess
import urllib.request
import zipfile


VERSION = '1.6.0'
PACKAGE_NAME = f'搞快点-{VERSION}-纯文本问答_本地自定义公式'
PAYLOAD_FILES = (
    'Outline.xll', 'Outline.dna', 'Outline.Addin.dll', 'Outline.Core.dll',
    'ExcelDna.Integration.dll', 'Newtonsoft.Json.dll',
)
# The setup project uses only .NET Framework references; do not glob build output.
ROOT_FILES = (
    '安装与卸载.exe', '安装与卸载.exe.config', '使用说明.md',
    'LICENSE', 'NOTICE',
    'Excel-DNA-LICENSE.txt', 'Newtonsoft.Json-LICENSE.md',
)
PRIVATE_KEY_NAME = 'gkd-workflow.key'  # Installer/settings also accept legacy installations.


def expected_files(private=False):
    files = set(ROOT_FILES)
    files.update(f'payload/{arch}/{name}' for arch in ('x86', 'x64') for name in PAYLOAD_FILES)
    if private:
        files.add(PRIVATE_KEY_NAME)
    return files


def write_manifest(dist, expected):
    actual = {f.relative_to(dist).as_posix() for f in dist.rglob('*') if f.is_file()}
    if actual != expected:
        raise RuntimeError(f'Unexpected package files: extra={sorted(actual - expected)}, missing={sorted(expected - actual)}')
    manifest = {}
    for name in sorted(actual):
        data = (dist / name).read_bytes()
        manifest[name] = {'bytes': len(data), 'sha256': hashlib.sha256(data).hexdigest()}
    (dist / '文件校验.json').write_text(json.dumps(manifest, ensure_ascii=False, indent=2), encoding='utf-8')


def verify_archive(archive, package_name, expected):
    with zipfile.ZipFile(archive) as package:
        bad = package.testzip()
        if bad is not None:
            raise RuntimeError(f'ZIP CRC failed: {bad}')
        prefix = package_name + '/'
        wanted = {prefix + name for name in expected | {'文件校验.json'}}
        names = package.namelist()
        if len(names) != len(wanted) or set(names) != wanted:
            raise RuntimeError('ZIP contents do not match the release allowlist')
        manifest = json.loads(package.read(prefix + '文件校验.json').decode('utf-8'))
        if set(manifest) != expected:
            raise RuntimeError('Checksum manifest does not match the release allowlist')
        for name, entry in manifest.items():
            data = package.read(prefix + name)
            if len(data) != entry['bytes'] or hashlib.sha256(data).hexdigest() != entry['sha256']:
                raise RuntimeError(f'ZIP checksum failed: {name}')


def main():
    parser = argparse.ArgumentParser()
    parser.add_argument('--dotnet', default='dotnet')
    parser.add_argument('--private-test-key-file', type=pathlib.Path)
    args = parser.parse_args()
    root = pathlib.Path(__file__).resolve().parents[1]
    dist = root / 'dist' / (PACKAGE_NAME + ('-私密版' if args.private_test_key_file else ''))
    # Version names contain dots; append the extension instead of replacing a suffix.
    archive = dist.parent / (dist.name + '.zip')
    if dist.exists() or archive.exists():
        raise SystemExit(f'Output already exists; refusing to overwrite: {dist}')
    if args.private_test_key_file and not args.private_test_key_file.is_file():
        raise SystemExit('Private test key file does not exist')
    for project in ('src/Outline.Addin/Outline.Addin.csproj', 'src/Outline.Setup/Outline.Setup.csproj'):
        subprocess.run([args.dotnet, 'build', str(root / project), '-c', 'Release', '--nologo'], check=True)
    url = 'https://api.nuget.org/v3-flatcontainer/exceldna.addin/1.9.0/exceldna.addin.1.9.0.nupkg'
    raw = urllib.request.urlopen(url, timeout=90).read()
    addin = root / 'src/Outline.Addin/bin/Release/net48'
    setup = root / 'src/Outline.Setup/bin/Release/net48/Outline.Setup.exe'
    dist.mkdir(parents=True, exist_ok=False)
    shutil.copy2(setup, dist / '安装与卸载.exe')
    shutil.copy2(pathlib.Path(str(setup) + '.config'), dist / '安装与卸载.exe.config')
    if args.private_test_key_file:
        shutil.copy2(args.private_test_key_file, dist / PRIVATE_KEY_NAME)
    with zipfile.ZipFile(io.BytesIO(raw)) as loaders:
        for arch, loader in (('x86', 'ExcelDna.xll'), ('x64', 'ExcelDna64.xll')):
            target = dist / 'payload' / arch
            target.mkdir(parents=True, exist_ok=True)
            (target / 'Outline.xll').write_bytes(loaders.read('tools/net462/' + loader))
            for name in PAYLOAD_FILES:
                if name.endswith('.dll'):
                    shutil.copy2(addin / name, target / name)
            (target / 'Outline.dna').write_text('<DnaLibrary Name="搞快点" RuntimeVersion="v4.0" Language="C#"><ExternalLibrary Path="Outline.Addin.dll" ExplicitExports="false" LoadFromBytes="true" /></DnaLibrary>', encoding='utf-8')
    shutil.copy2(root / 'docs/发行版使用说明.md', dist / '使用说明.md')
    for name in ('LICENSE', 'NOTICE'):
        shutil.copy2(root / name, dist / name)
    shutil.copy2(root / 'third-party/Excel-DNA-LICENSE.txt', dist / 'Excel-DNA-LICENSE.txt')
    shutil.copy2(root / 'third-party/Newtonsoft.Json-LICENSE.md', dist / 'Newtonsoft.Json-LICENSE.md')
    expected = expected_files(bool(args.private_test_key_file))
    write_manifest(dist, expected)
    with zipfile.ZipFile(archive, 'w', zipfile.ZIP_DEFLATED) as package:
        for name in sorted(expected | {'文件校验.json'}):
            package.write(dist / name, dist.name + '/' + name)
    verify_archive(archive, dist.name, expected)
    print('BUILD PACKAGE:', archive)


if __name__ == '__main__':
    main()
