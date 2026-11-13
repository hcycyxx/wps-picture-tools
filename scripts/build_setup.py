"""Package the installer; build_addin.py must run first."""
from pathlib import Path
import json
import subprocess
import sys

ROOT=Path(__file__).resolve().parents[1]
config=json.loads((ROOT/'project.json').read_text('utf-8'))
version=config['version']
name='WpsPictureTools-Setup-v'+version
dest=ROOT/'dist'/version
if (dest/(name+'.exe')).exists():raise RuntimeError('EXE 已存在，未覆盖。')
if not (dest/config['addin']).is_file():raise RuntimeError('请先构建 xlam。')
subprocess.run([sys.executable,'-m','PyInstaller','--noconfirm','--clean','--onefile','--windowed','--name',name,'--distpath',str(dest),'--workpath',str(ROOT/'build'/version),'--specpath',str(ROOT/'build'/version),'--hidden-import','win32timezone','--add-data',str(ROOT/'project.json')+':.','--add-data',str(dest/config['addin'])+':payload','--add-data',str(dest/'manifest.json')+':payload',str(ROOT/'src'/'setup.py')],cwd=ROOT,check=True)
print(dest/(name+'.exe'))
