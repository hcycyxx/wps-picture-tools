"""Create immutable source, skill and manual-install archives plus checksums."""
from pathlib import Path
import hashlib,json,subprocess,zipfile
ROOT=Path(__file__).resolve().parents[1]
cfg=json.loads((ROOT/'project.json').read_text('utf-8'))
version=cfg['version'];dest=ROOT/'dist'/version
names={
 'source':dest/f'WpsPictureTools-SourceDocs-v{version}.zip',
 'skill':dest/'wps-picture-tools-skill-v1.2.0.zip',
 'addin':dest/f'WpsPictureTools-XLAM-v{version}.zip'}
if any(p.exists() for p in names.values()) or (dest/'SHA256SUMS.txt').exists():raise RuntimeError('发布附件已存在，不覆盖。')
files=subprocess.check_output(['git','ls-files','-z'],cwd=ROOT).decode('utf-8').split('\0')
with zipfile.ZipFile(names['source'],'w',zipfile.ZIP_DEFLATED) as z:
 for name in files:
  if name:z.write(ROOT/name,arcname=f'wps-picture-tools-{version}/'+name)
with zipfile.ZipFile(names['skill'],'w',zipfile.ZIP_DEFLATED) as z:
 for p in sorted((ROOT/'skills/wps-picture-tools').rglob('*')):
  if p.is_file() and '__pycache__' not in p.parts:z.write(p,arcname=str(p.relative_to(ROOT/'skills')))
with zipfile.ZipFile(names['addin'],'w',zipfile.ZIP_DEFLATED) as z:
 for p in [dest/cfg['addin'],dest/'manifest.json',ROOT/'README.md',ROOT/'LICENSE']:z.write(p,arcname=p.name)
artifacts=[dest/f'WpsPictureTools-Setup-v{version}.exe',*names.values()]
(dest/'SHA256SUMS.txt').write_text(''.join(hashlib.sha256(p.read_bytes()).hexdigest()+'  '+p.name+'\n' for p in artifacts),'utf-8')
print(json.dumps([{'name':p.name,'bytes':p.stat().st_size} for p in artifacts],indent=2))
