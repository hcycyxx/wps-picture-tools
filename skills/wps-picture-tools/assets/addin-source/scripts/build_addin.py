"""Compile the standalone VBA add-in with an independent installed WPS process."""
from pathlib import Path
import hashlib
import json
import re
import tempfile
import zipfile
import xml.etree.ElementTree as E
import win32com.client

ROOT=Path(__file__).resolve().parents[1]

def build():
    config=json.loads((ROOT/'project.json').read_text('utf-8'))
    dest=ROOT/'dist'/config['version']
    target=dest/config['addin']
    if target.exists():raise RuntimeError('产物已存在，未覆盖。请使用新版本号。')
    original=None
    try:original=win32com.client.GetActiveObject('ket.Application')
    except Exception:pass
    app=win32com.client.DispatchEx('ket.Application')
    if original is not None and app.Hwnd==original.Hwnd:
        raise RuntimeError('未取得独立 WPS 进程，停止以免干扰正在编辑的表格。')
    book=None
    dest.mkdir(parents=True,exist_ok=True)
    names=[]
    try:
        app.Visible=False
        app.EnableEvents=False
        app.DisplayAlerts=False
        book=app.Workbooks.Add()
        for module_name, file_name in [('FloatPictures','PictureSelection.bas'),('AmountWords','AmountWords.bas')]:
            source=(ROOT/'src'/file_name).read_text('utf-8')
            code='\r\n'.join(line for line in source.splitlines() if not line.startswith('Attribute '))
            procedures=re.findall(r'(?im)^\s*(?:Public |Private )?(?:Sub|Function)\s+(\w+)',code)
            module=book.VBProject.VBComponents.Add(1)
            module.Name=module_name
            module.CodeModule.AddFromString(code)
            for procedure in procedures:
                if not module.CodeModule.ProcStartLine(procedure,0):raise RuntimeError('WPS 未识别过程：'+procedure)
            names.extend(procedures)
        book.IsAddin=True
        book.SaveAs(str(target),FileFormat=55)
        book.Close(SaveChanges=False)
        book=None
    finally:
        if book is not None:book.Close(SaveChanges=False)
        app.Quit()
    rel='http://schemas.openxmlformats.org/package/2006/relationships'
    E.register_namespace('',rel)
    with zipfile.ZipFile(target) as z:entries=[(i,z.read(i.filename)) for i in z.infolist()]
    relationships=E.fromstring(dict((i.filename,data) for i,data in entries)['_rels/.rels'])
    E.SubElement(relationships,'{'+rel+'}Relationship',{'Id':'WpsPictureToolsRibbon','Type':'http://schemas.microsoft.com/office/2006/relationships/ui/extensibility','Target':'customUI/customUI.xml'})
    with tempfile.NamedTemporaryFile(dir=dest,suffix='.xlam',delete=False) as f:patch=Path(f.name)
    with zipfile.ZipFile(patch,'w',zipfile.ZIP_DEFLATED) as z:
        for info,data in entries:
            if info.filename=='_rels/.rels':data=E.tostring(relationships,encoding='utf-8',xml_declaration=True)
            z.writestr(info,data)
        z.writestr('customUI/customUI.xml',(ROOT/'src'/'ribbon.xml').read_bytes())
    patch.replace(target)
    manifest={'project':config['name'],'version':config['version'],'addin':config['addin'],'sha256':hashlib.sha256(target.read_bytes()).hexdigest(),'vba_procedures':names,'source_suite_version':config['source_suite_version']}
    (dest/'manifest.json').write_text(json.dumps(manifest,ensure_ascii=False,indent=2),'utf-8')
    print(json.dumps(manifest,ensure_ascii=True))

if __name__=='__main__':build()
