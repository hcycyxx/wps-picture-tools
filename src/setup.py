"""Installer UI/CLI. It changes only this standalone add-in registration."""
from pathlib import Path
from datetime import datetime
import argparse
import hashlib
import json
import os
import shutil
import sys

ROOT=Path(getattr(sys,'_MEIPASS',Path(__file__).resolve().parents[1]))
CONFIG=json.loads((ROOT/'project.json').read_text('utf-8'))
PREFIX='KynkoPictureTools_'

def payload():
    return ROOT/'payload' if getattr(sys,'frozen',False) else ROOT/'dist'/CONFIG['version']

def package_info():
    folder=payload()
    manifest=json.loads((folder/'manifest.json').read_text('utf-8'))
    addin=folder/CONFIG['addin']
    if manifest['version']!=CONFIG['version'] or manifest['addin']!=CONFIG['addin']:
        raise RuntimeError('安装包版本信息不一致。')
    if hashlib.sha256(addin.read_bytes()).hexdigest()!=manifest['sha256']:
        raise RuntimeError('加载项文件校验失败，请重新获取完整安装包。')
    return manifest

def active_wps():
    import win32com.client
    try:return win32com.client.GetActiveObject('ket.Application')
    except Exception as e:raise RuntimeError('请先打开 WPS 表格，再使用安装或停用功能。') from e

def own_addins(app):
    return [app.AddIns.Item(i) for i in range(1,app.AddIns.Count+1) if str(app.AddIns.Item(i).Name).startswith((PREFIX,'KynkoPictureFloat')) and Path(str(app.AddIns.Item(i).FullName)).parent.parent == Path(os.environ['LOCALAPPDATA'])/'WpsPictureFloat']

def install():
    manifest=package_info()
    app=active_wps()
    previous=[a for a in own_addins(app) if a.Installed]
    for a in previous:
        previous_version=Path(str(a.FullName)).parent.name
        try:newer=tuple(map(int,previous_version.split('-')[0].split('.')))>tuple(map(int,CONFIG['version'].split('.')))
        except ValueError:raise RuntimeError('检测到未识别的独立工具版本，停止以免降级。')
        if newer:raise RuntimeError('当前已启用更新版本，未降级。')
    dest=Path(os.environ['LOCALAPPDATA'])/'WpsPictureFloat'/CONFIG['version']
    dest.mkdir(parents=True,exist_ok=True)
    for name in (CONFIG['addin'],'manifest.json'):
        source=payload()/name
        target=dest/name
        if target.exists():
            if hashlib.sha256(target.read_bytes()).digest()!=hashlib.sha256(source.read_bytes()).digest():
                raise RuntimeError('目标目录存在不同内容的同版本文件，未覆盖：'+str(target))
        else:shutil.copy2(source,target)
    disabled=[]
    new=None
    try:
        for a in previous:
            if os.path.normcase(str(a.FullName))!=os.path.normcase(str(dest/CONFIG['addin'])):
                a.Installed=False
                disabled.append(a)
        new=app.AddIns.Add(str(dest/CONFIG['addin']),False)
        new.Installed=True
        if not new.Installed:raise RuntimeError('WPS 未能启用加载项，请检查 VBA 支持与 WPS 的加载项提示。')
    except Exception:
        if new is not None:
            try:new.Installed=False
            except Exception:pass
        for a in disabled:a.Installed=True
        raise
    state={'project':CONFIG['name'],'version':CONFIG['version'],'path':str(dest/CONFIG['addin']),'installed':True,'time':datetime.now().isoformat(),'previous_own_addins':[str(a.FullName) for a in previous],'payload_sha256':manifest['sha256']}
    history=dest/('install-state-'+datetime.now().strftime('%Y%m%d-%H%M%S-%f')+'.json')
    history.write_text(json.dumps(state,ensure_ascii=False,indent=2),'utf-8')
    return '已启用！请在 WPS 顶部“图片工具”选择图片，再由你右键转换。'

def uninstall():
    app=active_wps()
    count=0
    for a in own_addins(app):
        if a.Installed:a.Installed=False;count+=1
    return f'已停用 {count} 个图片工具加载项。安装文件和其他套件均保留。'

def status():
    app=active_wps()
    return [{'name':str(a.Name),'path':str(a.FullName),'installed':bool(a.Installed)} for a in own_addins(app)]

def gui():
    import tkinter as tk
    from tkinter import ttk,messagebox
    window=tk.Tk()
    window.title('WPS 图片选择工具 · '+CONFIG['version'])
    window.geometry('580x330')
    window.resizable(False,False)
    panel=ttk.Frame(window,padding=24)
    panel.pack(fill='both',expand=True)
    ttk.Label(panel,text='WPS 图片选择工具',font=('Microsoft YaHei UI',19,'bold')).pack(anchor='w')
    ttk.Label(panel,text='只选择图片 · 目标有内容自动跳过',font=('Microsoft YaHei UI',11)).pack(anchor='w',pady=(8,16))
    ttk.Label(panel,text='先打开支持 VBA 的 WPS 表格，再点击安装。\n安装后：点选择按钮 → 右键转换 → 检查后保存。\n工具只改变当前工作表选区，不转换或保存表格。',font=('Microsoft YaHei UI',10),justify='left').pack(anchor='w')
    actions=ttk.Frame(panel)
    actions.pack(anchor='w',pady=18)
    progress=tk.StringVar(value='保留三个金额按钮；旧版图片加载项仅停用，文件保留。')
    def run_action(action):
        install_button.state(['disabled']);remove_button.state(['disabled'])
        try:
            progress.set('正在处理，请稍候……');window.update_idletasks()
            result=action();progress.set(result);messagebox.showinfo('WPS 图片选择工具',result,parent=window)
        except Exception as e:
            progress.set('未完成，请查看提示。');messagebox.showerror('WPS 图片选择工具',str(e),parent=window)
        finally:install_button.state(['!disabled']);remove_button.state(['!disabled'])
    install_button=ttk.Button(actions,text='安装到 WPS',command=lambda:run_action(install))
    install_button.pack(side='left',padx=(0,12),ipadx=12,ipady=4)
    remove_button=ttk.Button(actions,text='停用图片加载项',command=lambda:run_action(uninstall))
    remove_button.pack(side='left',ipadx=8,ipady=4)
    ttk.Label(panel,textvariable=progress,wraplength=520,font=('Microsoft YaHei UI',9)).pack(anchor='w')
    window.mainloop()

def main():
    parser=argparse.ArgumentParser(description='WPS 图片选择工具安装工具')
    actions=parser.add_mutually_exclusive_group()
    actions.add_argument('--info',action='store_true',help='只校验安装包并显示版本，不访问 WPS')
    actions.add_argument('--status',action='store_true',help='只读显示图片工具加载项状态')
    actions.add_argument('--install',action='store_true',help='安装/启用图片工具加载项')
    actions.add_argument('--uninstall',action='store_true',help='仅停用图片工具加载项，不删文件')
    parser.add_argument('--report',type=Path,help='将命令结果另存为 JSON（便于窗口版 EXE 的只读诊断）')
    args=parser.parse_args()
    if args.info:result=package_info()
    elif args.status:result=status()
    elif args.install:result=install()
    elif args.uninstall:result=uninstall()
    else:return gui()
    if args.report:
        if args.report.exists():raise RuntimeError('报告路径已存在，未覆盖。')
        args.report.write_text(json.dumps(result,ensure_ascii=False,indent=2),'utf-8')
    print(json.dumps(result,ensure_ascii=True,indent=2))

if __name__=='__main__':main()
