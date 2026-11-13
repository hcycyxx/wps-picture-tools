# 维护说明

仓库 https://github.com/hcycyxx/wps-picture-tools。使用当前项目版本，先检查 project.json、Git 和 WPS AddIns，不默认旧的安装状态仍有效。

源码在 src/PictureSelection.bas，模块名 FloatPictures，两个回调 ConvertAllPictures/EmbedAllPictures 实际仅选择。使用 SpecialCells 找公式、MergeArea 去重、Union 选择单元格；浮动图枚举 Shape 类型 13/11，以 TopLeftCell 判断占用，Shapes.Range 批量选择。没有合格对象时取消旧图片选择。其他三个按钮由 AmountWords 模块提供。

可移植源码快照在 assets/addin-source。Windows、Python、pywin32、PyInstaller、支持 VBA 的 WPS；先 python scripts/build_addin.py，再 python scripts/build_setup.py。安装器需先打开 WPS；--info 只检查包，--status 只读查看，--install 管理自己的加载项，--uninstall 只停用不删除。新版本用新目录，源码、项目版本、文档和 Release 一致。

校验当前表、合并区域、目标有内容跳过、多图冲突、金额按钮保留、不转换及不保存。以用户实际点击为准，不把编译或安装成功等同于验收；不擅自修改客户原件测试。
