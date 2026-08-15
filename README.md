# 调拨计算工具

一个基于 Python + Tkinter 的 WMS 库存查询与调拨计算桌面小工具。用户导入调拨 Excel 后，程序逐行调用 WMS 库存查询接口，按 FEFO（先到期先出）规则筛选散支并计算总调拨数，结果可在表格中预览并导出为新的 Excel 文件。

## 主要功能

- 导入调拨计算 Excel，自动识别列名：`SKU`、`Name`、`库存地`、`质量状态`、`箱数`、`warehouseId`
- 内置“导入模板下载”按钮，一键保存模板 `.xlsx`；模板已内嵌进程序，运行时不依赖外部文件
- 逐行调用 WMS 库存查询接口，按 Excel 中的 `warehouseId` 动态查询库存明细
- 按 FEFO 规则计算：整箱记录按失效日期升序累加至需求数量，确定“截止日期”，再筛选失效日期不晚于截止日期的散支
- 汇总“总调拨数 = 需求箱数 × 箱规 + 筛选后散支数量”，含散支的行置顶并标红
- 仓库 ID 校验：单张 Excel 只允许 `WH004004` 或 `WH004013` 其中之一，二者并存时拒绝导入
- 状态标注：无库存、无散支、无符合截止期散支、查询失败、接口未回告箱规等

## 技术栈

- Python 3
- Tkinter / ttk：桌面 GUI
- openpyxl：Excel 读写
- requests：调用 WMS 库存查询 HTTP 接口
- threading：后台线程逐行查询，避免界面卡顿
- PyInstaller：Windows 打包

## 运行方式

```bash
pip install openpyxl requests
python 调拨计算工具.py
```

运行程序只需要 `调拨计算工具.py` 和 `template_data.py` 两个文件，无需携带模板源文件；打包成 exe 后模板也会一并包含。

## 使用流程

1. 点击“导入模板下载”，将内嵌模板保存到本地（模板为原样内嵌，如需空白表头可自行清空数据后另存）
2. 点击“导入Excel”选择调拨 Excel，界面会预览数据并校验仓库 ID
3. 点击“查询库存并计算”，程序逐行查询库存并完成散支筛选与汇总
4. 在“库存明细”和“汇总”页签查看结果，点击“导出Excel”保存带时间戳的结果文件

## Windows 打包

在 Windows 环境中运行 `build_windows.bat`：

1. 脚本会自动安装 PyInstaller、openpyxl、requests
2. 自动查找主 Python 文件并执行 PyInstaller 打包
3. 生成 `dist/TransferTool.exe` 并复制到当前目录

