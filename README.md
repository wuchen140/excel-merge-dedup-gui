# Excel 多表合并去重工具

基于 `openpyxl + PySide6` 的桌面工具，用于将一个 `.xlsx` 文件内多个工作表合并并去重。

Pencil 界面草图（先设计后开发）：

![Pencil GUI Wireframe](./gui_wireframe_pencil.png)

## 功能

- 自动读取每个工作表，优先识别表头：`英文 / 日文 / 韩文`
- 若未识别到完整表头，回退为前 3 列
- 支持将 `.xlsx` 文件直接拖入窗口作为输入
- 去重方式：
  - `全部列`（英文+日文+韩文）
  - `仅英文列`
- 跳过全空行
- 可自定义输出工作表名
- 可直接修改输出文件名，也可单独选择输出目录

## 运行源码（开发）

```bash
python3 -m venv .venv
.venv/bin/python -m pip install -r requirements.txt
.venv/bin/python excel_merge_gui.py
```

## 打包（macOS）

推荐（启动更快，`onedir`）：

```bash
.venv/bin/pyinstaller --noconfirm --clean --windowed --name ExcelMergeDedup excel_merge_gui.py
```

兼容单文件（体积集中，但启动较慢）：

```bash
.venv/bin/pyinstaller --noconfirm --clean --onefile --windowed --name ExcelMergeDedup excel_merge_gui.py
```

打包产物（推荐 `onedir`）：

- `/Users/wuchen/Desktop/excel_merge_tool/dist/ExcelMergeDedup.app`

> 首次打开可能需要在系统设置里允许该应用运行（未做开发者证书签名）。
