# Word 文档差异比对工具

用于比较两份 `.docx` Word 文档差异，输出 Excel 差异结果。  
支持命令行模式和网页上传模式（本地 `8080`）。

## 功能特性

- 比较两份 Word 文档中的段落与表格内容
- 忽略常见空白差异（提高实际可用性）
  - 连续空白归一化
  - 中文字符间空格归一化（如 `其 他` -> `其他`）
  - 斜杠前后空格归一化（如 `A / B` -> `A/B`）
  - 标点前后多余空格归一化
- 相似度校验：若两份文档整体差异过大，拒绝生成结果（防止拿错文件）
- 输出 Excel（当前仅保留 `修改项` sheet）
- Web 页面支持拖拽上传与点击上传

## 技术栈

- Python 3
- `lxml`：解析 `docx` 内部 XML（`word/document.xml`）
- `difflib`：差异匹配
- `openpyxl`：生成 Excel
- `flask`：Web 服务
- `pyinstaller`：打包为可执行文件

## 核心比对逻辑

1. 将 `.docx` 当作 zip 读取，解析 `word/document.xml`
2. 递归提取：
   - 段落 `w:p`
   - 表格 `w:tbl`
   - 内容控件 `w:sdt` 中的文本
3. 文本归一化（空格/标点规则）
4. 相似度校验（防止两份无关文档被误比对）
5. 用 `SequenceMatcher` 计算差异
6. 写入 Excel `修改项` sheet

## 目录说明

- `compare_contracts.py`：核心比对逻辑（CLI 可直接运行）
- `web_app.py`：网页版服务（默认端口 `8080`）
- `templates/index.html`：网页界面
- `contract_compare.spec`：命令行版打包配置
- `contract_web.spec`：网页版打包配置
- `build_windows.bat`：Windows 一键打包脚本
- `run_web.bat`：Windows 一键启动网页版

## 本地运行（Python）

### 1) 安装依赖

```bash
pip install lxml openpyxl flask
```

### 2) 命令行模式

将两份 `.docx` 放在程序目录（排除 `~$` 临时文件），执行：

```bash
python compare_contracts.py
```

输出示例：

- `文档差异结果.xlsx`

### 3) 网页模式

```bash
python web_app.py
```

打开：<http://127.0.0.1:8080>

- 上传第一份（基准版）和第二份（修改版）文档
- 点击“开始分析”
- 结果保存到页面指定目录（默认程序目录）

## Windows 打包（PyInstaller）

> 建议在 Windows 环境中执行打包，生成 Windows 可执行文件。

### 方式 A：命令行手动打包

```bat
py -3 -m pip install -U pip pyinstaller openpyxl lxml flask
py -3 -m PyInstaller --noconfirm contract_compare.spec
py -3 -m PyInstaller --noconfirm contract_web.spec
```

产物：

- `dist\文档比对.exe`
- `dist\文档比对Web.exe`

### 方式 B：批处理一键打包

```bat
build_windows.bat
```

## 常见问题

### 1) 为什么“关闭浏览器后 8080 还占用”？

因为端口由服务进程占用，不是由浏览器占用。  
关闭运行 `web_app.py` / `文档比对Web.exe` 的进程后端口才会释放。

### 2) 两份文件同名会不会冲突？

不会。Web 上传后使用固定临时名分别保存（第一份/第二份分离）。

### 3) 为什么有时会拒绝生成结果？

触发了相似度校验，系统判定两份文档可能不是同一文档的不同版本。  
请检查上传文件是否正确。