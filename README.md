# 往来重分类分析工具

读取数据源 Excel（供应商应付类、客户合同负债类表格），按重分类规则提取数据、汇总分析并生成新文件，供往来重分类分析使用。

## 功能

1. 从各数据源表提取「总账科目：长文本 / 名称 / 公司代码货币价值」，导出明细表；
2. 生成汇总表：按「总账科目：长文本 + 名称」汇总金额，删除合计为零的行，排序；
3. 汇总表在公司代码货币价值后增加「重分类科目」列，依据规则文件判断；
4. 「应付含暂估」：按供应商汇总 应付款/暂估款/合计，并按合计方向重分类；
5. 「合同负债重分类其他流动负债」：汇总表合同负债负数行按客户提取，计算绝对值/可开票税额/已开票税额/差额/科目余额，给出差额正负合计与重分类金额结论；
6. 提供重分类规则的导入（`--rules`）与导出模板（`--export-template`）功能。

## 文件结构

- `reclass_core.py` — 核心逻辑（数据提取、汇总、重分类、Excel 生成）
- `reclass_entry.py` — 统一入口：无参数启动 tkinter GUI；带参数走 CLI
- `reclass_gui_macos.spec` — macOS GUI 打包参考配置（PyInstaller）
- `reclass_cli.spec` — CLI 打包参考配置（PyInstaller）
- `requirements.txt` — 运行依赖

## 使用方式

### 图形界面（GUI）

双击应用（macOS `.app` / Windows `.exe`），或直接运行：

```bash
python reclass_entry.py
```

支持拖拽导入数据源 / 规则文件，规则记忆自动加载，执行分析。

### 命令行（CLI）

```bash
python reclass_entry.py --data-dir <数据源目录> [--rules <规则文件>] [--out <输出路径>]
python reclass_entry.py --export-template <模板路径>
```

- `--data-dir` 数据源目录（默认当前目录）
- `--rules` 重分类规则文件（留空则自动加载记忆中的生效规则）
- `--out` 结果输出路径（默认 `<data-dir>/往来重分类分析结果.xlsx`）
- `--export-template` 导出重分类规则模板

## Windows 打包（GitHub Actions 交叉编译）

仓库自带 `.github/workflows/build-windows.yml`：使用 `windows-latest` runner 安装 Python + PyInstaller，分别打出 GUI 版（`--windowed --onefile`，含 tkinterdnd2）与 CLI 版（`--console --onefile`）exe，作为 Actions Artifact 提供下载。

本地打包命令（Windows）：

```bash
pip install -r requirements.txt pyinstaller
pyinstaller --onefile --windowed --name "往来重分类分析工具" --hidden-import tkinterdnd2 reclass_entry.py
pyinstaller --onefile --console --name "往来重分类分析工具-cli" --exclude-module tkinterdnd2 reclass_entry.py
```
