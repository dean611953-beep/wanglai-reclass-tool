#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
往来重分类分析工具
====================================
读取数据源 Excel（供应商应付类、客户合同负债类表格），按重分类规则提取数据、
汇总分析并生成新文件，供后续往来重分类分析使用。

功能：
  1) 从各数据源表提取「总账科目：长文本 / 名称 / 公司代码货币价值」，导出明细表；
  2) 生成汇总表：按「总账科目：长文本 + 名称」汇总金额，删除合计为零的行，排序；
  3) 汇总表在公司代码货币价值后增加「重分类科目」列，依据规则文件判断；
  4) 新增「应付含暂估」：按供应商汇总 应付款/暂估款/合计，并按合计方向重分类；
  5) 新增「合同负债重分类其他流动负债」：汇总表合同负债负数行按客户提取，计算
      绝对值 / 可开票税额（国内 1.13 税率、国外 0）/ 已开票税额 / 差额 / 科目余额，
      并给出差额正负合计与重分类金额结论；
6) 提供重分类规则的导入（--rules）与导出模板（--export-template）功能。

用法示例：
  # 常规分析（数据源目录 + 规则文件 + 输出文件）
  python 往来重分类分析工具.py \
      --data-dir "/Users/dean/Desktop/G/新需求" \
      --rules  "/Users/dean/Desktop/G/新需求/重分类规则.xlsx" \
      --out    "/Users/dean/.../output/往来重分类分析结果.xlsx"

  # 导出重分类规则模板（便于后续更新规则）
  python 往来重分类分析工具.py --export-template "/path/to/重分类规则模板.xlsx"

字段映射说明（“名称”统一列）：
  - 供应商应付类的余额.xlsx ：名称 = 「供应商科目：名称 1」
  - 供应商应付暂估.xlsx     ：名称 = 「供应商科目：名称 1」（回退「名称」列）
  - 客户合同负债类.xlsx     ：名称 = 「名称」
"""

import argparse
import hashlib
import json
import os
import shutil
import sys
from collections import defaultdict
from datetime import datetime

from openpyxl import load_workbook, Workbook
from openpyxl.styles import Font, Alignment, PatternFill
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------------------------
# 常量与数据源配置
# ---------------------------------------------------------------------------

# 数据源表配置：file=文件名；subject_col=总账科目长文本；name_col=统一名称来源；
# fallback_name_col=名称回退列（可空）；amount_col=金额列；direction_col=借/贷标识列（可空）
SOURCE_CONFIG = [
    {
        "file": "供应商应付类的余额.xlsx",
        "source": "供应商应付类的余额",
        "subject_col": "总账科目：长文本",
        "name_col": "供应商科目：名称 1",
        "fallback_name_col": None,
        "amount_col": "公司代码货币价值",
        "direction_col": "借/贷标识",
    },
    {
        "file": "供应商应付暂估.xlsx",
        "source": "供应商应付暂估",
        "subject_col": "总账科目：长文本",
        "name_col": "供应商科目：名称 1",
        "fallback_name_col": "名称",   # 该表表头存在多个“名称”列，取第一个非空出现者
        "amount_col": "公司代码货币价值",
        "direction_col": "借/贷标识",
    },
    {
        "file": "客户合同负债类.xlsx",
        "source": "客户合同负债类",
        "subject_col": "总账科目：长文本",
        "name_col": "名称",
        "fallback_name_col": None,
        "amount_col": "公司代码货币价值",
        "direction_col": "借/贷标识",
    },
]

# 数据源类型（与 SOURCE_CONFIG 下标对应）
SOURCE_TYPES = ["balance", "estimate", "customer"]
SOURCE_TYPE_INDEX = {"balance": 0, "estimate": 1, "customer": 2}
STANDARD_FILE_TO_TYPE = {cfg["file"]: t for t, cfg in zip(SOURCE_TYPES, SOURCE_CONFIG)}

# 文件名关键词（用于目录内“内容同类”识别；兼容 630 / 0630 等日期变体）
FILENAME_HINTS = {
    "balance": ["余额", "balance"],
    "estimate": ["暂估", "estimate"],
    "customer": ["客户", "合同负债", "customer"],
}

# 目录内扫描时排除的无关 xlsx（结果 / 规则 / 模板 / 记忆 / Excel 锁文件）
IGNORED_FILENAME_MARKERS = ["~$", "结果", "规则", "模板", "记忆"]

RULES_HEADER = ["总账科目：长文本", "公司代码货币价值", "余额方向", "重分类科目"]

# 规则永久记忆：在软件自身固定存储目录保存“当前生效规则”副本 + 元信息记录，
# 不依赖数据源目录（数据源目录中出现任何规则文件均不影响当前生效规则）。
MEMORY_RULES_FILENAME = "当前生效重分类规则.xlsx"
MEMORY_META_FILENAME = "重分类规则记忆.json"


def get_memory_dir():
    """软件自身固定存储目录（跨平台）：
       macOS: ~/Library/Application Support/往来重分类分析工具/
       Windows: %APPDATA%/往来重分类分析工具/
    不依赖数据源目录，规则导入后固化于此，不重新导入规则保持不变。"""
    if sys.platform == "darwin":
        base = os.path.expanduser("~/Library/Application Support")
    elif os.name == "nt":
        base = os.environ.get("APPDATA") or os.path.expanduser("~")
    else:
        base = os.path.expanduser("~/.config")
    d = os.path.join(base, "往来重分类分析工具")
    os.makedirs(d, exist_ok=True)
    return d


# ---------------------------------------------------------------------------
# 基础工具函数
# ---------------------------------------------------------------------------

def to_float(v):
    """兼容 int / float / 千分位字符串 的金额转换。"""
    if v is None:
        return 0.0
    if isinstance(v, (int, float)):
        return float(v)
    s = str(v).strip().replace(",", "").replace("，", "")
    if s in ("", "-", "--", "None"):
        return 0.0
    try:
        return float(s)
    except ValueError:
        return 0.0


def find_header_indexes(header, wanted_cols):
    """在表头行中定位目标列名 -> {wanted: 列索引}；重复列名取第一次出现。"""
    col_index = {}
    for i, cell in enumerate(header):
        if cell is None:
            continue
        name = str(cell).strip()
        if name in wanted_cols and name not in col_index:
            col_index[name] = i
    return col_index


def build_header_map(header):
    """构建 列名 -> 索引 映射（重复列名保留最后一次出现，用于多‘名称’列场景）。"""
    mapping = {}
    for i, cell in enumerate(header):
        if cell is None:
            continue
        mapping[str(cell).strip()] = i
    return mapping


def _read_header_cells(path):
    """轻量读取 xlsx 首行表头单元格列表；失败返回 None。"""
    try:
        wb = load_workbook(path, read_only=True, data_only=True)
        try:
            ws = wb.worksheets[0]
            for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
                return list(row)
            return []
        finally:
            wb.close()
    except Exception:
        return None


def classify_source_file(path):
    """按“内容同类”识别数据源文件所属类型（文件名可不同于标准名）。

    判定依据（列结构特征优先，文件名关键词辅助）：
      - 三类数据源均含「总账科目：长文本」「公司代码货币价值」；
      - 余额类：另有「供应商科目：名称 1」，无「名称」回退列（21 列结构）；
      - 暂估类：另有「供应商科目：名称 1」且有「名称」回退列（94 列结构）；
      - 客户类：无「供应商科目：名称 1」，有「名称」列（22 列结构）。
    文件名关键词（余额/暂估/客户，兼容 630 / 0630 等变体）用于辅助说明；
    文件名与结构冲突时以列结构为准（列结构决定字段映射）。

    返回 (source_type, cfg) 或 None。
    """
    base = os.path.basename(path)
    if any(m in base for m in IGNORED_FILENAME_MARKERS):
        return None
    header = _read_header_cells(path)
    if not header:
        return None
    hmap = build_header_map(header)
    has_subj = "总账科目：长文本" in hmap
    has_amount = "公司代码货币价值" in hmap
    has_supplier = "供应商科目：名称 1" in hmap
    has_name = "名称" in hmap
    if not (has_subj and has_amount):
        return None
    if has_supplier and has_name:
        stype = "estimate"
    elif has_supplier:
        stype = "balance"
    elif has_name:
        stype = "customer"
    else:
        return None
    return stype, SOURCE_CONFIG[SOURCE_TYPE_INDEX[stype]]


def resolve_source_paths(data_dir, explicit_files=None):
    """确定三个数据源类型的实际文件路径（内容同类识别）。

    优先级：
      1) 显式指定文件（explicit_files：dict {标准文件名/类型: 路径} 或 路径列表）——
         按内容分类归入对应类型；dict 的 key 若已是 SOURCE_TYPES 类型名（如
         "balance"/"estimate"/"customer"），视为调用方（GUI 列表加载时）已缓存的
         分类结果，直接复用、不再读表头；
      2) 目录内扫描识别（列结构 + 文件名关键词），同类型多候选时标准文件名优先；
      3) 均未找到 -> 该类型为 None（调用方据此报错）。

    返回 {source_type: path_or_None}。
    """
    explicit = {}
    if explicit_files:
        if isinstance(explicit_files, dict):
            for key, p in explicit_files.items():
                if not p or not os.path.isfile(p):
                    continue
                if key in SOURCE_TYPES:
                    # 已分类结果（GUI 列表加载时缓存）：直接复用，不再读表头
                    explicit[key] = p
                    continue
                cls = classify_source_file(p)
                if cls:
                    explicit[cls[0]] = p
                elif key in STANDARD_FILE_TO_TYPE:
                    # 兼容旧调用：{标准数据源文件名: 实际路径}
                    explicit[STANDARD_FILE_TO_TYPE[key]] = p
        else:
            for p in explicit_files:
                if not p or not os.path.isfile(p):
                    continue
                cls = classify_source_file(p)
                if cls:
                    explicit[cls[0]] = p
    dir_found = {}
    if data_dir and os.path.isdir(data_dir):
        for fn in sorted(os.listdir(data_dir)):
            if not fn.lower().endswith(".xlsx"):
                continue
            if any(m in fn for m in IGNORED_FILENAME_MARKERS):
                continue
            p = os.path.join(data_dir, fn)
            cls = classify_source_file(p)
            if not cls:
                continue
            stype, cfg = cls
            if stype in dir_found:
                # 同类型多候选：标准文件名优先于其它变体
                if fn == cfg["file"] and os.path.basename(dir_found[stype]) != cfg["file"]:
                    dir_found[stype] = p
                continue
            dir_found[stype] = p
    return {t: (explicit.get(t) or dir_found.get(t)) for t in SOURCE_TYPES}


# ---------------------------------------------------------------------------
# 重分类规则：导入 / 模板 / 匹配
# ---------------------------------------------------------------------------

def load_rules(rules_path):
    """导入重分类规则文件，返回规则列表 [{科目, 符号, 方向, 重分类科目}, ...]。
    支持附件原始规则表（表头前有空占位列）与按模板导出的标准格式。"""
    wb = load_workbook(rules_path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    header = None
    header_idx = None
    for row in ws.iter_rows(values_only=True):
        if row:
            joined = "".join(str(c) for c in row if c is not None)
            if "总账科目" in joined and "重分类科目" in joined:
                header = list(row)
                header_idx = {str(c).strip(): i for i, c in enumerate(header) if c is not None}
                break
    if header is None:
        raise ValueError(f"规则文件 {rules_path} 未找到包含“总账科目”与“重分类科目”的表头行")

    subj_i = next((i for k, i in header_idx.items() if "总账科目" in k), None)
    sign_i = next((i for k, i in header_idx.items() if "公司代码货币价值" in k), None)
    dir_i = next((i for k, i in header_idx.items() if "余额方向" in k), None)
    reclass_i = next((i for k, i in header_idx.items() if "重分类科目" in k), None)
    if subj_i is None or sign_i is None or reclass_i is None:
        raise ValueError(f"规则文件 {rules_path} 表头缺少必需列：总账科目：长文本 / 公司代码货币价值 / 重分类科目")

    rules = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        subj = str(row[subj_i]).strip() if row[subj_i] is not None else ""
        reclass = str(row[reclass_i]).strip() if row[reclass_i] is not None else ""
        if not subj or not reclass:
            continue
        sign = str(row[sign_i]).strip() if row[sign_i] is not None else ""
        direction = str(row[dir_i]).strip() if dir_i is not None and row[dir_i] is not None else ""
        if sign not in ("正数", "负数"):
            continue
        rules.append({
            "科目": subj,
            "符号": sign,          # 正数 / 负数
            "方向": direction,     # 借方 / 贷方（辅助信息）
            "重分类科目": reclass,
        })
    wb.close()
    if not rules:
        raise ValueError(f"规则文件 {rules_path} 未解析到有效规则行")
    return rules


# 汇总表重分类排除科目：完整科目名精确匹配，命中则不做重分类（重分类科目列留空）
NON_RECLASS_SUBJECTS = ("合同负债-待确认收入", "长期应收款-未确认融资费用")


def match_reclass(subject, amount, rules):
    """依据规则匹配重分类科目：
    1) 总账科目长文本包含规则关键词；
    2) 金额符号与规则“公司代码货币价值（正数/负数）”一致；
    3) 多条命中时取关键词最长者（更精确，如“其他应付款”优先于“应付款”）。
    金额为 0 或未命中返回空字符串（不做重分类）。
    完整科目名命中 NON_RECLASS_SUBJECTS 时同样返回空字符串（精确排除，避免误伤其它科目）。"""
    if amount == 0 or not subject:
        return ""
    if subject.strip() in NON_RECLASS_SUBJECTS:
        return ""
    sign = "正数" if amount > 0 else "负数"
    candidates = []
    for r in rules:
        kw = r["科目"]
        if kw and kw in subject and r["符号"] == sign:
            candidates.append((len(kw), r["重分类科目"]))
    if not candidates:
        return ""
    candidates.sort(key=lambda x: x[0], reverse=True)
    return candidates[0][1]


def match_sheet3_reclass(payable, est, total, rules):
    """sheet3 供应商行「重分类科目」（按合计金额方向匹配）：
    1) 该表仅汇总「应付账款」及其暂估变体科目，故判定科目域固定为「应付账款」；
    2) 符号按「合计」金额方向：合计>0 为正数（借方余额），合计<0 为负数（贷方余额）；
    3) 匹配规则文件中「应付账款」科目的对应方向行：
        正数 -> 预付款项（规则命中，已付/预付性质）；
        负数 -> 规则文件未配置「应付账款/负数」，按会计语义保持应付负债，填「应付账款」；
    4) 合计为零（应付款与暂估款相互抵消）时不分类，留空。"""
    if abs(total) < 1e-9:
        return ""
    sign = "正数" if total > 0 else "负数"
    for r in rules:
        if r["科目"] == "应付账款" and r["符号"] == sign:
            return r["重分类科目"]
    return "应付账款" if sign == "负数" else ""


def export_rule_template(template_path):
    """导出重分类规则模板，便于用户后续更新规则。"""
    wb = Workbook()
    ws = wb.active
    ws.title = "重分类规则"
    ws.append(RULES_HEADER)
    ws.append(["应收账款", "负数", "贷方", "预收款项"])   # 示例行
    ws.append(["应付账款", "正数", "借方", "预付款项"])   # 示例行

    note = wb.create_sheet("使用说明")
    lines = [
        "1. 规则判断逻辑：",
        "   总账科目：长文本 包含“规则科目关键词” 且 金额方向与“公司代码货币价值”列一致（正数/负数）时，",
        "   将该行重分类为“重分类科目”。",
        "2. “公司代码货币价值”列只能填写：正数 或 负数（金额>0为借方余额/正数，金额<0为贷方余额/负数）。",
        "3. “余额方向”列为辅助信息（借方/贷方），导入时以金额符号判断为主。",
        "4. 科目关键词越长匹配越优先（如“其他应付款”优先于“应付款”）。",
        "5. 表头请勿修改；示例行为演示数据，使用时可删除或覆盖。",
        "6. 将更新后的文件通过工具 --rules 参数导入即可生效。",
    ]
    for ln in lines:
        note.append([ln])

    header_fill = PatternFill("solid", fgColor="D9E1F2")
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = header_fill
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for col in ("A", "B", "C", "D"):
        ws.column_dimensions[col].width = 24
    for col in ("A",):
        note.column_dimensions[col].width = 120

    wb.save(template_path)
    return template_path


# ---------------------------------------------------------------------------
# 规则永久记忆（持久化 + 自动加载）
# ---------------------------------------------------------------------------

def _file_md5(path):
    """计算文件 MD5（逐块读取，兼容大文件）。"""
    h = hashlib.md5()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def persist_rule_memory(source_rules_path, rule_count=None):
    """将本次导入/生效的重分类规则固化保存到软件自身固定存储目录：
       <软件存储目录>/当前生效重分类规则.xlsx（覆盖旧副本）+ 重分类规则记忆.json。
    返回 (副本路径, 元信息 dict)；写入失败时返回 (None, None)。"""
    try:
        mem_dir = get_memory_dir()
        dst = os.path.join(mem_dir, MEMORY_RULES_FILENAME)
        # 源与目标为同一物理文件时跳过复制，仅刷新记忆记录
        try:
            same = os.path.realpath(source_rules_path) == os.path.realpath(dst)
        except Exception:
            same = False
        if not same:
            shutil.copy2(source_rules_path, dst)
        # 确保副本可写：源规则文件若为只读（copy2 保留只读属性），
        # 后续导入新规则覆盖副本会 Permission denied，导致记忆静默失效。
        try:
            os.chmod(dst, 0o644)
        except Exception:
            pass
        meta = {
            "rules_path": dst,
            "source_path": os.path.abspath(source_rules_path),
            "update_time": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            "md5": _file_md5(dst),
        }
        if rule_count is not None:
            meta["count"] = int(rule_count)
        meta_path = os.path.join(mem_dir, MEMORY_META_FILENAME)
        with open(meta_path, "w", encoding="utf-8") as f:
            json.dump(meta, f, ensure_ascii=False, indent=2)
        return dst, meta
    except Exception as e:
        print(f"      [记忆] 警告：规则记忆写入失败（不影响本次分析）：{e}")
        return None, None


def load_rule_memory():
    """读取软件自身存储目录下的规则记忆元信息，返回 dict 或 None。"""
    try:
        mem_dir = get_memory_dir()
        meta_path = os.path.join(mem_dir, MEMORY_META_FILENAME)
        if not os.path.isfile(meta_path):
            return None
        with open(meta_path, "r", encoding="utf-8") as f:
            meta = json.load(f)
        return meta if isinstance(meta, dict) else None
    except Exception:
        return None


def resolve_rules_path(explicit_rules=""):
    """确定本次分析使用的规则文件路径（来源优先级）：
       1) 显式指定规则文件（--rules / GUI 选择） -> 直接使用，并覆盖更新软件自身记忆；
       2) 软件自身记忆中的规则（当前生效重分类规则.xlsx 副本 + 记忆记录） -> 自动加载；
       3) 无显式指定且无记忆 -> 返回 (None, False, None)，提示先导入规则。
     不再扫描数据源目录：数据源目录中任何规则文件均不影响当前生效规则。"""
    if explicit_rules:
        return explicit_rules, False, None
    # 2) 软件自身记忆
    meta = load_rule_memory()
    if meta:
        mem_path = meta.get("rules_path") or ""
        if mem_path and os.path.isfile(mem_path):
            return mem_path, True, meta
    return None, False, None


# ---------------------------------------------------------------------------
# 数据源提取
# ---------------------------------------------------------------------------

def extract_source(data_dir, cfg, explicit_path=None):
    """从单个数据源文件提取目标字段行。跳过无总账科目（合计行/汇总行）。
    explicit_path: 可选，指定实际数据源文件路径（GUI 拖拽文件时覆盖默认文件名）。"""
    path = explicit_path or os.path.join(data_dir, cfg["file"])
    wb = load_workbook(path, read_only=True, data_only=True)
    ws = wb.worksheets[0]

    header = None
    for row in ws.iter_rows(min_row=1, max_row=1, values_only=True):
        header = list(row)
    if not header:
        wb.close()
        raise ValueError(f"文件 {cfg['file']} 无表头")

    header_map = build_header_map(header)
    subj_i = header_map.get(cfg["subject_col"])
    name_i = header_map.get(cfg["name_col"])
    fallback_i = header_map.get(cfg["fallback_name_col"]) if cfg.get("fallback_name_col") else None
    amount_i = header_map.get(cfg["amount_col"])
    dir_i = header_map.get(cfg["direction_col"]) if cfg.get("direction_col") else None

    if subj_i is None or amount_i is None:
        wb.close()
        raise ValueError(
            f"文件 {cfg['file']} 缺少必需列：总账科目：长文本 / 公司代码货币价值"
        )

    rows = []
    for row in ws.iter_rows(min_row=2, values_only=True):
        if row is None:
            continue
        subject = str(row[subj_i]).strip() if subj_i < len(row) and row[subj_i] is not None else ""
        if not subject:            # 跳过无科目的合计行
            continue
        amount = to_float(row[amount_i]) if amount_i < len(row) else 0.0
        name = ""
        if name_i is not None and name_i < len(row) and row[name_i] is not None:
            name = str(row[name_i]).strip()
        if not name and fallback_i is not None and fallback_i < len(row) and row[fallback_i] is not None:
            name = str(row[fallback_i]).strip()
        direction = ""
        if dir_i is not None and dir_i < len(row) and row[dir_i] is not None:
            direction = str(row[dir_i]).strip()
        rows.append({
            "source": cfg["source"],
            "subject": subject,
            "name": name,
            "amount": amount,
            "direction": direction,
        })
    wb.close()
    return rows


# ---------------------------------------------------------------------------
# 结果文件生成
# ---------------------------------------------------------------------------

def build_workbook(detail_rows, rules, sheet1_rows, sheet3_rows, ht_rows):
    """生成最终结果工作簿（提取明细 / 汇总表 / 应付含暂估 / 合同负债重分类其他流动负债 / 规则说明）。"""
    wb = Workbook()
    header_font = Font(bold=True, color="FFFFFF")
    header_fill = PatternFill("solid", fgColor="4472C4")
    center = Alignment(horizontal="center", vertical="center")

    def style_header(ws, row_idx, ncols):
        for c in range(1, ncols + 1):
            cell = ws.cell(row=row_idx, column=c)
            cell.font = header_font
            cell.fill = header_fill
            cell.alignment = center

    # ---- Sheet 0：提取明细 ----
    ws0 = wb.active
    ws0.title = "提取明细"
    ws0.append(["来源表", "总账科目：长文本", "名称", "公司代码货币价值", "借/贷标识"])
    style_header(ws0, 1, 5)
    for r in detail_rows:
        ws0.append([r["source"], r["subject"], r["name"], round(r["amount"], 2), r["direction"]])
    for col, w in zip("ABCDE", (22, 36, 34, 20, 12)):
        ws0.column_dimensions[col].width = w
    ws0.freeze_panes = "A2"

    # ---- Sheet 1（汇总表）：按科目+名称汇总 ----
    ws1 = wb.create_sheet("汇总表")
    ws1.append(["总账科目：长文本", "名称", "公司代码货币价值", "重分类科目"])
    style_header(ws1, 1, 4)
    for subject, name, amount, reclass in sheet1_rows:
        ws1.append([subject, name, amount, reclass])
    for col, w in zip("ABCD", (36, 34, 20, 20)):
        ws1.column_dimensions[col].width = w
    ws1.freeze_panes = "A2"

    # ---- Sheet 3：应付含暂估（按供应商汇总应付款/暂估款/合计，并按合计方向重分类） ----
    ws3 = wb.create_sheet("应付含暂估")
    ws3.append(["供应商", "应付款", "暂估款", "合计", "重分类科目"])
    style_header(ws3, 1, 5)
    for name, payable, est, total, reclass in sheet3_rows:
        ws3.append([name, payable, est, total, reclass])

    # 表尾仅保留两行汇总：按重分类科目分组合计，标签列与金额列着色（浅黄/浅蓝）+ 千位分隔两位小数
    yellow_fill = PatternFill("solid", fgColor="FFF2CC")
    blue_fill = PatternFill("solid", fgColor="DDEBF7")
    ap_total = round(sum(r[3] for r in sheet3_rows if r[4] == "应付账款"), 2)
    prepay_total = round(sum(r[3] for r in sheet3_rows if r[4] == "预付款项"), 2)

    ws3.append(["应付账款合计", "", "", ap_total, ""])
    ws3.cell(row=ws3.max_row, column=1).font = Font(bold=True)
    ws3.cell(row=ws3.max_row, column=4).font = Font(bold=True)
    ws3.cell(row=ws3.max_row, column=4).number_format = "#,##0.00"
    ws3.cell(row=ws3.max_row, column=1).fill = yellow_fill
    ws3.cell(row=ws3.max_row, column=4).fill = yellow_fill

    ws3.append(["预付款项合计", "", "", prepay_total, ""])
    ws3.cell(row=ws3.max_row, column=1).font = Font(bold=True)
    ws3.cell(row=ws3.max_row, column=4).font = Font(bold=True)
    ws3.cell(row=ws3.max_row, column=4).number_format = "#,##0.00"
    ws3.cell(row=ws3.max_row, column=1).fill = blue_fill
    ws3.cell(row=ws3.max_row, column=4).fill = blue_fill

    for col, w in zip("ABCDE", (40, 20, 20, 20, 18)):
        ws3.column_dimensions[col].width = w
    ws3.freeze_panes = "A2"

    # ---- Sheet 4：合同负债重分类其他流动负债 ----
    # 数据行：[名称, 公司代码货币价值(负), 绝对值, 可开票税额, 已开票税额, 差额, 科目余额]
    ws4 = wb.create_sheet("合同负债重分类其他流动负债")
    ws4.append(["名称", "公司代码货币价值", "公司代码货币价值绝对值",
                "可开票税额", "已开票税额", "差额", "科目余额"])
    style_header(ws4, 1, 7)
    for row in ht_rows:
        ws4.append(row)

    t_amt = round(sum(r[1] for r in ht_rows), 2)
    t_abs = round(sum(r[2] for r in ht_rows), 2)
    t_tax = round(sum(r[3] for r in ht_rows), 2)
    t_inv = round(sum(r[4] for r in ht_rows), 2)
    t_diff = round(sum(r[5] for r in ht_rows), 2)
    t_bal = round(sum(r[6] for r in ht_rows), 2)
    pos_diff = round(sum(r[5] for r in ht_rows if r[5] > 0), 2)
    neg_diff = round(sum(r[5] for r in ht_rows if r[5] < 0), 2)

    ws4.append([])
    ws4.append(["合计", t_amt, t_abs, t_tax, t_inv, t_diff, t_bal])
    ws4.cell(row=ws4.max_row, column=1).font = Font(bold=True)
    for c in range(2, 8):
        ws4.cell(row=ws4.max_row, column=c).font = Font(bold=True)

    ws4.append([])
    ws4.append(["差额大于零的合计", "", "", "", "", pos_diff, ""])
    ws4.cell(row=ws4.max_row, column=1).font = Font(bold=True)
    ws4.cell(row=ws4.max_row, column=6).font = Font(bold=True)
    ws4.append(["差额小于零的合计", "", "", "", "", neg_diff, ""])
    ws4.cell(row=ws4.max_row, column=1).font = Font(bold=True)
    ws4.cell(row=ws4.max_row, column=6).font = Font(bold=True)
    ws4.append(["结论：差额大于零的合计数为本期重分类其他流动负债金额", "", "", "", "", pos_diff, ""])
    ws4.cell(row=ws4.max_row, column=1).font = Font(bold=True, color="C00000")
    ws4.cell(row=ws4.max_row, column=6).font = Font(bold=True, color="C00000")

    for col, w in zip("ABCDEFG", (40, 18, 20, 18, 18, 18, 20)):
        ws4.column_dimensions[col].width = w
    ws4.freeze_panes = "A2"

    # ---- Sheet 5：规则说明 ----
    ws5 = wb.create_sheet("规则说明")
    ws5.append(["当前生效的重分类规则"])
    ws5.cell(row=1, column=1).font = Font(bold=True, size=12)
    ws5.append(RULES_HEADER)
    style_header(ws5, 2, 4)
    for r in rules:
        ws5.append([r["科目"], r["符号"], r["方向"], r["重分类科目"]])
    ws5.append([])
    ws5.append(["匹配逻辑：总账科目长文本包含规则关键词 且 金额符号（正数/负数）一致；多规则命中时取关键词最长者。"])
    for col, w in zip("ABCD", (36, 20, 14, 24)):
        ws5.column_dimensions[col].width = w
    ws5.column_dimensions["E"].width = 20
    ws5.freeze_panes = "A2"

    return wb


# ---------------------------------------------------------------------------
# 主流程
# ---------------------------------------------------------------------------

def run_analysis(data_dir, rules_path, out_path, source_files=None):
    """执行完整分析流程。
    source_files: 可选。显式指定的数据源文件，支持两种形式：
      - list[path]：按内容同类识别（classify_source_file）归入余额/暂估/客户类型；
      - dict {标准数据源文件名: 实际文件路径}：旧 GUI 拖拽形式，同样按内容识别。
    显式指定的类型优先；未显式指定的类型从 data_dir 目录内按内容同类识别
    （不再依赖固定文件名，兼容 供应商余额630 / 供应商暂估0630 / 客户余额630 等变体）。"""
    print("[1/6] 导入重分类规则 ...")
    rules = load_rules(rules_path)
    print(f"      已加载 {len(rules)} 条规则")

    print("[记忆] 将本次生效规则固化保存到软件自身存储目录（下次运行自动加载最新规则）...")
    mem_dst, _ = persist_rule_memory(rules_path, rule_count=len(rules))
    if mem_dst:
        print(f"      当前生效规则副本: {mem_dst}")
        print(f"      记忆记录文件:     {os.path.join(get_memory_dir(), MEMORY_META_FILENAME)}")
    else:
        print("      软件规则记忆目录不可写或写入失败，本次不更新规则记忆")

    print("[2/6] 读取并提取数据源字段（内容同类识别） ...")
    src_paths = resolve_source_paths(data_dir, source_files)
    detail_rows = []
    for cfg, stype in zip(SOURCE_CONFIG, SOURCE_TYPES):
        path = src_paths.get(stype)
        if not path or not os.path.isfile(path):
            raise FileNotFoundError(
                f"未找到 {cfg['source']} 数据源文件：目录内未识别到同类文件，"
                "也未显式指定。请检查数据源目录或直接选择对应文件。"
            )
        rows = extract_source(data_dir, cfg, path)
        detail_rows.extend(rows)
        print(f"      {cfg['source']}: 提取 {len(rows)} 行（文件 {os.path.basename(path)}）")
    print(f"      明细合计 {len(detail_rows)} 行")

    print("[3/6] 生成汇总表（按总账科目：长文本 + 名称 汇总）...")
    agg = defaultdict(float)
    for r in detail_rows:
        agg[(r["subject"], r["name"])] += r["amount"]

    sheet1_rows = []
    for (subject, name), amt in agg.items():
        amt_rounded = round(amt, 2)
        if abs(amt_rounded) < 1e-9:      # 删除合计为零的行
            continue
        sheet1_rows.append([
            subject, name, amt_rounded,
            match_reclass(subject, amt_rounded, rules),
        ])
    sheet1_rows.sort(key=lambda x: (x[0] or "", x[1] or ""))
    zero_removed = len(agg) - len(sheet1_rows)
    print(f"      汇总 {len(agg)} 组，删除合计为零 {zero_removed} 组，保留 {len(sheet1_rows)} 行")

    print("[4/6] 生成 sheet3（应付含暂估：按供应商汇总 应付款/暂估款/合计）...")
    # 匹配口径（基于 sheet1 实际科目变体确认）：
    #   - 应付账款科目        = 总账科目含「应付账款」
    #   - 应付款              = 应付账款科目中不含「暂估」的科目金额合计
    #                          （实际数据：材料款-集团内/集团外、费用、待摊费用 等）
    #   - 暂估款              = 应付账款科目中含「暂估」的科目金额合计
    #                          （实际数据：暂估材料款、暂估非材料款、暂估运费）
    #   - 合计                = 应付款 + 暂估款
    payable_agg = defaultdict(float)   # 供应商 -> 应付款（不含暂估）
    est_agg = defaultdict(float)       # 供应商 -> 暂估款（含暂估）
    for subject, name, amt, _reclass in sheet1_rows:
        if "应付账款" not in subject:
            continue
        if "暂估" in subject:
            est_agg[name] += amt
        else:
            payable_agg[name] += amt

    sheet3_rows = []
    all_suppliers = sorted(set(payable_agg) | set(est_agg), key=lambda x: x or "")
    for name in all_suppliers:
        payable = round(payable_agg.get(name, 0.0), 2)
        est = round(est_agg.get(name, 0.0), 2)
        total = round(payable + est, 2)
        reclass = match_sheet3_reclass(payable, est, total, rules)
        # 保留全部供应商（含应付款与暂估款相互抵消、合计为零者），
        # 保证 sheet3 与明细/独立聚合完全可核对
        sheet3_rows.append([name, payable, est, total, reclass])
    pos_cnt = sum(1 for r in sheet3_rows if r[3] > 0)
    neg_cnt = sum(1 for r in sheet3_rows if r[3] < 0)
    zero_cnt = len(sheet3_rows) - pos_cnt - neg_cnt
    print(f"      应付含暂估 {len(sheet3_rows)} 个供应商")
    print(f"      应付款合计 {round(sum(r[1] for r in sheet3_rows), 2)}，"
          f"暂估款合计 {round(sum(r[2] for r in sheet3_rows), 2)}，"
          f"总计 {round(sum(r[3] for r in sheet3_rows), 2)}")
    print(f"      重分类分布（按合计方向）：正数→预付款项 {pos_cnt} 个，"
          f"负数→应付账款 {neg_cnt} 个，合计为零未分类 {zero_cnt} 个")

    print("[5/6] 生成合同负债重分类其他流动负债 ...")
    # 数据源：汇总表（原 sheet1）中 科目含「合同负债」且 公司代码货币价值为负数的行
    # 实际科目变体：合同负债-集团外（负数行）、合同负债-待确认收入（均为正数，不构成负数行）
    ht_neg = [r for r in sheet1_rows if "合同负债" in r[0] and r[2] < 0]
    # 已开票税额：汇总表「合同负债-待确认收入」科目有数值的行，按同一客户（名称）取值
    invoiced_map = {}
    for r in sheet1_rows:
        if "合同负债-待确认收入" in r[0] and r[2] != 0:
            invoiced_map[r[1]] = round(r[2], 2)

    def _has_cn(name):
        return any("\u4e00" <= ch <= "\u9fff" for ch in (name or ""))

    ht_rows = []
    # 多行客户按负数行去重汇总：同一客户若存在多个合同负债负数行（不同科目变体），
    # 先按客户聚合金额为一行，正数/零行一律不进入本表（ht_neg 已严格过滤 < 0）
    ht_agg = defaultdict(float)
    for _subj, name, amt, _re in ht_neg:
        ht_agg[name] += amt
    for name, amt in sorted(ht_agg.items(), key=lambda x: x[0] or ""):
        abs_amt = round(-amt, 2)
        taxable = round(abs_amt / 1.13 * 0.13, 2) if _has_cn(name) else 0.0
        invoiced = invoiced_map.get(name, 0.0)
        diff = round(taxable - invoiced, 2)
        balance = round(abs_amt - invoiced, 2)
        ht_rows.append([name, round(amt, 2), abs_amt, taxable, invoiced, diff, balance])
    cn_cnt = sum(1 for r in ht_rows if _has_cn(r[0]))
    en_cnt = len(ht_rows) - cn_cnt
    pos_diff = round(sum(r[5] for r in ht_rows if r[5] > 0), 2)
    neg_diff = round(sum(r[5] for r in ht_rows if r[5] < 0), 2)
    print(f"      合同负债负数行 {len(ht_rows)} 个客户（国内 {cn_cnt} / 国外 {en_cnt}）")
    print(f"      可开票税额合计 {round(sum(r[3] for r in ht_rows), 2)}，已开票税额合计 {round(sum(r[4] for r in ht_rows), 2)}，差额合计 {round(sum(r[5] for r in ht_rows), 2)}")
    print(f"      差额大于零合计 {pos_diff}，差额小于零合计 {neg_diff}")
    print(f"      结论：差额大于零的合计数（{pos_diff}）为本期重分类其他流动负债金额")

    print("[6/6] 写出结果文件 ...")
    wb = build_workbook(detail_rows, rules, sheet1_rows, sheet3_rows, ht_rows)
    wb.save(out_path)
    print(f"      已保存: {out_path}")

    # 简要统计输出
    print()
    print("== 运行统计 ==")
    print(f"  提取明细行数        : {len(detail_rows)}")
    print(f"  汇总表行数         : {len(sheet1_rows)}（删除合计为零 {zero_removed} 组）")
    print(f"  sheet3 供应商数     : {len(sheet3_rows)}")
    print(f"  sheet3 应付款合计   : {round(sum(r[1] for r in sheet3_rows), 2)}")
    print(f"  sheet3 暂估款合计   : {round(sum(r[2] for r in sheet3_rows), 2)}")
    print(f"  sheet3 总计         : {round(sum(r[3] for r in sheet3_rows), 2)}")
    print(f"  sheet3 正数(预付款项): {sum(1 for r in sheet3_rows if r[3] > 0)}")
    print(f"  sheet3 负数(应付账款): {sum(1 for r in sheet3_rows if r[3] < 0)}")
    print(f"  sheet3 合计为零      : {sum(1 for r in sheet3_rows if r[3] == 0)}")
    print(f"  合同负债重分类行数   : {len(ht_rows)}")
    print(f"  合同负债可开票税额   : {round(sum(r[3] for r in ht_rows), 2)}")
    print(f"  合同负债已开票税额   : {round(sum(r[4] for r in ht_rows), 2)}")
    print(f"  合同负债差额>0合计   : {round(sum(r[5] for r in ht_rows if r[5] > 0), 2)}")
    print(f"  合同负债差额<0合计   : {round(sum(r[5] for r in ht_rows if r[5] < 0), 2)}")
    print(f"  合同负债科目余额合计 : {round(sum(r[6] for r in ht_rows), 2)}")
    return out_path


def main():
    parser = argparse.ArgumentParser(description="往来重分类分析工具")
    parser.add_argument("--data-dir", default=".",
                        help="数据源目录（包含三个数据源 xlsx 与规则文件），默认当前目录")
    parser.add_argument("--rules", default="",
                        help="重分类规则文件路径（默认 <data-dir>/重分类规则.xlsx）")
    parser.add_argument("--out", default="",
                        help="结果输出文件路径（默认 <data-dir>/往来重分类分析结果.xlsx）")
    parser.add_argument("--export-template", default="",
                        help="导出重分类规则模板到指定路径后退出")
    parser.add_argument("--balance", default="",
                        help="余额数据源文件路径（显式指定，优先于目录内同类识别）")
    parser.add_argument("--estimate", default="",
                        help="暂估数据源文件路径（显式指定，优先于目录内同类识别）")
    parser.add_argument("--customer", default="",
                        help="客户数据源文件路径（显式指定，优先于目录内同类识别）")
    args = parser.parse_args()

    if args.export_template:
        path = export_rule_template(args.export_template)
        print(f"重分类规则模板已导出: {path}")
        return

    data_dir = args.data_dir
    out_path = args.out or os.path.join(data_dir, "往来重分类分析结果.xlsx")
    if not os.path.isdir(data_dir):
        sys.exit(f"数据源目录不存在: {data_dir}")

    rules_path, auto_loaded, _ = resolve_rules_path(args.rules)
    if rules_path is None:
        sys.exit(
            "未找到可用的重分类规则。\n"
            "  软件自身尚未记忆规则：请先通过 --rules 导入规则文件，\n"
            "  或运行 --export-template 导出规则模板填写规则后再导入。\n"
            f"  规则固化目录: {get_memory_dir()}\n"
        )
    if auto_loaded:
        print(f"[记忆] 自动加载当前生效规则: {rules_path}")
    else:
        print(f"[规则] 使用规则文件: {rules_path}")

    explicit = [p for p in (args.balance, args.estimate, args.customer) if p]
    run_analysis(data_dir, rules_path, out_path, explicit or None)


if __name__ == "__main__":
    main()
