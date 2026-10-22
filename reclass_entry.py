#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
往来重分类分析工具 - 统一入口（GUI + CLI）
============================================
- 双击 .app / 无参数运行  -> 打开 tkinter 图形界面（紧凑卡片式布局）：
    * 顶部细窄标题栏单行；
    * 规则文件卡片单行：规则状态提示（已有规则 N 条已生效 · 更新于… /
      需要先载入规则文件）+「导入规则文件」「导出模板」两个按钮，
      导入后自动固化到软件自身记忆，无需再次选择路径；
    * 数据源与输出结果两列并排紧凑分组：
      - 数据源：已加载文件列表（文件名 / 类型 / 完整路径），每行支持
        勾选参与分析、单独移除；支持 浏览文件 / 选择目录 / 拖入 追加；
        提供「清除列表」一键清空；未勾选任何文件时自动尝试目录识别；
      - 输出结果：输出路径 + 浏览；
    * 执行分析、运行日志、窗口任意位置拖拽兜底识别。
- 终端带参数运行          -> 命令行模式，完全兼容原 CLI 参数：
    --data-dir  数据源目录（默认当前目录）
    --rules     重分类规则文件（留空则自动加载软件自身记忆中的当前生效规则）
    --out       结果输出路径（默认 <data-dir>/往来重分类分析结果.xlsx）
    --export-template <path>  导出重分类规则模板
"""

import sys
import os
import io
import queue
import threading
import contextlib
import datetime

import reclass_core as core


# ---------------------------------------------------------------------------
# 命令行模式（兼容原 CLI）
# ---------------------------------------------------------------------------

def cli_main():
    core.main()


# ---------------------------------------------------------------------------
# 可测逻辑层：不依赖 tkinter（单元测试直接调用）
# ---------------------------------------------------------------------------

class SrcFileEntry:
    """单个已加载数据源文件：路径 + 内容同类识别结果（缓存，避免重复读表头）。"""
    __slots__ = ("path", "stype", "source")

    def __init__(self, path, cls=None):
        self.path = path
        # cls 由调用方（模型层）传入，实现判识结果复用：每个文件最多读 1 次表头。
        if cls is None:
            cls = core.classify_source_file(path)
        if cls:
            self.stype, self.source = cls[0], cls[1]["source"]
        else:
            self.stype, self.source = None, "未识别"


class SrcListModel:
    """数据源文件列表模型：加载 / 勾选 / 移除 / 清除。
    不依赖 tkinter，便于单元测试验证增删清与勾选逻辑。"""

    def __init__(self):
        self.entries = []     # 有序 SrcFileEntry
        self.checked = {}     # path -> bool（默认 True）
        self._cls_cache = {}  # path -> (stype, cfg)：判识结果缓存（每文件只读 1 次表头）

    def add(self, paths):
        """追加可识别的数据源文件（自动去重，每文件仅读 1 次表头）。
        返回 (matched, ignored)。"""
        matched, ignored = [], []
        for p in (paths or []):
            if not p or not os.path.isfile(p):
                continue
            if p in self.checked:
                continue
            cls = core.classify_source_file(p)
            if cls:
                ent = SrcFileEntry(p, cls=cls)
                self.entries.append(ent)
                self.checked[p] = True
                self._cls_cache[p] = cls
                matched.append(p)
            else:
                ignored.append(p)
        return matched, ignored

    def add_identified(self, path_cls):
        """追加数据源文件（判识结果由调用方/后台线程提供，复用缓存不再读表头）。
        path_cls: [(path, cls_or_None), ...]。返回 (matched, ignored)。"""
        matched, ignored = [], []
        for p, cls in (path_cls or []):
            if not p or not os.path.isfile(p):
                continue
            if p in self.checked:
                continue
            if cls:
                ent = SrcFileEntry(p, cls=cls)
                self.entries.append(ent)
                self.checked[p] = True
                self._cls_cache[p] = cls
                matched.append(p)
            else:
                ignored.append(p)
        return matched, ignored

    def cls_of(self, path):
        """返回 path 已缓存的判识结果 (stype, cfg)，未缓存返回 None。"""
        return self._cls_cache.get(path)

    def remove(self, path):
        if path in self.checked:
            self.entries = [e for e in self.entries if e.path != path]
            self.checked.pop(path, None)
            self._cls_cache.pop(path, None)

    def clear(self):
        self.entries = []
        self.checked = {}
        self._cls_cache = {}

    def set_checked(self, path, value):
        if path in self.checked:
            self.checked[path] = bool(value)

    def selected(self):
        return [e.path for e in self.entries if self.checked.get(e.path)]

    def count_by_type(self):
        """返回 {stype: n} 统计（按内容识别类型）。"""
        res = {}
        for e in self.entries:
            if e.stype:
                res[e.stype] = res.get(e.stype, 0) + 1
        return res


def compute_rule_status():
    """基于软件自身记忆计算规则状态文案与颜色。
    返回 (text, color_key)，color_key in ("ok", "warn")。"""
    path, auto, meta = core.resolve_rules_path("")
    if path and os.path.isfile(path):
        n = 0
        try:
            n = len(core.load_rules(path))
        except Exception:
            n = 0
        ut = (meta or {}).get("update_time", "") if isinstance(meta, dict) else ""
        if not ut:
            try:
                ut = datetime.datetime.fromtimestamp(
                    os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
            except Exception:
                ut = ""
        suffix = (" · 更新于 " + ut) if ut else ""
        return "已有规则 {} 条已生效{}".format(n, suffix), "ok"
    return "需要先载入规则文件（点击「导入规则文件」）", "warn"


# ---------------------------------------------------------------------------
# 图形界面模式（tkinter + ttk，clam 主题自定义样式，紧凑布局）
# ---------------------------------------------------------------------------

def gui_main():
    import tkinter as tk
    from tkinter import ttk, filedialog, messagebox, scrolledtext
    # tkinterdnd2 必须在创建 TkinterDnD.Tk() 根窗口前导入（本函数内位于
    # 根窗口创建之前，满足"不能延迟加载"要求）；PyInstaller 打包时由
    # 官方 hook 自动收集 tkdnd 平台资源。CLI 分支不进入本函数，故 CLI
    # 打包可安全 excludes tkinterdnd2。
    from tkinterdnd2 import TkinterDnD, DND_FILES

    # ---- 配色 ----
    BG = "#f1f5f9"
    CARD = "#ffffff"
    BORDER = "#e2e8f0"
    HEADER_BG = "#12324f"
    HEADER_ACCENT = "#3b82f6"
    TEXT = "#0f172a"
    MUTED = "#64748b"
    OK = "#15803d"
    WARN = "#b45309"

    root = TkinterDnD.Tk()
    root.title("往来重分类分析工具")
    root.geometry("1020x820")
    root.minsize(900, 660)
    root.configure(bg=BG)

    # ---- ttk 样式（clam 主题，紧凑内边距） ----
    style = ttk.Style(root)
    try:
        style.theme_use("clam")
    except Exception:
        pass
    style.configure("Accent.TButton",
                    background=HEADER_ACCENT, foreground="#ffffff",
                    borderwidth=0, focusthickness=0, padding=(12, 5),
                    font=("PingFang SC", 10, "bold"))
    style.map("Accent.TButton",
              background=[("pressed", "#1d4ed8"), ("active", "#3f8cff"),
                          ("disabled", "#9db8dd")],
              foreground=[("disabled", "#e8eef5")])
    style.configure("Ghost.TButton",
                    background=CARD, foreground="#334155",
                    borderwidth=1, relief="solid", padding=(9, 4),
                    font=("PingFang SC", 9))
    style.map("Ghost.TButton",
              background=[("active", "#eef2f7"), ("pressed", "#e2e8f0")])
    # 「移除」按钮：默认可见样式（浅红底 + 深红字 + 实线边框），避免与卡片同色不可见
    style.configure("Remove.TButton",
                    background="#fee2e2", foreground="#b91c1c",
                    borderwidth=1, relief="solid", padding=(9, 3),
                    font=("PingFang SC", 9))
    style.map("Remove.TButton",
              background=[("active", "#fecaca"), ("pressed", "#fca5a5")],
              foreground=[("disabled", "#c7a3a3")])
    # 列表管理按钮：标准可见样式（白底 + 深灰边框 + 黑字），参考图2
    style.configure("White.TButton",
                    background="#ffffff", foreground="#111827",
                    borderwidth=2, relief="solid",
                    bordercolor="#475569",
                    padding=(13, 5),
                    font=("PingFang SC", 9))
    style.map("White.TButton",
              background=[("active", "#f1f5f9"), ("pressed", "#e2e8f0")],
              bordercolor=[("active", "#334155"), ("pressed", "#1e293b")],
              foreground=[("disabled", "#9ca3af")])
    style.configure("Run.TButton",
                    background="#166534", foreground="#ffffff",
                    borderwidth=0, focusthickness=0, padding=(22, 6),
                    font=("PingFang SC", 11, "bold"))
    style.map("Run.TButton",
              background=[("pressed", "#14532d"), ("active", "#1f8a4c"),
                          ("disabled", "#a7c4b1")],
              foreground=[("disabled", "#e8f0ea")])

    var_data_dir = tk.StringVar()
    var_out = tk.StringVar()
    src_model = SrcListModel()   # 数据源文件列表（勾选/移除/清除）
    q = queue.Queue()
    state = {"running": False}

    # ==================================================================
    # 顶部细窄标题栏（一行）
    # ==================================================================
    header = tk.Frame(root, bg=HEADER_BG, height=46)
    header.pack(fill="x")
    header.pack_propagate(False)
    tk.Label(header, text="往来重分类分析工具", bg=HEADER_BG, fg="#ffffff",
             font=("PingFang SC", 13, "bold")).place(x=16, y=11)
    tk.Label(header, text="数据源提取 · 重分类规则 · 应付含暂估分析",
             bg=HEADER_BG, fg="#a9bccd",
             font=("PingFang SC", 9)).place(relx=1.0, x=-16, anchor="e", y=15)

    # ==================================================================
    # 主体（紧凑卡片）
    # ==================================================================
    body = tk.Frame(root, bg=BG)
    body.pack(fill="both", expand=True, padx=14, pady=8)

    def make_card(parent, title, hint=""):
        card = tk.Frame(parent, bg=CARD, highlightbackground=BORDER,
                        highlightthickness=1, bd=0)
        head = tk.Frame(card, bg=CARD)
        head.pack(fill="x", padx=12, pady=(7, 2))
        tk.Label(head, text=title, bg=CARD, fg=TEXT,
                 font=("PingFang SC", 11, "bold")).pack(side="left")
        if hint:
            tk.Label(head, text=hint, bg=CARD, fg=MUTED,
                     font=("PingFang SC", 8)).pack(side="left", padx=6)
        return card

    # ---------------- 1. 规则文件（单行：状态提示 + 导入 + 导出） ----------------
    rules_card = make_card(body, "规则文件",
                           "导入即自动记忆，无需重复选择路径")
    rules_card.pack(fill="x", pady=(0, 6))

    rules_row = tk.Frame(rules_card, bg=CARD)
    rules_row.pack(fill="x", padx=12, pady=(0, 9))
    rules_status = tk.Label(rules_row, text="正在识别规则…", bg=CARD, fg=MUTED,
                            font=("PingFang SC", 9, "bold"), anchor="w")
    rules_status.pack(side="left", fill="x", expand=True)
    btn_export = ttk.Button(rules_row, text="导出模板", style="Ghost.TButton",
                            command=None)
    btn_import = ttk.Button(rules_row, text="导入规则文件", style="Accent.TButton",
                            command=None)
    btn_export.pack(side="right")
    btn_import.pack(side="right", padx=(0, 6))

    # ---------------- 2/3. 数据源 + 输出结果（两列并排） ----------------
    mid = tk.Frame(body, bg=BG)
    mid.pack(fill="x", pady=(0, 6))
    mid.columnconfigure(0, weight=3)
    mid.columnconfigure(1, weight=2)

    data_card = tk.Frame(mid, bg=CARD, highlightbackground=BORDER,
                         highlightthickness=1, bd=0)
    data_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
    data_card.columnconfigure(0, weight=1)
    data_card.rowconfigure(1, weight=1)
    # 区域标题：顶部左对齐、常规字体（无加粗），参考图2
    tk.Label(data_card, text="待处理文件列表（可拖拽 .xlsx/.xls 到此处）",
             bg=CARD, fg=TEXT, font=("PingFang SC", 9),
             anchor="w").grid(row=0, column=0, sticky="ew",
                              padx=12, pady=(8, 2))

    # 列表区：Treeview + 垂直滚动条并排，独占一行（grid 布局保证按钮区在底部整行）
    list_area = tk.Frame(data_card, bg=CARD)
    list_area.grid(row=1, column=0, sticky="nsew", padx=12, pady=(2, 4))
    list_area.columnconfigure(0, weight=1)

    # 无列头/无勾选框的滚动列表（Treeview 单列；空列表时留白）
    tree = ttk.Treeview(list_area, columns=(), show="tree",
                        selectmode="extended", height=8)
    style.configure("Treeview",
                    background=CARD, fieldbackground=CARD, foreground=TEXT,
                    borderwidth=0, font=("PingFang SC", 9), rowheight=22)
    style.configure("Treeview.Heading", font=("PingFang SC", 9, "bold"))
    tree.column("#0", width=600, minwidth=360, anchor="w", stretch=True)
    tree_scroll = ttk.Scrollbar(list_area, orient="vertical",
                                command=tree.yview)
    tree.configure(yscrollcommand=tree_scroll.set)

    def _on_mousewheel(event):
        tree.yview_scroll(int(-event.delta / 120), "units")
    tree.bind("<MouseWheel>", _on_mousewheel)

    tree.pack(side="left", fill="both", expand=True)
    tree_scroll.pack(side="right", fill="y", padx=(6, 0))

    # 底部三按钮：添加文件 / 移除选中 / 清空列表（白底深灰边框，水平并排，整行完整显示）
    list_btns = tk.Frame(data_card, bg=CARD)
    list_btns.grid(row=2, column=0, sticky="ew", padx=12, pady=(2, 6))
    btn_add = ttk.Button(list_btns, text="添加文件", style="White.TButton",
                         command=None)
    btn_remove_sel = ttk.Button(list_btns, text="移除选中", style="White.TButton",
                                command=None)
    btn_clear = ttk.Button(list_btns, text="清空列表", style="White.TButton",
                           command=None)
    btn_add.pack(side="left")
    btn_remove_sel.pack(side="left", padx=(8, 0))
    btn_clear.pack(side="left", padx=(8, 0))

    src_summary = tk.Label(data_card, text="", bg=CARD, fg=MUTED,
                           font=("PingFang SC", 8), anchor="w", justify="left")
    src_summary.grid(row=3, column=0, sticky="ew", padx=12, pady=(0, 6))

    out_card = make_card(mid, "输出结果", "默认生成到数据源目录下")
    out_card.grid(row=0, column=1, sticky="nsew", padx=(6, 0))
    out_row = tk.Frame(out_card, bg=CARD)
    out_row.pack(fill="x", padx=12, pady=(0, 4))
    ent_out = tk.Entry(out_row, textvariable=var_out, bg="#f8fafc",
                       relief="solid", bd=1, highlightthickness=1,
                       highlightbackground=BORDER, highlightcolor=HEADER_ACCENT,
                       font=("PingFang SC", 10))
    ent_out.pack(side="left", fill="x", expand=True, ipady=3)
    btn_out_browse = ttk.Button(out_row, text="浏览…", style="Ghost.TButton",
                                command=None)
    btn_out_browse.pack(side="left", padx=(6, 0))

    out_hint = tk.Label(
        out_card,
        text="生成文件：往来重分类分析结果.xlsx\n含 提取明细 / 汇总表 / 应付含暂估 / 合同负债重分类其他流动负债 / 规则说明",
        bg=CARD, fg=MUTED, font=("PingFang SC", 8), anchor="w", justify="left")
    out_hint.pack(fill="x", padx=12, pady=(0, 8))

    # ---------------- 执行按钮（紧凑行） ----------------
    run_frame = tk.Frame(body, bg=BG)
    run_frame.pack(fill="x", pady=(0, 6))
    btn_run = ttk.Button(run_frame, text="执行分析", style="Run.TButton",
                         command=None)
    btn_run.pack(anchor="center")

    # ---------------- 运行日志 ----------------
    log_card = make_card(body, "运行日志", "")
    log_card.pack(fill="both", expand=True)
    text = scrolledtext.ScrolledText(
        log_card, wrap="word", font=("Menlo", 9),
        bg="#0d1520", fg="#d7e0e8", insertbackground="#d7e0e8",
        relief="solid", bd=0, highlightthickness=1, highlightbackground=BORDER,
        padx=8, pady=5, state="disabled")
    text.pack(fill="both", expand=True, padx=12, pady=(0, 10))

    def log(msg):
        text.configure(state="normal")
        text.insert("end", msg)
        text.see("end")
        text.configure(state="disabled")

    # ==================================================================
    # 状态刷新：规则（条数/更新时间）与数据源（文件列表就绪）
    # ==================================================================
    def refresh_rule_status():
        txt, key = compute_rule_status()
        rules_status.config(text=txt, fg=(OK if key == "ok" else WARN))

    def refresh_data_status():
        data_dir = var_data_dir.get().strip()
        ok_dir = bool(data_dir) and os.path.isdir(data_dir)
        counts = src_model.count_by_type()
        part = []
        for stype, src in zip(core.SOURCE_TYPES, core.SOURCE_CONFIG):
            if counts.get(stype):
                part.append("{} {}".format(counts[stype], src["source"]))
        if src_model.entries:
            summary = "已加载 {} 个文件（{}）".format(
                len(src_model.entries), " · ".join(part) if part else "待识别")
            fg = OK
        else:
            summary = ""
            fg = MUTED
        if ok_dir:
            summary += (" ｜ 数据源目录：{}".format(data_dir) if summary else
                        "数据源目录：{}".format(data_dir))
        src_summary.config(text=summary, fg=fg)

    # ==================================================================
    # 数据源列表渲染（Treeview 增量插入，避免增/删/清全量重建）
    # ==================================================================
    def _row_text(ent):
        return "{}   {}   {}".format(os.path.basename(ent.path),
                                     ent.source or "未识别", ent.path)

    def _render_src_list():
        """全量重建（仅用于初始化/清空）；日常增删走增量插入。空列表留白。"""
        for i in tree.get_children():
            tree.delete(i)
        if not src_model.entries:
            refresh_data_status()
            return
        for ent in src_model.entries:
            tree.insert("", "end", iid=ent.path, text=_row_text(ent))
        refresh_data_status()

    def _append_src_rows(matched_paths):
        """增量插入新行：仅对新增文件插行，不销毁已有行。"""
        if not matched_paths:
            return
        for ent in src_model.entries:
            if ent.path in matched_paths and not tree.exists(ent.path):
                tree.insert("", "end", iid=ent.path, text=_row_text(ent))
        refresh_data_status()

    def _remove_src(path):
        src_model.remove(path)
        log("[列表] 已移除数据源文件: {}\n".format(path))
        if tree.exists(path):
            tree.delete(path)
        refresh_data_status()

    def _remove_selected():
        sel = tree.selection()
        if not sel:
            log("[列表] 未选中任何文件，无需移除\n")
            return
        for path in sel:
            if tree.exists(path):
                tree.delete(path)
            if any(e.path == path for e in src_model.entries):
                src_model.remove(path)
        log("[列表] 已移除选中的 {} 个数据源文件\n".format(len(sel)))
        refresh_data_status()

    def _clear_src():
        if not src_model.entries:
            return
        n = len(src_model.entries)
        src_model.clear()
        for i in tree.get_children():
            tree.delete(i)
        log("[列表] 已清空全部 {} 个数据源文件\n".format(n))
        refresh_data_status()

    # ==================================================================
    # 拖拽数据解析（tkinterdnd2 / macOS Finder 格式）
    # ==================================================================
    def parse_drop_data(data):
        """解析拖拽数据为路径列表。
        macOS Finder 传入格式为 {/path/a} {/path/b}（花括号包裹路径，
        路径内含空格也安全）；Windows 为 "C:/a" "C:/b"。优先 Tcl 原生
        splitlist 解析，失败则逐字符按花括号深度解析，最后回退空白分割。"""
        if not data:
            return []
        try:
            parts = root.tk.splitlist(data)
            if parts:
                return [str(p) for p in parts if str(p).strip()]
        except Exception:
            pass
        paths = []
        cur = []
        depth = 0
        for ch in data:
            if ch == "{":
                depth += 1
                if depth == 1:
                    cur = []
                else:
                    cur.append(ch)
            elif ch == "}":
                if depth > 0:
                    depth -= 1
                    if depth == 0:
                        p = "".join(cur).strip()
                        if p:
                            paths.append(p)
                        cur = []
                else:
                    cur.append(ch)
            else:
                if depth > 0:
                    cur.append(ch)
                else:
                    if ch.isspace():
                        if cur:
                            p = "".join(cur).strip()
                            if p:
                                paths.append(p)
                            cur = []
                    else:
                        cur.append(ch)
        if cur:
            p = "".join(cur).strip()
            if p:
                paths.append(p)
        return paths

    # ------------------------------------------------------------------
    # 拖拽处理器
    # ------------------------------------------------------------------
    def handle_data_drop(event):
        """拖入数据源目录 / 一个或多个数据源 Excel：后台线程判识，主线程增量渲染。"""
        paths = parse_drop_data(event.data)
        dirs = [p for p in paths if os.path.isdir(p)]
        files = [p for p in paths if os.path.isfile(p)]
        added = False
        if dirs:
            var_data_dir.set(dirs[0])
            log("[拖拽] 数据源目录已设置为: " + dirs[0] + "\n")
            added = True
        if files:
            def _done(matched, ignored):
                if matched:
                    common = os.path.dirname(matched[0])
                    if all(os.path.dirname(p) == common for p in matched) and os.path.isdir(common):
                        var_data_dir.set(common)
                        log("[拖拽] 数据源目录已自动设置为: " + common + "\n")
            add_src_files_async(files, on_done=_done)
            added = True
        if not added:
            log("[拖拽] 未识别到可用的数据源，请拖入 数据源目录，或 余额/暂估/客户 三类 xlsx "
                "（文件名不限，按列结构自动识别）\n")

    def handle_rules_drop(event):
        """拖入规则文件：直接导入并固化到软件自身记忆。"""
        paths = parse_drop_data(event.data)
        xlsx = [p for p in paths if os.path.isfile(p) and p.lower().endswith(".xlsx")]
        if xlsx:
            p = xlsx[0]
            n = 0
            try:
                n = len(core.load_rules(p))
            except Exception:
                n = 0
            dst, meta = core.persist_rule_memory(p, rule_count=n)
            if dst:
                log("[拖拽] 规则文件已导入并记忆: {}（共 {} 条）→ {}\n".format(p, n, dst))
                refresh_rule_status()
            else:
                log("[拖拽] 规则记忆写入失败，请检查软件存储目录权限\n")
        else:
            log("[拖拽] 未识别到 Excel 规则文件（请拖入 .xlsx 规则表）\n")

    def on_drop_anywhere(event):
        """窗口空白处兜底：按文件名关键词粗分流（避免预判识重复读表头），
        数据源判识交由后台线程，规则文件按「规则」关键词命中导入。"""
        paths = parse_drop_data(event.data)
        files = [p for p in paths if os.path.isfile(p)]
        if files and any("规则" in os.path.basename(p) for p in files):
            handle_rules_drop(event)
        else:
            handle_data_drop(event)

    # ==================================================================
    # 对话框动作
    # ==================================================================
    def add_src_files(paths):
        """将显式数据源文件追加到列表（按内容同类识别，去重，每文件仅读 1 次表头）。
        返回 (matched, ignored)。"""
        matched, ignored = src_model.add(paths)
        if matched:
            for p in matched:
                cls = src_model.cls_of(p)
                label = cls[1]["source"] if cls else "未识别"
                log("[识别] 数据源 {} <- {}\n".format(label, p))
            _append_src_rows(set(matched))
        return matched, ignored

    def add_src_files_async(paths, on_done=None):
        """后台线程判识（不阻塞 GUI），完成后在主线程增量渲染。
        on_done(matched, ignored) 在主线程回调。"""
        def _bg():
            results = []
            for p in (paths or []):
                if not p or not os.path.isfile(p):
                    results.append((p, None))
                    continue
                results.append((p, core.classify_source_file(p)))
            root.after(0, lambda: _apply_src_results(results, on_done))
        threading.Thread(target=_bg, daemon=True).start()

    def _apply_src_results(results, on_done=None):
        """后台判识结果回到主线程：模型写入 + 增量渲染 + 日志提示。"""
        matched, ignored = src_model.add_identified(results)
        if matched:
            for p in matched:
                cls = src_model.cls_of(p)
                label = cls[1]["source"] if cls else "未识别"
                log("[识别] 数据源 {} <- {}\n".format(label, p))
            _append_src_rows(set(matched))
        if ignored:
            log("[提示] 忽略无法识别的文件（未匹配到余额/暂估/客户同类结构）: {}\n".format(
                "; ".join(ignored)))
        if on_done:
            on_done(matched, ignored)

    def ask_source_files():
        """浏览并追加数据源文件（支持多选：余额/暂估/客户 任意组合，后台判识不卡界面）。"""
        ps = filedialog.askopenfilenames(
            title="选择数据源文件（可多选：余额/暂估/客户 三类自动识别）",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if ps:
            def _done(matched, ignored):
                if not matched and ignored:
                    messagebox.showerror(
                        "无法识别",
                        "所选文件未匹配到余额/暂估/客户任一数据源结构（列结构特征），请确认文件正确。",
                    )
            add_src_files_async(list(ps), on_done=_done)

    def import_rules_file():
        """选择规则文件并固化到软件自身记忆（导入后免再次选择）。"""
        if state["running"]:
            return
        p = filedialog.askopenfilename(
            title="选择重分类规则文件（导入后将自动记忆）",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if not p:
            return
        n = 0
        try:
            n = len(core.load_rules(p))
        except Exception as e:
            messagebox.showerror("读取失败", "无法读取规则文件：\n" + str(e))
            return
        dst, meta = core.persist_rule_memory(p, rule_count=n)
        if dst:
            log("已导入并记忆规则：{}（共 {} 条）→ {}\n".format(p, n, dst))
            refresh_rule_status()
            messagebox.showinfo("导入成功",
                                "规则已导入并记忆（共 {} 条），下次运行免重复选择。\n\n副本：\n{}".format(n, dst))
        else:
            messagebox.showerror("导入失败", "规则记忆写入失败，请检查软件存储目录权限。")

    def ask_save():
        p = filedialog.asksaveasfilename(
            title="选择结果输出路径",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="往来重分类分析结果.xlsx",
        )
        if p:
            var_out.set(p)

    def export_tpl():
        if state["running"]:
            return
        p = filedialog.asksaveasfilename(
            title="导出重分类规则模板",
            defaultextension=".xlsx",
            filetypes=[("Excel 文件", "*.xlsx")],
            initialfile="重分类规则模板.xlsx",
        )
        if p:
            try:
                core.export_rule_template(p)
                log("规则模板已导出：" + p + "\n")
                messagebox.showinfo("完成", "规则模板已导出：\n" + p)
            except Exception as e:
                messagebox.showerror("导出失败", str(e))

    def worker(data_dir, rules_path, out_path, src_files=None):
        buf = io.StringIO()
        try:
            with contextlib.redirect_stdout(buf), contextlib.redirect_stderr(buf):
                core.run_analysis(data_dir, rules_path, out_path, src_files or None)
            q.put(("done", buf.getvalue()))
        except Exception as e:
            q.put(("error", buf.getvalue() + "\n[错误] " + str(e) + "\n"))

    def poll():
        finished = False
        try:
            while True:
                kind, msg = q.get_nowait()
                log(msg)
                if kind in ("done", "error"):
                    finished = True
                    if kind == "done":
                        log("\n[完成] 分析结果已生成。\n")
                    else:
                        log("\n[失败] 请检查上方提示。\n")
        except queue.Empty:
            pass
        if finished:
            state["running"] = False
            btn_run.configure(text="执行分析", state="normal")
            return
        if state["running"]:
            root.after(100, poll)

    def run():
        if state["running"]:
            messagebox.showwarning("提示", "分析正在执行中，请稍候")
            return
        data_dir = var_data_dir.get().strip()
        if data_dir and not os.path.isdir(data_dir):
            messagebox.showerror("参数错误", "数据源目录不存在：" + data_dir)
            return
        # 参与分析的数据源文件 = 列表内已加载的全部文件（执行分析全部文件）
        explicit_paths = [ent.path for ent in src_model.entries]
        if explicit_paths and not all(os.path.isfile(p) for p in explicit_paths):
            messagebox.showerror("参数错误", "指定的数据源文件不存在，请重新选择")
            return
        # 解析最终数据源：复用列表加载时缓存的内容识别结果（stype -> path），
        # 已加载文件不再重复读表头；未显式覆盖的类型由目录内扫描兜底识别。
        explicit_map = {}
        if explicit_paths:
            for ent in src_model.entries:
                if ent.path in explicit_paths and ent.stype:
                    explicit_map[ent.stype] = ent.path
        src_paths = core.resolve_source_paths(data_dir, explicit_map or None)
        missing = []
        for cfg, stype in zip(core.SOURCE_CONFIG, core.SOURCE_TYPES):
            p = src_paths.get(stype)
            if not p or not os.path.isfile(p):
                missing.append(cfg["source"])
        if missing:
            messagebox.showerror(
                "参数错误",
                "以下数据源未就绪（请添加文件，或选择包含其的数据源目录）：\n"
                + "\n".join(missing),
            )
            return
        # 规则固定取自软件自身记忆（导入即固化）
        rules_path, auto_loaded, _ = core.resolve_rules_path("")
        out_path = var_out.get().strip()
        if not out_path:
            if data_dir and os.path.isdir(data_dir):
                out_path = os.path.join(data_dir, "往来重分类分析结果.xlsx")
            elif explicit_paths:
                out_path = os.path.join(
                    os.path.dirname(explicit_paths[0]), "往来重分类分析结果.xlsx"
                )
            else:
                messagebox.showerror("参数错误", "请选择输出路径")
                return
        if not rules_path or not os.path.isfile(rules_path):
            messagebox.showerror(
                "参数错误",
                "未找到可用的重分类规则。\n\n请点击「导入规则文件」导入规则，"
                "或点击「导出模板」生成模板并填写规则。",
            )
            return
        if auto_loaded:
            log("自动加载当前生效规则：" + rules_path + "\n")
        else:
            log("使用规则文件：" + rules_path + "\n")
        if explicit_paths:
            log("使用已加载的数据源文件（{} 个）：\n".format(len(explicit_paths)))
            for cfg, stype in zip(core.SOURCE_CONFIG, core.SOURCE_TYPES):
                p = src_paths.get(stype)
                if p:
                    log("    {} <- {}\n".format(cfg["source"], p))
        else:
            log("数据源文件（目录内按内容同类识别）：\n")
            for cfg, stype in zip(core.SOURCE_CONFIG, core.SOURCE_TYPES):
                p = src_paths.get(stype)
                if p:
                    log("    {} <- {}\n".format(cfg["source"], os.path.basename(p)))
        log("开始分析：\n")
        state["running"] = True
        btn_run.configure(text="分析中…", state="disabled")
        threading.Thread(
            target=worker,
            args=(data_dir, rules_path, out_path, explicit_map or None),
            daemon=True,
        ).start()
        root.after(100, poll)

    # ==================================================================
    # 绑定：按钮 / 拖拽 / 状态联动
    # ==================================================================
    btn_import.configure(command=import_rules_file)
    btn_export.configure(command=export_tpl)
    btn_add.configure(command=ask_source_files)
    btn_remove_sel.configure(command=_remove_selected)
    btn_clear.configure(command=_clear_src)
    btn_out_browse.configure(command=ask_save)
    btn_run.configure(command=run)

    # 数据源列表区 / 输出区 注册拖拽；窗口空白处兜底
    data_card.drop_target_register(DND_FILES)
    data_card.dnd_bind("<<Drop>>", handle_data_drop)
    root.drop_target_register(DND_FILES)
    root.dnd_bind("<<Drop>>", on_drop_anywhere)

    var_data_dir.trace_add("write", lambda *_: refresh_data_status())
    var_out.trace_add("write", lambda *_: None)

    log("欢迎使用往来重分类分析工具\n")
    log("提示：规则文件点击「导入规则文件」或拖入即可，导入后自动记忆；"
        "数据源可点击「添加文件」或拖入追加（可多选，自动识别余额/暂估/客户），"
        "列表选中行后「移除选中」，「清空列表」一键清空；执行分析处理全部已加载文件。\n")
    _render_src_list()
    refresh_rule_status()
    refresh_data_status()
    root.mainloop()


# ---------------------------------------------------------------------------
# 入口：带参数 -> CLI；无参数 -> GUI
# ---------------------------------------------------------------------------

def main():
    if len(sys.argv) > 1:
        cli_main()
    else:
        gui_main()


if __name__ == "__main__":
    main()
