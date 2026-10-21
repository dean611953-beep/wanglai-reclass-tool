#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
往来重分类分析工具 - 统一入口（GUI + CLI）
============================================
- 双击 .app / 无参数运行  -> 打开 tkinter 图形界面（紧凑卡片式布局）：
    * 顶部标题栏细窄；规则文件卡片置顶（独立"导入规则文件"按钮 + "已有规则"
      状态提示（条数 / 更新时间）+ "导出模板"）；
    * 数据源与输出结果两列并排紧凑分组；数据源支持整体选/拖入目录，或分别
      选/拖入 余额/暂估/客户 三个数据源单文件（文件名不限，按列结构同类识别），
      并实时显示三数据源就绪状态（已指定 / 目录识别 / 未就绪）；
    * 支持拖拽导入数据源 / 规则文件、规则记忆自动加载、执行分析、运行日志。
- 终端带参数运行          -> 命令行模式，完全兼容原 CLI 参数：
    --data-dir  数据源目录（默认当前目录）
    --rules     重分类规则文件（留空则自动加载上次导入并记忆的当前生效规则）
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
    root.geometry("880x700")
    root.minsize(780, 620)
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
    style.configure("Run.TButton",
                    background="#166534", foreground="#ffffff",
                    borderwidth=0, focusthickness=0, padding=(22, 6),
                    font=("PingFang SC", 11, "bold"))
    style.map("Run.TButton",
              background=[("pressed", "#14532d"), ("active", "#1f8a4c"),
                          ("disabled", "#a7c4b1")],
              foreground=[("disabled", "#e8f0ea")])

    var_data_dir = tk.StringVar()
    var_rules = tk.StringVar()
    var_out = tk.StringVar()
    # 数据源单入口：显式文件路径列表（可多选/拖入/手动输入，自动按内容同类识别归入三类）
    var_src_files = tk.StringVar()
    src_files = []              # 显式指定的数据源文件（余额/暂估/客户 按内容同类识别归入）
    _src_sync_guard = [False]   # 防递归：更新显示文本时禁止再次解析

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

    # ---------------- 1. 规则文件（置顶，全宽） ----------------
    rules_card = make_card(body, "规则文件",
                           "导入规则后自动记忆，下次运行免重复选择")
    rules_card.pack(fill="x", pady=(0, 6))

    rules_top = tk.Frame(rules_card, bg=CARD)
    rules_top.pack(fill="x", padx=12, pady=(0, 3))
    rules_status = tk.Label(rules_top, text="正在识别规则…", bg=CARD, fg=MUTED,
                            font=("PingFang SC", 9, "bold"), anchor="w")
    rules_status.pack(side="left")
    btn_export = ttk.Button(rules_top, text="导出模板", style="Ghost.TButton",
                            command=None)
    btn_import = ttk.Button(rules_top, text="导入规则文件", style="Accent.TButton",
                            command=None)
    btn_export.pack(side="right")
    btn_import.pack(side="right", padx=(0, 6))

    rules_row = tk.Frame(rules_card, bg=CARD)
    rules_row.pack(fill="x", padx=12, pady=(0, 9))
    ent_rules = tk.Entry(rules_row, textvariable=var_rules, bg="#f8fafc",
                         relief="solid", bd=1, highlightthickness=1,
                         highlightbackground=BORDER, highlightcolor=HEADER_ACCENT,
                         font=("PingFang SC", 10))
    ent_rules.pack(side="left", fill="x", expand=True, ipady=3)
    btn_rules_browse = ttk.Button(rules_row, text="浏览…", style="Ghost.TButton",
                                   command=None)
    btn_rules_browse.pack(side="left", padx=(6, 0))

    # ---------------- 2/3. 数据源 + 输出结果（两列并排） ----------------
    mid = tk.Frame(body, bg=BG)
    mid.pack(fill="x", pady=(0, 6))
    mid.columnconfigure(0, weight=3)
    mid.columnconfigure(1, weight=2)

    data_card = make_card(mid, "数据源",
                          "单个入口：文件 / 多选 / 目录 / 拖入，自动按内容同类识别")
    data_card.grid(row=0, column=0, sticky="nsew", padx=(0, 6))
    data_row = tk.Frame(data_card, bg=CARD)
    data_row.pack(fill="x", padx=12, pady=(0, 3))
    ent_data = tk.Entry(data_row, textvariable=var_src_files, bg="#f8fafc",
                        relief="solid", bd=1, highlightthickness=1,
                        highlightbackground=BORDER, highlightcolor=HEADER_ACCENT,
                        font=("PingFang SC", 10))
    ent_data.pack(side="left", fill="x", expand=True, ipady=3)
    btn_data_browse = ttk.Button(data_row, text="浏览…", style="Ghost.TButton",
                                 command=None)
    btn_data_browse.pack(side="left", padx=(6, 0))
    btn_data_dir = ttk.Button(data_row, text="目录…", style="Ghost.TButton",
                              command=None)
    btn_data_dir.pack(side="left", padx=(6, 0))

    data_status_box = tk.Frame(data_card, bg=CARD)
    data_status_box.pack(fill="x", padx=12, pady=(2, 6))
    data_status_labels = []
    for cfg in core.SOURCE_CONFIG:
        lb = tk.Label(data_status_box, text="○ " + cfg["source"], bg=CARD,
                      fg="#94a3b8", font=("PingFang SC", 8), anchor="w")
        lb.pack(fill="x")
        data_status_labels.append((cfg, lb))

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
    # 状态刷新：规则（条数/更新时间）与数据源（三文件就绪）
    # ==================================================================
    def _rule_count(path):
        try:
            return len(core.load_rules(path))
        except Exception:
            return 0

    def refresh_rule_status():
        explicit = var_rules.get().strip()
        data_dir = var_data_dir.get().strip()
        if explicit:
            if os.path.isfile(explicit):
                n = _rule_count(explicit)
                rules_status.config(
                    text="已指定规则文件 · 共 {} 条 · {}".format(n, os.path.basename(explicit)),
                    fg=OK)
            else:
                rules_status.config(text="规则文件不存在，请重新选择", fg=WARN)
        elif data_dir and os.path.isdir(data_dir):
            path, auto, meta = core.resolve_rules_path(data_dir, "")
            if path and os.path.isfile(path):
                n = _rule_count(path)
                ut = (meta or {}).get("update_time", "")
                if not ut:
                    try:
                        ut = datetime.datetime.fromtimestamp(
                            os.path.getmtime(path)).strftime("%Y-%m-%d %H:%M:%S")
                    except Exception:
                        ut = ""
                tail = (" · 更新于 " + ut) if ut else ""
                rules_status.config(
                    text="已有规则 ✓ 共 {} 条{}".format(n, tail), fg=OK)
            else:
                rules_status.config(
                    text="尚未记忆规则，请点击「导入规则文件」或导出模板填写", fg=WARN)
        else:
            rules_status.config(
                text="选择/拖入数据源目录后自动识别规则记忆", fg=MUTED)

    def refresh_data_status():
        data_dir = var_data_dir.get().strip()
        ok_dir = bool(data_dir) and os.path.isdir(data_dir)
        explicit_map = {}
        for p in src_files:
            cls = core.classify_source_file(p)
            if cls:
                explicit_map[cls[0]] = p
        dir_resolved = core.resolve_source_paths(data_dir) if ok_dir else {}
        for (cfg, lb), stype in zip(data_status_labels, core.SOURCE_TYPES):
            exp = explicit_map.get(stype, "")
            if exp and os.path.isfile(exp):
                text = "✓ " + cfg["source"] + " · 已指定 " + os.path.basename(exp)
                fg = OK
            else:
                dp = (dir_resolved or {}).get(stype)
                if dp and os.path.isfile(dp):
                    text = "✓ " + cfg["source"] + " · 目录识别 " + os.path.basename(dp)
                    fg = OK
                else:
                    text = "○ " + cfg["source"] + " · 未就绪（选文件或选目录）"
                    fg = "#94a3b8"
            lb.config(text=text, fg=fg)

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
        """拖入数据源目录 / 一个或多个数据源 Excel：按内容同类识别自动归入三类。"""
        paths = parse_drop_data(event.data)
        dirs = [p for p in paths if os.path.isdir(p)]
        files = [p for p in paths if os.path.isfile(p)]
        added = False
        if dirs:
            var_data_dir.set(dirs[0])
            log("[拖拽] 数据源目录已设置为: " + dirs[0] + "\n")
            added = True
        if files:
            matched = add_src_files(files)
            if matched:
                common = os.path.dirname(matched[0])
                if all(os.path.dirname(p) == common for p in matched) and os.path.isdir(common):
                    var_data_dir.set(common)
                    log("[拖拽] 数据源目录已自动设置为: " + common + "\n")
                added = True
        if not added:
            log("[拖拽] 未识别到可用的数据源，请拖入 数据源目录，或 余额/暂估/客户 三类 xlsx "
                "（文件名不限，按列结构自动识别）\n")

    def handle_rules_drop(event):
        """拖入规则文件：填入规则框。"""
        paths = parse_drop_data(event.data)
        xlsx = [p for p in paths if os.path.isfile(p) and p.lower().endswith(".xlsx")]
        if xlsx:
            var_rules.set(xlsx[0])
            log("[拖拽] 规则文件已设置为: " + xlsx[0] + "\n")
        else:
            log("[拖拽] 未识别到 Excel 规则文件（请拖入 .xlsx 规则表）\n")

    def on_drop_anywhere(event):
        """窗口空白处兜底：按内容自动分发（数据源同类识别优先，规则关键词次之，单文件兜底规则）。"""
        paths = parse_drop_data(event.data)
        files = [p for p in paths if os.path.isfile(p)]
        if files and any(core.classify_source_file(p) for p in files):
            handle_data_drop(event)
        elif files and any("规则" in os.path.basename(p) for p in files):
            handle_rules_drop(event)
        elif len(files) == 1:
            handle_rules_drop(event)
        else:
            handle_data_drop(event)

    # ==================================================================
    # 对话框动作
    # ==================================================================
    def ask_dir():
        p = filedialog.askdirectory(title="选择数据源目录")
        if p:
            var_data_dir.set(p)

    def add_src_files(paths):
        """将显式数据源文件加入单入口列表：按内容同类识别归入三类，去重并刷新显示。
        返回成功加入（可识别）的文件路径列表。"""
        matched = []
        for p in paths:
            if not p or not os.path.isfile(p):
                continue
            if p in src_files:
                continue
            cls = core.classify_source_file(p)
            if cls:
                stype, cfg = cls
                src_files.append(p)
                log("[识别] 数据源 {} <- {}\n".format(cfg["source"], p))
                matched.append(p)
            else:
                log("[提示] 忽略无法识别的文件（未匹配到余额/暂估/客户同类结构）: {}\n".format(p))
        if matched:
            _src_sync_guard[0] = True
            try:
                var_src_files.set("; ".join(src_files))
            finally:
                _src_sync_guard[0] = False
        return matched

    def _on_src_display_change(*_args):
        """输入框内容变化（手动输入/粘贴）：按 ; 或换行解析为显式文件路径列表。"""
        if _src_sync_guard[0]:
            return
        text = var_src_files.get()
        cands = [s.strip() for s in text.replace(";", "\n").split("\n") if s.strip()]
        kept = []
        for c in cands:
            if os.path.isfile(c):
                cls = core.classify_source_file(c)
                if cls:
                    kept.append(c)
        if kept != src_files:
            src_files[:] = kept
        refresh_data_status()

    def ask_source_files():
        """浏览并选择数据源文件（支持单选/多选：余额/暂估/客户 任意组合）。"""
        ps = filedialog.askopenfilenames(
            title="选择数据源文件（可多选：余额/暂估/客户 三类自动识别）",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if ps:
            matched = add_src_files(list(ps))
            if not matched:
                messagebox.showerror(
                    "无法识别",
                    "所选文件未匹配到余额/暂估/客户任一数据源结构（列结构特征），请确认文件正确。",
                )

    def ask_file():
        p = filedialog.askopenfilename(
            title="选择重分类规则文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if p:
            var_rules.set(p)

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

    q = queue.Queue()
    state = {"running": False}

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
        # 显式指定的数据源文件（单入口：可多选/拖入；未覆盖的类型从目录内按内容同类识别）
        explicit_paths = list(src_files)
        if explicit_paths and not all(os.path.isfile(p) for p in explicit_paths):
            messagebox.showerror("参数错误", "指定的数据源文件不存在，请重新选择")
            return
        # 解析最终数据源：显式文件优先，其余目录内识别
        src_paths = core.resolve_source_paths(data_dir, explicit_paths or None)
        missing = []
        for cfg, stype in zip(core.SOURCE_CONFIG, core.SOURCE_TYPES):
            p = src_paths.get(stype)
            if not p or not os.path.isfile(p):
                missing.append(cfg["source"])
        if missing:
            messagebox.showerror(
                "参数错误",
                "以下数据源未就绪（请分别选择文件，或选择包含其的数据源目录）：\n"
                + "\n".join(missing),
            )
            return
        rules_path, auto_loaded, _ = core.resolve_rules_path(
            data_dir, var_rules.get().strip()
        )
        out_path = var_out.get().strip()
        if not out_path:
            if data_dir and os.path.isdir(data_dir):
                out_path = os.path.join(data_dir, "往来重分类分析结果.xlsx")
            else:
                out_path = os.path.join(
                    os.path.dirname(explicit_paths[0]), "往来重分类分析结果.xlsx"
                )
        if not rules_path or not os.path.isfile(rules_path):
            messagebox.showerror(
                "参数错误",
                "未找到可用的重分类规则。\n\n可先点击「导入规则文件」导入规则，"
                "或点击「导出模板」生成模板并填写规则。",
            )
            return
        if auto_loaded:
            log("自动加载当前生效规则：" + rules_path + "\n")
        else:
            log("使用规则文件：" + rules_path + "\n")
        if explicit_paths:
            log("使用显式指定的数据源文件（{} 个）：\n".format(len(explicit_paths)))
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
            args=(data_dir, rules_path, out_path, list(explicit_paths)),
            daemon=True,
        ).start()
        root.after(100, poll)

    # ==================================================================
    # 绑定：按钮 / 拖拽 / 状态联动
    # ==================================================================
    btn_import.configure(command=ask_file)
    btn_export.configure(command=export_tpl)
    btn_rules_browse.configure(command=ask_file)
    btn_data_browse.configure(command=ask_source_files)
    btn_data_dir.configure(command=ask_dir)
    btn_out_browse.configure(command=ask_save)
    btn_run.configure(command=run)

    ent_data.drop_target_register(DND_FILES)
    ent_data.dnd_bind("<<Drop>>", handle_data_drop)
    ent_rules.drop_target_register(DND_FILES)
    ent_rules.dnd_bind("<<Drop>>", handle_rules_drop)
    root.drop_target_register(DND_FILES)
    root.dnd_bind("<<Drop>>", on_drop_anywhere)

    var_data_dir.trace_add("write", lambda *_: (refresh_data_status(), refresh_rule_status()))
    var_rules.trace_add("write", lambda *_: refresh_rule_status())
    var_src_files.trace_add("write", _on_src_display_change)

    log("欢迎使用往来重分类分析工具\n")
    log("提示：单个入口即可选择/拖入 数据源目录 或 余额/暂估/客户 数据源文件（可多选、"
        "文件名不限，自动按列结构识别归入三类）；规则 xlsx 可拖入规则框；也可拖到窗口任意位置自动识别。\n")
    refresh_data_status()
    refresh_rule_status()
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
