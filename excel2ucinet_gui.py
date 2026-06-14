#!/usr/bin/env python3
"""
excel2ucinet_gui.py  —  舆情 Excel → Ucinet 转换工具（批量 GUI 版）

用法:
  python excel2ucinet_gui.py
"""

import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import threading
import os
import sys
import re
import csv
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    messagebox.showerror("缺少依赖", "请先安装 openpyxl:\npip install openpyxl")
    raise SystemExit(1)

# ───────────────────────── 核心逻辑（同 CLI 版） ─────────────────────────

ENTITY_RE = re.compile(r"(\S+?)（(\d+)）")


def parse_entity_column(text):
    if not text or not isinstance(text, str):
        return []
    return [(name.strip(), int(cnt)) for name, cnt in ENTITY_RE.findall(text)]


def read_excel(filepath):
    wb = openpyxl.load_workbook(filepath, read_only=True)
    ws = wb.active
    rows_entities = []
    for i, row in enumerate(ws.iter_rows(values_only=True), 1):
        if i == 1:
            continue
        if len(row) < 4:
            continue
        entities = parse_entity_column(row[3])
        if entities:
            rows_entities.append(entities)
    wb.close()
    return rows_entities


def filter_nodes(rows_entities, min_freq):
    freq = defaultdict(int)
    for entities in rows_entities:
        for name, _ in entities:
            freq[name] += 1
    keep = {n for n, f in freq.items() if f >= min_freq}
    result = []
    for entities in rows_entities:
        filtered = [(n, c) for n, c in entities if n in keep]
        if filtered:
            result.append(filtered)
    return result


def build_cooccurrence(rows_entities):
    node_count = defaultdict(int)
    co_matrix = defaultdict(int)
    for entities in rows_entities:
        names = sorted(set(e[0] for e in entities))
        for name in names:
            node_count[name] += 1
        for i in range(len(names)):
            for j in range(i + 1, len(names)):
                a, b = names[i], names[j]
                if a > b:
                    a, b = b, a
                co_matrix[(a, b)] += 1
    return dict(node_count), dict(co_matrix)


def write_ucinet_dl(node_count, co_matrix, filepath):
    nodes = sorted(node_count.keys())
    with open(filepath, "w", encoding="utf-8") as f:
        f.write("dl\n")
        f.write(f"n = {len(nodes)}\n")
        f.write("format = edgelist1\n")
        f.write("labels embedded\n")
        f.write("data:\n")
        for (a, b), w in sorted(co_matrix.items()):
            f.write(f"{a} {b} {w}\n")


def write_comatrix_csv(node_count, co_matrix, filepath):
    nodes = sorted(node_count.keys())
    idx = {n: i for i, n in enumerate(nodes)}
    n = len(nodes)
    mat = [[0] * n for _ in range(n)]
    for (a, b), w in co_matrix.items():
        i, j = idx[a], idx[b]
        mat[i][j] = w
        mat[j][i] = w
    with open(filepath, "w", newline="", encoding="utf-8-sig") as f:
        writer = csv.writer(f)
        writer.writerow([""] + nodes)
        for i, name in enumerate(nodes):
            writer.writerow([name] + mat[i])


# ───────────────────────── GUI ─────────────────────────

class App:
    def __init__(self, root):
        self.root = root
        root.title("舆情 Excel → Ucinet 转换工具（批量版）")
        root.geometry("720x620")
        root.minsize(640, 520)

        # ── 状态 ──
        self.file_list: list[str] = []          # 待处理的 xlsx 路径列表
        self.output_dir = tk.StringVar()
        self.min_freq = tk.IntVar(value=1)
        self.output_name = tk.StringVar(value="merged")

        self._build_ui()
        self._set_icon()

    # ────────── 图标 ──────────

    def _set_icon(self):
        try:
            if getattr(sys, 'frozen', False):
                base = sys._MEIPASS
            else:
                base = os.path.dirname(os.path.abspath(__file__))
            ico = os.path.join(base, 'icon.ico')
            if os.path.exists(ico):
                self.root.iconbitmap(default=ico)
        except Exception:
            pass

    # ────────── 界面构建 ──────────

    def _build_ui(self):
        # ── 文件选择 ──
        frame_file = tk.LabelFrame(self.root, text="输入文件（可多选）", padx=10, pady=10)
        frame_file.pack(fill="x", padx=10, pady=(10, 5))

        row_btn = tk.Frame(frame_file)
        row_btn.pack(fill="x", pady=2)
        tk.Button(row_btn, text="选择多个文件...", command=self._add_files,
                  width=16).pack(side="left", padx=(0, 5))
        tk.Button(row_btn, text="选择文件夹...", command=self._add_folder,
                  width=16).pack(side="left", padx=5)
        tk.Button(row_btn, text="清空列表", command=self._clear_files,
                  width=10).pack(side="right", padx=(5, 0))

        # 文件列表
        cols = ("文件名", "路径")
        self.tree_files = ttk.Treeview(frame_file, columns=cols, show="headings",
                                       height=5)
        self.tree_files.heading("文件名", text="文件名")
        self.tree_files.heading("路径", text="完整路径")
        self.tree_files.column("文件名", width=200)
        self.tree_files.column("路径", width=400)
        self.tree_files.pack(fill="x", pady=5)

        self.file_count_var = tk.StringVar(value="共 0 个文件")
        tk.Label(frame_file, textvariable=self.file_count_var,
                 fg="gray").pack(anchor="w")

        # ── 参数 ──
        frame_opts = tk.LabelFrame(self.root, text="转换参数", padx=10, pady=10)
        frame_opts.pack(fill="x", padx=10, pady=5)

        r1 = tk.Frame(frame_opts)
        r1.pack(fill="x", pady=3)
        tk.Label(r1, text="最低节点频次:").pack(side="left")
        tk.Spinbox(r1, from_=1, to=999, width=6,
                   textvariable=self.min_freq).pack(side="left", padx=5)
        tk.Label(r1, text="(出现次数低于此值的实体将被过滤)").pack(side="left", padx=(5, 0))

        r2 = tk.Frame(frame_opts)
        r2.pack(fill="x", pady=3)
        tk.Label(r2, text="输出文件名:").pack(side="left")
        tk.Entry(r2, textvariable=self.output_name, width=20).pack(side="left", padx=5)
        tk.Label(r2, text="（将生成 _ucinet_dl.txt 和 _comatrix.csv）",
                 fg="gray").pack(side="left")

        # ── 输出目录 ──
        frame_out = tk.LabelFrame(self.root, text="输出目录", padx=10, pady=10)
        frame_out.pack(fill="x", padx=10, pady=5)

        row_out = tk.Frame(frame_out)
        row_out.pack(fill="x", pady=2)
        tk.Entry(row_out, textvariable=self.output_dir, state="readonly").pack(
            side="left", fill="x", expand=True)
        tk.Button(row_out, text="选择目录...", command=self._browse_output,
                  width=12).pack(side="right", padx=(5, 0))
        tk.Button(row_out, text="当前目录", command=self._use_cwd,
                  width=10).pack(side="right", padx=(5, 0))

        # ── 日志 ──
        frame_log = tk.LabelFrame(self.root, text="处理日志", padx=5, pady=5)
        frame_log.pack(fill="both", expand=True, padx=10, pady=5)

        self.log_text = tk.Text(frame_log, height=8, state="disabled",
                                wrap="word", font=("Consolas", 9))
        scroll = tk.Scrollbar(frame_log, command=self.log_text.yview)
        self.log_text.configure(yscrollcommand=scroll.set)
        scroll.pack(side="right", fill="y")
        self.log_text.pack(fill="both", expand=True)

        # ── 底部按钮 ──
        btn_bar = tk.Frame(self.root)
        btn_bar.pack(fill="x", padx=10, pady=(5, 10))

        self.convert_btn = tk.Button(btn_bar, text="开始转换",
                                     command=self._convert,
                                     bg="#4CAF50", fg="white",
                                     width=14, font=("", 10, "bold"))
        self.convert_btn.pack(side="left")

        self.status_var = tk.StringVar(value="就绪")
        tk.Label(btn_bar, textvariable=self.status_var,
                 fg="gray").pack(side="right", padx=(10, 0))

        # ── 底部签名 ──
        tk.Label(self.root, text="by aLICE", fg="lightgray",
                 font=("", 8)).pack(side="bottom", anchor="sw", padx=12, pady=(0, 4))

    # ──────────

    def _add_files(self):
        paths = filedialog.askopenfilenames(
            title="选择 Excel 文件",
            filetypes=[("Excel 文件", "*.xlsx"), ("所有文件", "*.*")],
        )
        if paths:
            for p in paths:
                if p not in self.file_list:
                    self.file_list.append(p)
            self._refresh_file_list()

    def _add_folder(self):
        folder = filedialog.askdirectory(title="选择包含 xlsx 的文件夹")
        if folder:
            found = sorted(Path(folder).glob("*.xlsx"))
            added = 0
            for f in found:
                s = str(f)
                if s not in self.file_list:
                    self.file_list.append(s)
                    added += 1
            self._refresh_file_list()
            self.log(f"扫描文件夹: {folder}  → 新增 {added} 个 xlsx")

    def _clear_files(self):
        self.file_list.clear()
        self._refresh_file_list()

    def _refresh_file_list(self):
        for item in self.tree_files.get_children():
            self.tree_files.delete(item)
        for p in self.file_list:
            pp = Path(p)
            self.tree_files.insert("", "end", values=(pp.name, str(pp)))
        self.file_count_var.set(f"共 {len(self.file_list)} 个文件")

    def _browse_output(self):
        path = filedialog.askdirectory(title="选择输出目录")
        if path:
            self.output_dir.set(path)

    def _use_cwd(self):
        self.output_dir.set(os.getcwd())

    # ────────── 日志 ──────────

    def log(self, msg):
        self.log_text.configure(state="normal")
        self.log_text.insert("end", msg + "\n")
        self.log_text.see("end")
        self.log_text.configure(state="disabled")
        self.root.update_idletasks()

    # ────────── 转换 ──────────

    def _convert(self):
        if not self.file_list:
            messagebox.showwarning("提示", "请先添加需要转换的 xlsx 文件")
            return
        out = self.output_dir.get()
        if not out:
            out = os.getcwd()
            self.output_dir.set(out)

        self.convert_btn.config(state="disabled", text="转换中...")
        self.status_var.set("正在转换...")
        self.log("=" * 50)
        self.log("开始批量转换...")
        threading.Thread(target=self._do_convert, args=(out,), daemon=True).start()

    def _do_convert(self, out_dir):
        try:
            # 读取全部文件
            all_rows = []
            for p in self.file_list:
                rows = read_excel(p)
                self.log(f"  {Path(p).name}: {len(rows)} 行")
                all_rows.extend(rows)

            self.log(f"  合计: {len(all_rows)} 行数据")
            if not all_rows:
                self.root.after(0, lambda: self._done("未解析到实体"))
                return

            # 过滤
            mf = self.min_freq.get()
            if mf > 1:
                before = len(all_rows)
                all_rows = filter_nodes(all_rows, mf)
                self.log(f"  过滤后: {len(all_rows)} 行 (min-freq={mf})")

            # 构建
            nc, cm = build_cooccurrence(all_rows)
            self.log(f"  唯一实体数: {len(nc)}")
            self.log(f"  共现对数:   {len(cm)}")

            # 输出
            stem = self.output_name.get().strip() or "merged"
            dl_path = os.path.join(out_dir, f"{stem}_ucinet_dl.txt")
            csv_path = os.path.join(out_dir, f"{stem}_comatrix.csv")
            write_ucinet_dl(nc, cm, dl_path)
            write_comatrix_csv(nc, cm, csv_path)

            self.log(f"  >> DL:  {dl_path}")
            self.log(f"  >> CSV: {csv_path}")
            self.log("  完成！")

            msg = (f"转换完成！\n"
                   f"  文件数: {len(self.file_list)}\n"
                   f"  实体数: {len(nc)}\n"
                   f"  共现边: {len(cm)}\n\n"
                   f"输出:\n  {os.path.basename(dl_path)}\n  {os.path.basename(csv_path)}")
            self.root.after(0, lambda: self._done(msg, is_ok=True))

        except Exception as e:
            self.log(f"  !! 出错: {e}")
            self.root.after(0, lambda: self._done(f"出错: {e}"))

    def _done(self, msg, is_ok=False):
        self.convert_btn.config(state="normal", text="开始转换")
        self.status_var.set(msg.split("\n")[0])
        if is_ok:
            messagebox.showinfo("转换结果", msg)
        else:
            messagebox.showerror("转换失败", msg)


if __name__ == "__main__":
    root = tk.Tk()
    App(root)
    root.mainloop()
