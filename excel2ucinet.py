#!/usr/bin/env python3
"""
excel2ucinet.py  —  舆情 Excel → Ucinet DL edgelist + 共现矩阵 CSV（批量版）

功能:
  批量读取多个 xlsx（或整个文件夹），提取「内容」列中的实体关键词，
  合并构建实体共现网络，统一输出一个 Ucinet DL 文件和一个 CSV 共现矩阵。

输入格式:
  列4 (内容) : "实体名（次数） 实体名（次数）…"

用法:
  # 单个文件
  python excel2ucinet.py resource/example.xlsx

  # 多个文件（合并输出）
  python excel2ucinet.py a.xlsx b.xlsx c.xlsx

  # 通配符（shell 展开）
  python excel2ucinet.py resource/*.xlsx

  # 文件夹（扫描其中所有 xlsx）
  python excel2ucinet.py resource/  --batch-dir

  # 过滤低频节点 + 指定输出目录
  python excel2ucinet.py *.xlsx --min-freq 2 -o ./output/

  # 自定义输出文件名
  python excel2ucinet.py *.xlsx --name 鹅腿阿姨_合并
"""

import re
import os
import sys
import csv
import argparse
from collections import defaultdict
from pathlib import Path

try:
    import openpyxl
except ImportError:
    print("  !! 请先安装 openpyxl: pip install openpyxl")
    sys.exit(1)

# ───────────────────────── 核心逻辑 ─────────────────────────

ENTITY_RE = re.compile(r"(\S+?)（(\d+)）")


def parse_entity_column(text):
    """解析 "实体名（次数）" 格式文本，返回 [(name, count), ...]"""
    if not text or not isinstance(text, str):
        return []
    return [(name.strip(), int(cnt)) for name, cnt in ENTITY_RE.findall(text)]


def read_excel(filepath):
    """读取一个 xlsx，返回该文件中所有行的实体列表"""
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


def read_batch(paths):
    """批量读取多个 xlsx 文件，返回合并后的所有行实体列表"""
    all_rows = []
    for p in paths:
        p = Path(p)
        if not p.exists():
            print(f"  !! 跳过（不存在）: {p}")
            continue
        rows = read_excel(str(p))
        print(f"    {p.name}  → {len(rows)} 行")
        all_rows.extend(rows)
    return all_rows


def filter_nodes(rows_entities, min_freq):
    """按节点总频次过滤低频实体"""
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
    """构建共现矩阵，返回 (node_count, co_matrix)"""
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
    print(f"  >> Ucinet DL:  {filepath}")
    print(f"     节点数 {len(nodes)},  边数 {len(co_matrix)}")


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
    print(f"  >> 共现矩阵: {filepath}")
    print(f"     维度 {n} x {n}")


# ───────────────────────── 入口 ─────────────────────────


def main():
    ap = argparse.ArgumentParser(
        description="批量 Excel → Ucinet DL edgelist + 共现矩阵 CSV",
        formatter_class=argparse.RawDescriptionHelpFormatter,
        epilog=__doc__,
    )
    ap.add_argument("inputs", nargs="*", default=[],
                     help="xlsx 文件或目录（支持通配符）")
    ap.add_argument("--batch-dir", action="store_true",
                     help="将 inputs 中的目录展开为其中的所有 xlsx")
    ap.add_argument("-o", "--output-dir", default=None,
                     help="输出目录（默认：当前目录）")
    ap.add_argument("--name", default=None,
                     help="输出文件名前缀（默认: merged）")
    ap.add_argument("--min-freq", type=int, default=1,
                     help="节点最低出现频次（默认 1 = 不过滤）")
    args = ap.parse_args()

    # ── 收集文件列表 ──
    raw = args.inputs
    # 如果没有传参，尝试读取当前目录所有 xlsx
    if not raw:
        raw = ["."]

    files = []
    for item in raw:
        p = Path(item)
        if not p.exists():
            print(f"  !! 跳过（不存在）: {item}")
            continue
        if p.is_dir():
            if args.batch_dir:
                found = sorted(p.glob("*.xlsx"))
                files.extend(found)
                print(f"  扫描目录 {p}  → {len(found)} 个 xlsx")
            else:
                print(f"  !! 跳过目录（加 --batch-dir 展开）: {item}")
        elif p.suffix.lower() == ".xlsx":
            files.append(p)
        else:
            print(f"  !! 跳过（非 xlsx）: {item}")

    if not files:
        print("  !! 没有找到可处理的 xlsx 文件")
        sys.exit(1)

    # ── 批量读取 ──
    print(f"\n  批量读取 {len(files)} 个文件 ...")
    all_rows = read_batch(files)
    print(f"  合计: {len(all_rows)} 行数据")

    if not all_rows:
        print("  !! 未解析到任何实体")
        sys.exit(0)

    # ── 过滤 ──
    if args.min_freq > 1:
        before = len(all_rows)
        all_rows = filter_nodes(all_rows, args.min_freq)
        print(f"  过滤后: {len(all_rows)} 行 (min-freq={args.min_freq}, 过滤前 {before})")

    # ── 构建共现 ──
    node_count, co_matrix = build_cooccurrence(all_rows)
    print(f"  唯一实体数: {len(node_count)}")
    print(f"  共现对数:   {len(co_matrix)}\n")

    # ── 输出 ──
    out_dir = Path(args.output_dir) if args.output_dir else Path.cwd()
    out_dir.mkdir(parents=True, exist_ok=True)

    stem = args.name or "merged"
    write_ucinet_dl(node_count, co_matrix, out_dir / f"{stem}_ucinet_dl.txt")
    write_comatrix_csv(node_count, co_matrix, out_dir / f"{stem}_comatrix.csv")
    print("\n  >> 完成！")


if __name__ == "__main__":
    main()
