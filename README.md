# yuqing.nybkz-Ucinet

> **yuqing.nybkz 舆情系统转 Ucinet** —— 将舆情分析平台导出的 Excel 数据转换为 Ucinet DL edgelist 格式和共现矩阵 CSV。

---

## 功能

- 读取舆情系统导出的 Excel（格式：列4为实体关键词，`实体名（次数）`）
- 构建**实体共现网络**：同一行中出现的实体视为共现
- 批量处理多个 xlsx / 整个文件夹
- 输出两种格式：
  - **Ucinet DL edgelist** — 可直接导入 Ucinet / NetDraw 进行网络分析
  - **共现矩阵 CSV** — 实体×实体对称矩阵，适合导入其他分析工具
- 支持**低频节点过滤**（`--min-freq`）

---

## 使用方式

### 🖥 GUI 版（推荐）

双击 `excel2ucinet_gui.exe` 打开图形界面：

```
① 添加文件  →  选择多个 xlsx 或整个文件夹
② 设置参数  →  最低频次过滤、输出文件名
③ 开始转换  →  实时日志，完成后弹窗汇总
```

### ⌨️ CLI 版

```bash
# 单个文件
excel2ucinet.exe resource/example.xlsx

# 多个文件（合并输出）
excel2ucinet.exe a.xlsx b.xlsx c.xlsx

# 文件夹中所有 xlsx
excel2ucinet.exe resource/ --batch-dir

# 过滤低频节点 + 自定义输出名
excel2ucinet.exe *.xlsx --min-freq 2 --name 我的数据 --output-dir ./output/
```

### 🐍 从源码运行

```bash
pip install openpyxl

# CLI
python excel2ucinet.py resource/example.xlsx

# GUI
python excel2ucinet_gui.py
```

---

## 输入格式

Excel 文件需包含以下列（舆情系统标准导出）：

| 列 | 内容 | 示例 |
|----|------|------|
| 列4 | 实体关键词 | `微博（1） 荒漠（1） 北京（1）` |
| 列6 | 主话题（可选） | `鹅腿阿姨(3)` |

> 实体格式：`名称（次数）`，中文全角括号，多个实体以空格分隔。

---

## 输出示例

### Ucinet DL edgelist (`*_ucinet_dl.txt`)

```
dl
n = 23
format = edgelist1
labels embedded
data:
CAKE HTERABB 1
CAKE LOVEVOU2 1
北京 荒漠 2
...
```

### 共现矩阵 CSV (`*_comatrix.csv`)

| | 北京 | 微博 | 荒漠 |
|---|------|------|------|
| 北京 | 0 | 1 | 2 |
| 微博 | 1 | 0 | 1 |
| 荒漠 | 2 | 1 | 0 |

---

## 自行打包

```bash
pip install pyinstaller

# CLI
pyinstaller --onefile --name excel2ucinet excel2ucinet.py

# GUI（无控制台窗口）
pyinstaller --onefile --windowed --name excel2ucinet_gui excel2ucinet_gui.py
```

---

## 依赖

- Python ≥ 3.8
- openpyxl

---

## 许可证

MIT License © 2026 [Bili-Enderman](https://github.com/Bili-Enderman)
