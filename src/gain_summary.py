# -*- coding: utf-8 -*-
r"""
光谱仪测试数据「按需汇总」工具

用途
----
仪器每测一次导出一个文本文件，一天几十个。手工抄进 Excel 慢且容易错位。
本工具把这件事自动化：给文件（或文件夹），出统计表。

工作流程
--------
1. 解析仪器导出文件 —— 表头按文字匹配，列顺序变了也能读；
2. 可选手写参数表（样品编号、用纤长度、输入输出功率等）补进统计；
3. 每次运行产出一份多 sheet 的 Excel（统计 / 明细 / 说明），并追加到累计 CSV；
4. 可选把数据回填到自己的模板副本，原件只读。

用法
----
    python src/gain_summary.py data/sample/demo_001.csv
    python src/gain_summary.py data/sample/
    python src/gain_summary.py data/sample/ --params data/sample/params.csv
    python src/gain_summary.py --dry-run          # 只解析打印，不落盘
    python src/gain_summary.py data/sample/ --fill-template

设计约定
--------
- 仪器已算好的指标（增益、噪声系数）只读取、不重算，避免引入算法偏差；
- 数据点数不写死，从文件推断，换仪器配置也能跑；
- 型号前缀写在配置区，不绑定任何具体产品线。

示例数据说明
------------
data/sample/ 下的文件由脚本随机生成，仅用于验证代码逻辑，数值无工程意义。
"""

from __future__ import annotations

# ==================== 配置区（一般只改这里） ====================

# 不传任何目标时，默认扫描的目录（相对本文件所在目录）
SCAN_SUBDIR = "data/sample"

# 输出 / 日志目录
OUTPUT_SUBDIR = "output"
LOG_SUBDIR = "logs"

# 回填模板（--fill-template 时才用到；只读，写的是副本）
TEMPLATE_FILE = "templates/demo_template.xlsx"

# 样品型号识别：按自己的命名习惯改这两个正则
# 例：文件夹名 DEMO-001 或 FIBER-003-300M
MODEL_PATTERN = r"(DEMO|FIBER)[-_](\d+)"
LENGTH_PATTERN = r"[-_](\d+)\s*M\b"

# 判重状态文件，仅目录扫描模式使用
STATE_FILENAME = ".summary_state.json"

# 点数兜底校验：设成整数则强制校验，设成 None 表示不校验（推荐 None）
FALLBACK_POINTS = None

# 小数位
GAIN_DIGITS = 3
WL_DIGITS = 4

# ===============================================================

import argparse
import csv
import hashlib
import json
import re
import sys
import traceback
from datetime import datetime
from pathlib import Path

import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

BASE_DIR = Path(__file__).resolve().parent.parent   # 项目根目录
OUTPUT_DIR = BASE_DIR / OUTPUT_SUBDIR
LOG_DIR = BASE_DIR / LOG_SUBDIR
STATE_FILE = BASE_DIR / STATE_FILENAME

_LOG_PATH: Path | None = None


def log(message: str = "") -> None:
    """同时打印到控制台和当天日志文件。"""
    print(message)
    if _LOG_PATH is not None:
        with _LOG_PATH.open("a", encoding="utf-8") as fh:
            fh.write(message + "\n")


# ---------------------------------------------------------------- 手写参数

PARAM_FIELDS = (
    "样品编号",
    "长度m",
    "输入信号光功率dBm",
    "输出光功率dBm",
    "熔接方向",
    "起始波长nm",
    "泵浦功率mW",
)

PARAM_COLUMNS = {
    "输入信号光功率dBm": "手写-输入光功率dBm",
    "输出光功率dBm": "手写-输出光功率dBm",
    "熔接方向": "手写-熔接方向",
    "起始波长nm": "手写-起始波长nm",
    "泵浦功率mW": "手写-泵浦功率mW",
}


def load_params(path: Path) -> dict:
    """
    读手写参数表（CSV，UTF-8 或 GBK 均可）。

    写法：
        文件名,样品编号,长度m,输入信号光功率dBm,输出光功率dBm,熔接方向,起始波长nm,泵浦功率mW
        demo_001.csv,DEMO-001,300,-5.0,8.0,前向,1550,500

    「文件名」和「文件夹」两列任给其一，用来对应到具体数据文件。
    """
    text = None
    for encoding in ("utf-8-sig", "gbk"):
        try:
            text = path.read_text(encoding=encoding)
            break
        except UnicodeDecodeError:
            continue
    if text is None:
        raise ValueError(f"参数表编码无法识别：{path.name}")

    table: dict[str, dict] = {}
    for record in csv.DictReader(text.splitlines()):
        cleaned = {k: (record.get(k) or "").strip() for k in PARAM_FIELDS}
        for key_column in ("文件名", "文件夹"):
            key = (record.get(key_column) or "").strip()
            if key:
                table[key] = cleaned
    return table


def lookup_params(table: dict, path: Path) -> dict:
    """按文件名优先、文件夹名次之，取该文件对应的手写参数。"""
    if not table:
        return {}
    return table.get(path.name) or table.get(path.parent.name) or {}


def as_int(value) -> int | None:
    """把 '400' / '400.0' 这类写法稳妥转成 int，转不了返回 None。"""
    if value in (None, ""):
        return None
    try:
        return int(float(str(value)))
    except (TypeError, ValueError):
        return None


# ---------------------------------------------------------------- 解析

# 表头别名：不同仪器型号/固件命名略有差异，按前缀匹配
COLUMN_ALIASES = {
    "wavelength": ("WAVELENGTH", "波长"),
    "input": ("INPUTLVL", "INPUT"),
    "output": ("OUTPUTLVL", "OUTPUT"),
    "ase": ("ASELVL", "ASE"),
    "resoln": ("RESOLN", "RESOLUTION"),
    "gain": ("GAIN", "增益"),
    "nf": ("NF", "噪声系数"),
}


def _normalize(name: str) -> str:
    """去掉空格、转大写，便于宽松匹配表头。"""
    return re.sub(r"\s+", "", name).upper()


def _locate_columns(header_line: str) -> dict[str, int]:
    """按表头文字定位各字段的列号，返回 {字段: 列索引}。"""
    names = [_normalize(part) for part in header_line.split(",")]
    mapping: dict[str, int] = {}
    for field, prefixes in COLUMN_ALIASES.items():
        for index, name in enumerate(names):
            if any(name.startswith(prefix) for prefix in prefixes):
                mapping[field] = index
                break
    missing = [f for f in ("wavelength", "gain") if f not in mapping]
    if missing:
        raise ValueError(f"表头缺少必需字段 {missing}：{header_line[:80]}")
    return mapping


def _read_text(path: Path) -> str:
    """仪器导出文件编码可能是 UTF-8(BOM) 或 GBK，双兜底。"""
    for encoding in ("utf-8-sig", "gbk", "utf-8"):
        try:
            return path.read_text(encoding=encoding)
        except UnicodeDecodeError:
            continue
    raise ValueError(f"编码无法识别：{path.name}")


def parse_data_file(path: Path) -> dict:
    """
    解析单个仪器导出文件。

    返回：
        {
          "path": Path,
          "measured_at": str | None,          # 文件里的测试时间
          "points": [{"wavelength":..,"gain":..,"nf":..,...}, ...],
        }
    """
    lines = _read_text(path).splitlines()

    # 数据表头固定以 "NO.," 开头，用来定位数据区起点
    header_index = next(
        (i for i, line in enumerate(lines) if line.strip().upper().startswith("NO.,")),
        None,
    )
    if header_index is None:
        raise ValueError(f"找不到 'NO.,' 表头行：{path.name}")

    columns = _locate_columns(lines[header_index])

    points = []
    for line in lines[header_index + 1:]:
        parts = line.strip().split(",")
        if len(parts) <= max(columns.values()):
            continue
        try:
            point = {
                "wavelength": float(parts[columns["wavelength"]]),
                "gain": float(parts[columns["gain"]]),
            }
            if "nf" in columns:
                point["nf"] = float(parts[columns["nf"]])
            if "input" in columns:
                point["input"] = float(parts[columns["input"]])
            if "output" in columns:
                point["output"] = float(parts[columns["output"]])
        except (ValueError, IndexError):
            # 参数行、尾部空行会落到这里
            continue
        points.append(point)

    if not points:
        raise ValueError(f"没有解析到任何数据点：{path.name}")
    if FALLBACK_POINTS is not None and len(points) != FALLBACK_POINTS:
        raise ValueError(
            f"{path.name} 数据点数为 {len(points)}，应为 {FALLBACK_POINTS}，请人工确认"
        )

    # 测试时间通常在前 5 行，形如 "2026 Oct 17 15:46"
    measured_at = None
    for line in lines[:5]:
        if re.search(r"\d{4}\s+[A-Za-z]{3}\s+\d{1,2}\s+\d{1,2}:\d{2}", line):
            measured_at = line.strip()
            break

    return {"path": path, "measured_at": measured_at, "points": points}


# ---------------------------------------------------------------- 样品编号识别

REEL_PATTERN = re.compile(MODEL_PATTERN, re.IGNORECASE)
LENGTH_RE = re.compile(LENGTH_PATTERN, re.IGNORECASE)


def identify_reel(path: Path) -> dict:
    """
    从路径里识别型号 / 编号 / 用纤长度。

    约定：文件夹名形如 DEMO-001 或 FIBER-003-300M。
    长度识别不到不算错误，手写参数表里的「长度m」会补上。
    """
    for part in reversed(path.parts):
        match = REEL_PATTERN.search(part)
        if match:
            model = match.group(1).upper()
            number = match.group(2)
            length_match = LENGTH_RE.search(part)
            return {
                "model": model,
                "reel": number,
                "length_m": int(length_match.group(1)) if length_match else None,
                "header": (
                    f"{model}-{number}-{length_match.group(1)}M" if length_match else None
                ),
            }
    return {"model": None, "reel": None, "length_m": None, "header": None}


def resolve_identity(path: Path, params: dict) -> dict:
    """
    综合「文件夹名」和「手写参数表」定出 样品编号 / 长度 / 模板列头。
    手写参数优先，因为它才是现场记录的真实口径。
    """
    detected = identify_reel(path)

    params_reel = params.get("样品编号")
    params_length = as_int(params.get("长度m"))

    if params_reel:
        display = params_reel
    elif detected["reel"]:
        display = f"{detected['model']}-{detected['reel']}"
    else:
        display = "未识别"

    length = params_length if params_length is not None else detected["length_m"]
    header = f"{display}-{length}M" if display != "未识别" and length else detected["header"]

    return {"display": display, "length_m": length, "header": header}


# ---------------------------------------------------------------- 统计

def summarize(parsed: dict) -> dict:
    """
    按解析出的全部点计算关键指标。

    口径说明：
      - 平均值 / 极值都直接在原始点上算，不插值；
      - 定点波长取「最近点」，容差 ±1 nm，超出记空白；
      - 平坦度 = 全程最大增益 − 最小增益。
    """
    points = parsed["points"]
    gains = [p["gain"] for p in points]
    wavelengths = [p["wavelength"] for p in points]

    peak_index = max(range(len(gains)), key=lambda i: gains[i])
    dip_index = min(range(len(gains)), key=lambda i: gains[i])

    def gain_at(target: float) -> float | None:
        nearest = min(range(len(wavelengths)), key=lambda i: abs(wavelengths[i] - target))
        return gains[nearest] if abs(wavelengths[nearest] - target) <= 1.0 else None

    nfs = [p.get("nf") for p in points if p.get("nf") is not None]
    inputs = [p.get("input") for p in points if p.get("input") is not None]
    outputs = [p.get("output") for p in points if p.get("output") is not None]

    return {
        "点数": len(points),
        "起始波长nm": round(wavelengths[0], WL_DIGITS),
        "终止波长nm": round(wavelengths[-1], WL_DIGITS),
        "平均增益dB": round(sum(gains) / len(gains), GAIN_DIGITS),
        "最大增益dB": round(gains[peak_index], GAIN_DIGITS),
        "最大增益@nm": round(wavelengths[peak_index], WL_DIGITS),
        "最小增益dB": round(gains[dip_index], GAIN_DIGITS),
        "最小增益@nm": round(wavelengths[dip_index], WL_DIGITS),
        "增益平坦度dB": round(gains[peak_index] - gains[dip_index], GAIN_DIGITS),
        "1530nm增益dB": _r(gain_at(1530.0), GAIN_DIGITS),
        "1550nm增益dB": _r(gain_at(1550.0), GAIN_DIGITS),
        "1626nm增益dB": _r(gain_at(1626.0), GAIN_DIGITS),
        "平均NFdB": _r(sum(nfs) / len(nfs) if nfs else None, GAIN_DIGITS),
        "最小NFdB": _r(min(nfs) if nfs else None, GAIN_DIGITS),
        "输入功率均值dBm": _r(sum(inputs) / len(inputs) if inputs else None, 3),
        "输出功率均值dBm": _r(sum(outputs) / len(outputs) if outputs else None, 3),
    }


def _r(value, digits):
    return round(value, digits) if value is not None else None


# ---------------------------------------------------------------- 判重状态（仅扫描模式）

def file_fingerprint(path: Path) -> dict:
    """用大小 + 修改时间 + SHA1 给文件做指纹，用于判重。"""
    stat = path.stat()
    digest = hashlib.sha1(path.read_bytes()).hexdigest()
    return {"size": stat.st_size, "mtime": round(stat.st_mtime, 3), "sha1": digest}


def load_state() -> dict:
    if STATE_FILE.is_file():
        try:
            return json.loads(STATE_FILE.read_text(encoding="utf-8"))
        except (json.JSONDecodeError, OSError):
            log(f"警告：状态文件损坏，按全新处理：{STATE_FILE.name}")
    return {"files": {}}


def save_state(state: dict) -> None:
    STATE_FILE.write_text(
        json.dumps(state, ensure_ascii=False, indent=2), encoding="utf-8"
    )


def collect_files(roots: list[Path]) -> list[Path]:
    """递归收集所有 D 开头的 CSV（也兼容 .D / .TXT 导出）。"""
    found: list[Path] = []
    for root in roots:
        if not root.is_dir():
            log(f"警告：目录不存在，已跳过：{root}")
            continue
        for path in root.rglob("*"):
            if not path.is_file():
                continue
            if path.suffix.upper() not in (".CSV", ".D", ".TXT"):
                continue
            if not path.stem.upper().startswith("D"):
                continue
            found.append(path)
    return sorted(set(found), key=lambda p: p.stat().st_mtime)


# ---------------------------------------------------------------- 输出：Excel

HEADER_FILL = PatternFill("solid", fgColor="DDEBF7")
TITLE_FONT = Font(bold=True, size=12)

STAT_COLUMNS = [
    "样品编号", "用纤长度/m", "文件名", "测试时间", "点数",
    "起始波长nm", "终止波长nm",
    "平均增益dB", "最大增益dB", "最大增益@nm",
    "最小增益dB", "最小增益@nm", "增益平坦度dB",
    "1530nm增益dB", "1550nm增益dB", "1626nm增益dB",
    "平均NFdB", "最小NFdB", "输入功率均值dBm", "输出功率均值dBm",
    "手写-输入光功率dBm", "手写-输出光功率dBm", "手写-熔接方向",
    "手写-起始波长nm", "手写-泵浦功率mW",
    "文件路径",
]


def _style_header(ws, width_by_col: dict[str, int]) -> None:
    for cell in ws[1]:
        cell.font = Font(bold=True)
        cell.fill = HEADER_FILL
        cell.alignment = Alignment(horizontal="center", vertical="center")
    for letter, width in width_by_col.items():
        ws.column_dimensions[letter].width = width
    ws.freeze_panes = "A2"


def write_workbook(rows: list[dict], run_time: datetime, params_source: str | None) -> Path:
    """写出本次运行的汇总工作簿：统计 + 明细 + 说明。"""
    wb = openpyxl.Workbook()

    # ---- 统计表（放第一个，最常看）
    ws_stat = wb.active
    ws_stat.title = "统计"
    ws_stat.append(STAT_COLUMNS)
    for row in rows:
        ws_stat.append([row.get(col) for col in STAT_COLUMNS])
    _style_header(
        ws_stat,
        {"A": 14, "B": 12, "C": 16, "D": 20, "E": 8, "F": 12, "G": 12,
         "H": 12, "I": 12, "J": 14, "K": 12, "L": 14, "M": 14,
         "N": 14, "O": 14, "P": 14, "Q": 11, "R": 11, "S": 16, "T": 16,
         "U": 18, "V": 18, "W": 14, "X": 16, "Y": 16, "Z": 60},
    )

    # ---- 明细表
    ws_detail = wb.create_sheet("明细")
    ws_detail.append(["样品编号", "文件名", "波长nm", "增益dB", "NF dB", "输入dBm", "输出dBm"])
    for row in rows:
        for point in row["数据点"]:
            ws_detail.append([
                row.get("样品编号"),
                row["文件名"],
                round(point["wavelength"], WL_DIGITS),
                round(point["gain"], GAIN_DIGITS),
                _r(point.get("nf"), GAIN_DIGITS),
                _r(point.get("input"), 3),
                _r(point.get("output"), 3),
            ])
    _style_header(
        ws_detail,
        {"A": 14, "B": 16, "C": 12, "D": 10, "E": 10, "F": 12, "G": 12},
    )

    # ---- 说明表：把口径写清楚，别人拿到表不用问
    ws_note = wb.create_sheet("说明")
    notes = [
        ["光谱数据汇总"],
        [],
        ["运行时间", run_time.strftime("%Y-%m-%d %H:%M:%S")],
        ["处理文件数", len(rows)],
        ["手写参数表", params_source or "未提供"],
        ["脚本", Path(__file__).name],
        [],
        ["数据口径"],
        ["增益 / NF 均为仪器在原始文件内直接输出，脚本只读取，不重算、不插值。"],
        ["1530 / 1550 / 1626 nm 取波长最近点，容差 ±1 nm；超出容差记为空白。"],
        ["增益平坦度 = 全程最大增益 − 最小增益。"],
        ["「手写-」开头的列来自手写参数表，未经脚本核算。"],
        [],
        ["注意"],
        ["本文件由脚本生成；补数据请重新运行脚本，不要手工覆盖。"],
    ]
    for line in notes:
        ws_note.append(line)
    ws_note["A1"].font = TITLE_FONT
    ws_note.column_dimensions["A"].width = 24
    ws_note.column_dimensions["B"].width = 60

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    target = OUTPUT_DIR / f"汇总_{run_time:%Y-%m-%d_%H%M%S}.xlsx"
    wb.save(target)
    return target


def append_cumulative_csv(rows: list[dict], run_time: datetime) -> Path:
    """累计统计表：每次运行按「文件名+SHA1」去重后追加，便于长期看趋势。"""
    target = OUTPUT_DIR / "累计统计.csv"
    known: set[str] = set()
    if target.is_file():
        with target.open("r", encoding="utf-8-sig", newline="") as fh:
            for record in csv.DictReader(fh):
                known.add(f"{record.get('文件名')}|{record.get('SHA1')}")

    columns = [
        "运行时间", "样品编号", "用纤长度/m", "文件名", "测试时间",
        "平均增益dB", "最大增益dB", "最小增益dB", "增益平坦度dB",
        "1530nm增益dB", "1626nm增益dB", "平均NFdB", "SHA1", "文件路径",
    ]
    is_new = not target.is_file()
    with target.open("a", encoding="utf-8-sig", newline="") as fh:
        writer = csv.DictWriter(fh, fieldnames=columns, extrasaction="ignore")
        if is_new:
            writer.writeheader()
        for row in rows:
            key = f"{row['文件名']}|{row.get('SHA1')}"
            if key in known:
                continue
            record = {"运行时间": run_time.strftime("%Y-%m-%d %H:%M")}
            record.update({col: row.get(col) for col in columns})
            writer.writerow(record)
            known.add(key)
    return target


# ---------------------------------------------------------------- 输出：模板回填

def fill_template(rows: list[dict], run_time: datetime, template_path: Path) -> Path | None:
    """
    把数据回填到模板的「汇总」sheet。

    安全约定：
      - 只读模板原件，写的是副本 <模板名>_回填_<日期>.xlsx；
      - 按第 1 行列头精确匹配（如 DEMO-001-300M），匹配不上就跳过并报告；
      - 模板原有的公式保留不覆盖；
      - 参数区不自动写，那部分口径需现场确认，脚本不猜。
    """
    template = template_path
    if not template.is_file():
        log(f"警告：模板不存在，跳过回填：{template}")
        return None

    wb = openpyxl.load_workbook(template)
    if "汇总" not in wb.sheetnames:
        log(f"警告：模板里没有「汇总」sheet，现有：{wb.sheetnames}")
        return None
    ws = wb["汇总"]

    # 扫描第 1 行，建立「列头文字 → 列号」
    header_to_col: dict[str, int] = {}
    for col in range(1, ws.max_column + 1):
        value = ws.cell(row=1, column=col).value
        if isinstance(value, str) and value.strip():
            header_to_col.setdefault(value.strip(), col)

    filled, skipped = 0, []
    for row in rows:
        header = row.get("模板列头")
        col = header_to_col.get(header) if header else None

        # 回退：文件夹没写长度时，若模板里同编号列头唯一，就用它
        if col is None and row.get("样品编号"):
            candidates = [n for n in header_to_col if n.startswith(f"{row['样品编号']}-")]
            if len(candidates) == 1:
                col = header_to_col[candidates[0]]
                log(f"  提示：{row['文件名']} 用唯一列头「{candidates[0]}」回填")
            elif len(candidates) > 1:
                skipped.append(
                    f"{row['文件名']}：同编号有 {len(candidates)} 个列头 "
                    f"（{', '.join(sorted(candidates))}），长度不明确，请在参数表里补「长度m」"
                )
                continue

        if col is None:
            skipped.append(f"{row['文件名']}：模板第 1 行没有列头「{header}」")
            continue

        for offset, point in enumerate(row["数据点"]):
            ws.cell(row=2 + offset, column=col, value=round(point["gain"], GAIN_DIGITS))
        filled += 1

    OUTPUT_DIR.mkdir(parents=True, exist_ok=True)
    target = OUTPUT_DIR / f"{template.stem}_回填_{run_time:%Y-%m-%d_%H%M%S}.xlsx"
    wb.save(target)

    log(f"模板回填：成功 {filled} 个，跳过 {len(skipped)} 个")
    for message in skipped:
        log(f"  跳过 → {message}")
    log("提醒：openpyxl 保存会丢失模板里的散点图数据引用，图需在 Excel 里重新指定数据源。")
    return target


# ---------------------------------------------------------------- 主流程

def main() -> int:
    parser = argparse.ArgumentParser(
        description="光谱数据按需汇总：给文件或文件夹，出 Excel 统计表",
        epilog="示例：python src/gain_summary.py data/sample/demo_001.csv",
    )
    parser.add_argument("targets", nargs="*", metavar="文件或目录",
                        help="要汇总的数据文件或文件夹，可以给多个")
    parser.add_argument("--params", "--手写参数", dest="params_file", metavar="CSV",
                        help="手写参数表（样品编号、长度、输入输出光功率等）")
    parser.add_argument("--scan", action="append", metavar="DIR",
                        help="按目录扫描模式处理（会判重），可重复指定")
    parser.add_argument("--all", action="store_true", help="扫描模式下忽略判重，全部重算")
    parser.add_argument("--dry-run", action="store_true", help="只解析打印，不写文件")
    parser.add_argument("--fill-template", action="store_true", help="额外生成模板回填副本")
    parser.add_argument("--template", dest="template", metavar="XLSX",
                        help="回填用的模板文件，默认用配置区 TEMPLATE_FILE")
    args = parser.parse_args()

    global _LOG_PATH
    run_time = datetime.now()

    if not args.dry_run:
        LOG_DIR.mkdir(parents=True, exist_ok=True)
        _LOG_PATH = LOG_DIR / f"run_{run_time:%Y-%m-%d}.log"

    log("=" * 62)
    log(f"光谱数据汇总  {run_time:%Y-%m-%d %H:%M:%S}")
    log("=" * 62)

    # ---- 1. 确定要处理哪些文件
    explicit: list[Path] = []
    roots: list[Path] = []
    for target in args.targets:
        path = Path(target)
        if path.is_file():
            explicit.append(path)
        elif path.is_dir():
            roots.append(path)
        else:
            log(f"警告：路径不存在，已跳过：{target}")

    scan_mode = not explicit and not roots
    if scan_mode:
        roots = [Path(p) for p in (args.scan or [BASE_DIR / SCAN_SUBDIR])]
        log("未指定文件，进入扫描模式：")
        for root in roots:
            log(f"  - {root}")

    candidates = sorted(set(explicit) | set(collect_files(roots)),
                        key=lambda p: p.stat().st_mtime)
    if not candidates:
        log("没有找到任何数据文件。")
        return 0

    log(f"\n待处理文件：{len(candidates)} 个")

    # ---- 2. 判重（只在扫描模式下生效）
    state = load_state()
    processed = state.setdefault("files", {})
    dedupe = scan_mode and not args.all

    pending: list[tuple[Path, dict]] = []
    for path in candidates:
        fingerprint = file_fingerprint(path)
        key = str(path.resolve())
        if dedupe and processed.get(key, {}).get("sha1") == fingerprint["sha1"]:
            log(f"  [跳过] {path.name} 已汇总过（需要重算加 --all）")
            continue
        pending.append((path, fingerprint))

    if not pending:
        log("没有需要处理的新文件。")
        return 0

    # ---- 3. 读手写参数
    params_table: dict = {}
    if args.params_file:
        try:
            params_table = load_params(Path(args.params_file))
            log(f"\n手写参数表：{Path(args.params_file).name}（{len(params_table)} 条）")
        except (OSError, ValueError) as exc:
            log(f"警告：手写参数表读取失败，本次忽略：{exc}")

    # ---- 4. 解析 + 统计
    rows: list[dict] = []
    failures: list[str] = []
    for path, fingerprint in pending:
        try:
            parsed = parse_data_file(path)
        except (ValueError, OSError) as exc:
            failures.append(f"{path.name}：{exc}")
            log(f"  [失败] {exc}")
            continue

        params = lookup_params(params_table, path)
        identity = resolve_identity(path, params)
        row = {
            "样品编号": identity["display"],
            "用纤长度/m": identity["length_m"],
            "文件名": path.name,
            "测试时间": parsed["measured_at"],
            "模板列头": identity["header"],
            "数据点": parsed["points"],
            "SHA1": fingerprint["sha1"],
            "文件路径": str(path),
        }
        row.update(summarize(parsed))
        for source, target in PARAM_COLUMNS.items():
            row[target] = params.get(source) or None
        rows.append(row)
        processed[str(path.resolve())] = fingerprint

        log(f"  [成功] {identity['display']}  {path.name}")
        log(f"         平均增益 {row['平均增益dB']} dB │ 平坦度 {row['增益平坦度dB']} dB"
            f" │ 1626nm {row['1626nm增益dB']} dB │ 平均NF {row['平均NFdB']} dB")

    if not rows:
        log("\n本次没有可汇总的有效数据。")
        return 2

    # ---- 5. 出表
    if args.dry_run:
        log("\n[dry-run] 不写文件。统计预览：")
        for row in rows:
            log(f"  {row['样品编号']:<14} {row['文件名']:<14} "
                f"平均 {row['平均增益dB']:>7} │ 平坦 {row['增益平坦度dB']:>6} │ "
                f"min {row['最小增益dB']:>7}")
        return 0

    workbook = write_workbook(rows, run_time, args.params_file)
    log(f"\n已生成汇总表：{workbook}")
    csv_path = append_cumulative_csv(rows, run_time)
    log(f"累计统计已更新：{csv_path}")

    if args.fill_template:
        fill_template(rows, run_time, Path(args.template) if args.template else BASE_DIR / TEMPLATE_FILE)

    if scan_mode:
        save_state(state)

    log(f"\n本次成功 {len(rows)} 个，失败 {len(failures)} 个。")
    for message in failures:
        log(f"  失败 → {message}")
    log("=" * 62)
    return 0 if not failures else 1


if __name__ == "__main__":
    try:
        sys.exit(main())
    except Exception:
        traceback.print_exc()
        sys.exit(2)
