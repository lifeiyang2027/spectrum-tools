# -*- coding: utf-8 -*-
"""
生成模拟的光谱仪导出文件，用于验证 gain_summary.py 的解析与统计逻辑。

输出格式模仿仪器导出：文件头若干参数行 + 一行以 "NO.," 开头的数据表头 + 数据块。

重要：这里的数值全部由公式和随机扰动生成，不具备任何工程意义，
      不来自任何真实设备测量，不可用于调试、校准或工艺判断。
"""

from __future__ import annotations

import math
import random
from pathlib import Path

# 输出目录
OUT_DIR = Path(__file__).resolve().parent.parent / "data" / "sample"

# 扫描参数（行业通用的 C+L 波段范围，取自公开文献常见设置）
WL_START = 1524.0
WL_END = 1626.0
POINTS = 120

# 每个样品不同的形态参数，让两份数据看起来不一样
SAMPLES = [
    {"name": "demo_001", "peak_nm": 1560.0, "peak_gain": 28.0, "width": 60.0, "nf_base": 5.0},
    {"name": "demo_002", "peak_nm": 1545.0, "peak_gain": 24.0, "width": 45.0, "nf_base": 5.6},
]


def gain_curve(wl: float, peak_nm: float, peak_gain: float, width: float) -> float:
    """用一个钟形函数模拟增益谱：峰值处最高，两侧衰减。"""
    return peak_gain * math.exp(-((wl - peak_nm) ** 2) / (2 * width ** 2))


def noise_figure(wl: float, nf_base: float) -> float:
    """噪声系数随波长缓慢上升，短波侧略低。"""
    return nf_base + (wl - WL_START) / (WL_END - WL_START) * 1.8


def build_file(sample: dict, seed: int) -> str:
    rng = random.Random(seed)

    wavelengths = [
        WL_START + i * (WL_END - WL_START) / (POINTS - 1) for i in range(POINTS)
    ]

    lines = [
        "Simulated Spectrum Export",
        "Model, DEMO-SPECTRUM",
        "Date, 2026 Nov 15 10:00",
        "Sweep, 1524.000-1626.000nm",
        "Resolution, 0.5nm",
        "Points, 120",
        "Input Level, -5.00dBm",
        "Operator, demo",
        "Note, generated for code validation only",
        "-----",
        "",
        "",
        "NO.,WAVELENGTH[nm],INPUT LVL,OUTPUT LVL,ASE LVL,RESOLN,GAIN[dB],NF[dB]",
    ]

    input_level = -5.0
    for index, wl in enumerate(wavelengths, start=1):
        gain = gain_curve(wl, sample["peak_nm"], sample["peak_gain"], sample["width"])
        gain += rng.gauss(0, 0.12)                       # 轻微抖动
        nf = noise_figure(wl, sample["nf_base"]) + rng.gauss(0, 0.05)
        output = input_level + gain
        ase = output - 45.0 - rng.uniform(0, 2)
        lines.append(
            f"{index},{wl:.3f},{input_level:.2f},{output:.3f},{ase:.3f},0.500,{gain:.3f},{nf:.3f}"
        )

    return "\n".join(lines) + "\n"


def main() -> None:
    OUT_DIR.mkdir(parents=True, exist_ok=True)
    for i, sample in enumerate(SAMPLES):
        text = build_file(sample, seed=1000 + i)
        target = OUT_DIR / f"{sample['name']}.csv"
        target.write_text(text, encoding="utf-8")
        print(f"已生成 {target}（{POINTS} 点）")


if __name__ == "__main__":
    main()
