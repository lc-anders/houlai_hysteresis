# -*- coding: utf-8 -*-
"""
houlai_hysteresis 滞回曲线处理与抗震评价程序（Python/Streamlit 版）
作者：ChatGPT 生成，可自由修改
功能：
1. 导入 xlsx/xls/csv/txt 文件，选择位移列与荷载/反力列；
2. 停顿点修正、隔行取数、力值平滑；
3. 滞回环自动分解；
4. 输出每圈滞回环信息、累计耗能、割线刚度、等效黏滞阻尼系数、残余变形；
5. 输出三类骨架曲线：位移峰值法、荷载峰值法、滞回包络线法；
6. 计算延性系数：割线刚度退化法、能量等效法、Park 近似法；
7. 输出 Excel、PNG 图、GIF 动图，并打包为 ZIP 下载。

运行：
    pip install -r requirements.txt
    streamlit run app.py
"""

from __future__ import annotations

import io
import math
import os
import re
import tempfile
import zipfile
from dataclasses import dataclass, asdict
from pathlib import Path
from typing import Dict, List, Tuple, Optional, Any

import numpy as np
import pandas as pd
import matplotlib.pyplot as plt

try:
    import streamlit as st
except Exception:  # 允许命令行环境仅进行语法检查
    st = None

try:
    import imageio.v2 as imageio
except Exception:
    imageio = None


# -----------------------------
# 基础数据结构
# -----------------------------

@dataclass
class ProcessConfig:
    # 基础数据
    disp_col: str
    force_col: str
    sheet_name: Optional[str] = None
    has_header: bool = True
    disp_unit: str = "mm"
    force_unit: str = "kN"
    zero_origin: bool = False

    # 预处理
    remove_pause: bool = True
    pause_dx: float = 0.02
    pause_dy: float = 0.02
    downsample_step: int = 1
    smooth_method: str = "不平滑"
    smooth_window: int = 11
    savgol_polyorder: int = 3
    control_points: int = 80

    # 滞回环分解
    loop_method: str = "加载级位移峰值法（推荐）"
    zero_eps: float = 1e-6
    min_points_per_loop: int = 15
    level_tolerance: float = 2.0
    loading_levels: str = "5,10,15,20,25,30"
    peak_prominence: float = 1.0
    peak_distance: int = 20
    min_peak_abs: float = 2.0

    # 指标计算可选项
    skeleton_first_cycle_only: bool = True
    loop_area_method: str = "路径积分trapz取绝对值"
    stiffness_peak_reference: str = "位移峰值点"
    damping_formula: str = "JGJ常用三角形能量法"
    ductility_skeleton_source: str = "位移峰值骨架"
    init_stiffness_method: str = "低荷载线性拟合"
    init_stiffness_force_ratio: float = 0.4
    ultimate_strength_ratio: float = 0.85
    ultimate_displacement_policy: str = "未下降则取最大骨架位移"

    # 输出与图表
    output_language: str = "中文"
    excel_column_language: str = "保留英文变量名"
    figure_format: str = "png"
    figure_dpi: int = 160
    font_family: str = "Microsoft YaHei"
    line_width: float = 1.0
    marker_size: float = 3.0
    show_axis_zero_line: bool = True
    max_plot_points: int = 15000
    plot_raw_processed: bool = True
    plot_separated_loops: bool = True
    plot_skeleton: bool = True
    plot_cumulative_energy: bool = True
    plot_secant_stiffness: bool = True
    plot_equivalent_damping: bool = True
    export_raw_data: bool = True
    export_processed_data: bool = True
    export_detected_peaks: bool = True
    export_loop_info: bool = True
    export_separated_loops: bool = True
    export_skeleton_data: bool = True
    export_ductility: bool = True
    export_level_summary: bool = True
    make_gif: bool = True
    gif_step: int = 30
    gif_interval_ms: int = 60


# -----------------------------
# 语言、字体和输出辅助
# -----------------------------

TEXT = {
    "中文": {
        "app_title": "houlai_hysteresis 滞回曲线处理与抗震评价程序 v1.0",
        "app_caption": "适用于低周往复加载的位移–荷载/反力滞回曲线。支持中英文图表、可选计算参数和批量导出。",
        "disp": "位移",
        "force": "荷载/反力",
        "raw": "原始数据",
        "processed": "处理后数据",
        "hys_title": "滞回曲线：原始数据与处理后数据对比",
        "loops_title": "滞回环分解图",
        "skeleton_title": "骨架曲线",
        "disp_skeleton": "位移峰值骨架曲线",
        "force_skeleton": "荷载峰值骨架曲线",
        "envelope_skeleton": "滞回包络骨架曲线",
        "loop_id": "滞回环编号",
        "cum_energy": "累计耗能",
        "cum_energy_title": "累计耗能曲线",
        "stiffness": "割线刚度",
        "stiffness_title": "割线刚度退化曲线",
        "heq": "等效黏滞阻尼系数",
        "heq_title": "等效黏滞阻尼系数",
        "animation_title": "滞回曲线生成动画",
    },
    "English": {
        "app_title": "houlai_hysteresis Hysteresis Analysis and Seismic Evaluation Program v1.0",
        "app_caption": "For cyclic displacement–force hysteresis data. Supports bilingual figures, optional calculation settings, and batch export.",
        "disp": "Displacement",
        "force": "Force",
        "raw": "Raw",
        "processed": "Processed",
        "hys_title": "Hysteresis Curve: Raw vs. Processed",
        "loops_title": "Separated Hysteresis Loops",
        "skeleton_title": "Skeleton Curves",
        "disp_skeleton": "Displacement-peak skeleton",
        "force_skeleton": "Force-peak skeleton",
        "envelope_skeleton": "Envelope skeleton",
        "loop_id": "Loop ID",
        "cum_energy": "Cumulative Energy",
        "cum_energy_title": "Cumulative Energy Dissipation",
        "stiffness": "Secant Stiffness",
        "stiffness_title": "Secant Stiffness Degradation",
        "heq": "Equivalent Viscous Damping",
        "heq_title": "Equivalent Viscous Damping Coefficient",
        "animation_title": "Hysteresis Animation",
    },
}

COLUMN_CN_MAP = {
    "LoopID": "滞回环编号",
    "StartPointID": "起始点编号",
    "EndPointID": "结束点编号",
    "PointCount": "点数",
    "AmplitudeLevel": "加载位移级别",
    "Dmax_pos": "正向最大位移",
    "F_at_Dmax_pos": "正向最大位移对应力",
    "Dmin_neg": "负向最大位移",
    "F_at_Dmin_neg": "负向最大位移对应力",
    "Fmax_pos": "正向最大力",
    "D_at_Fmax_pos": "正向最大力对应位移",
    "Fmin_neg": "负向最大力",
    "D_at_Fmin_neg": "负向最大力对应位移",
    "Area_signed": "带符号面积",
    "LoopArea_Eh": "单圈耗能面积",
    "CumulativeEnergy": "累计耗能",
    "SecantStiffness_K": "割线刚度",
    "SecantStiffness_pos": "正向割线刚度",
    "SecantStiffness_neg": "负向割线刚度",
    "ElasticEnergy_triangle": "三角形弹性能",
    "EnergyDissipationCoeff": "能量耗散系数",
    "EquivalentViscousDamping_heq": "等效黏滞阻尼系数",
    "ResidualDisp_pos": "正向残余变形",
    "ResidualDisp_neg": "负向残余变形",
    "ResidualDisp_absmax": "最大绝对残余变形",
    "StrengthDegradation_pos_to_prev": "同级正向强度退化系数",
    "StrengthDegradation_neg_to_prev": "同级负向强度退化系数",
    "StiffnessDegradation_to_prev": "同级刚度退化系数",
    "D": "位移",
    "F": "力",
    "F_smooth": "处理后力",
    "RawIndex": "原始行号",
    "PointID": "点编号",
    "Branch": "分支",
    "Method": "方法",
    "Direction": "方向",
    "Ductility_mu": "延性系数",
}


def T(cfg: ProcessConfig, key: str) -> str:
    lang = "English" if getattr(cfg, "output_language", "中文") == "English" else "中文"
    return TEXT.get(lang, TEXT["中文"]).get(key, key)


def setup_matplotlib_style(cfg: ProcessConfig):
    plt.rcParams["axes.unicode_minus"] = False
    if getattr(cfg, "output_language", "中文") == "中文":
        fonts = [getattr(cfg, "font_family", "Microsoft YaHei"), "Microsoft YaHei", "SimHei", "Noto Sans CJK SC", "Arial Unicode MS", "DejaVu Sans"]
        plt.rcParams["font.sans-serif"] = fonts
    else:
        plt.rcParams["font.sans-serif"] = ["DejaVu Sans", "Arial", "Liberation Sans"]


def fig_path(folder: Path, stem: str, cfg: ProcessConfig) -> Path:
    ext = str(getattr(cfg, "figure_format", "png")).lower().strip().lstrip(".") or "png"
    if ext == "jpeg":
        ext = "jpg"
    return folder / f"{stem}.{ext}"


def maybe_translate_columns(df: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    if getattr(cfg, "excel_column_language", "保留英文变量名") == "中文列名":
        return df.rename(columns=COLUMN_CN_MAP)
    return df


# -----------------------------
# 文件读取
# -----------------------------

def sanitize_filename(name: str) -> str:
    base = Path(name).stem
    base = re.sub(r"[^\w\u4e00-\u9fa5\-]+", "_", base)
    return base[:80] if base else "hysteresis"


def list_excel_sheets(file_obj) -> List[str]:
    pos = file_obj.tell() if hasattr(file_obj, "tell") else 0
    try:
        xls = pd.ExcelFile(file_obj)
        return list(xls.sheet_names)
    finally:
        try:
            file_obj.seek(pos)
        except Exception:
            pass


def read_input_table(file_obj, cfg: ProcessConfig) -> pd.DataFrame:
    """读取 Excel/CSV/TXT 文件。"""
    name = getattr(file_obj, "name", "")
    suffix = Path(name).suffix.lower()
    header = 0 if cfg.has_header else None

    if suffix in [".xlsx", ".xls", ".xlsm"]:
        df = pd.read_excel(file_obj, sheet_name=cfg.sheet_name or 0, header=header)
    elif suffix in [".csv"]:
        df = pd.read_csv(file_obj, header=header)
    elif suffix in [".txt", ".dat"]:
        # 自动识别逗号、制表符、空格分隔
        raw = file_obj.read()
        if isinstance(raw, bytes):
            raw = raw.decode("utf-8", errors="ignore")
        file_like = io.StringIO(raw)
        try:
            df = pd.read_csv(file_like, header=header, sep=None, engine="python")
        except Exception:
            file_like.seek(0)
            df = pd.read_csv(file_like, header=header, delim_whitespace=True)
    else:
        raise ValueError(f"暂不支持文件格式：{suffix}")

    if not cfg.has_header:
        df.columns = [f"列{i+1}" for i in range(df.shape[1])]
    return df


def extract_xy(df: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    if cfg.disp_col not in df.columns or cfg.force_col not in df.columns:
        # 多文件批量时列名不一致，自动退回前两列
        dcol, fcol = df.columns[0], df.columns[1]
    else:
        dcol, fcol = cfg.disp_col, cfg.force_col

    out = pd.DataFrame({
        "D": pd.to_numeric(df[dcol], errors="coerce"),
        "F": pd.to_numeric(df[fcol], errors="coerce"),
    })
    out = out.dropna().reset_index(drop=True)
    out.insert(0, "RawIndex", np.arange(len(out)))

    if cfg.zero_origin and len(out) > 0:
        out["D"] = out["D"] - out["D"].iloc[0]
        out["F"] = out["F"] - out["F"].iloc[0]

    return out


# -----------------------------
# 预处理：停顿点、隔行取数、平滑
# -----------------------------

def local_extrema_indices(y: np.ndarray) -> np.ndarray:
    """简单寻找局部极值点，用于避免停顿点删除时误删峰值。"""
    y = np.asarray(y, dtype=float)
    if len(y) < 3:
        return np.array([], dtype=int)
    dy = np.diff(y)
    sign = np.sign(dy)
    # 填补 0，避免平台导致极值识别失败
    for i in range(1, len(sign)):
        if sign[i] == 0:
            sign[i] = sign[i - 1]
    for i in range(len(sign) - 2, -1, -1):
        if sign[i] == 0:
            sign[i] = sign[i + 1]
    turn = np.where(np.diff(sign) != 0)[0] + 1
    return turn.astype(int)


def remove_pause_points(df: pd.DataFrame, dx_thr: float, dy_thr: float) -> pd.DataFrame:
    """删除相邻变化很小的停顿点，同时强制保留位移/力的局部极值与首尾点。"""
    if len(df) <= 3:
        return df.copy()

    d = df["D"].to_numpy()
    f = df["F"].to_numpy()
    dx = np.r_[np.inf, np.abs(np.diff(d))]
    dy = np.r_[np.inf, np.abs(np.diff(f))]

    keep = (dx > dx_thr) | (dy > dy_thr)
    keep[0] = True
    keep[-1] = True

    extrema = np.unique(np.r_[local_extrema_indices(d), local_extrema_indices(f)])
    extrema = extrema[(extrema >= 0) & (extrema < len(df))]
    keep[extrema] = True

    return df.loc[keep].reset_index(drop=True)


def downsample(df: pd.DataFrame, step: int) -> pd.DataFrame:
    step = max(1, int(step))
    if step <= 1 or len(df) <= 2:
        return df.copy()
    idx = np.arange(0, len(df), step)
    if idx[-1] != len(df) - 1:
        idx = np.r_[idx, len(df) - 1]
    return df.iloc[idx].reset_index(drop=True)


def ensure_odd(n: int, minimum: int = 3) -> int:
    n = max(minimum, int(n))
    return n if n % 2 == 1 else n + 1


def smooth_force(df: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    out = df.copy()
    method = cfg.smooth_method
    if method == "不平滑" or len(out) < 5:
        out["F_smooth"] = out["F"]
        return out

    f = out["F"].to_numpy(dtype=float)
    n = len(f)
    window = min(ensure_odd(cfg.smooth_window), n if n % 2 == 1 else n - 1)
    if window < 3:
        out["F_smooth"] = out["F"]
        return out

    if method == "移动平均":
        out["F_smooth"] = pd.Series(f).rolling(window=window, center=True, min_periods=1).mean().to_numpy()
    elif method == "移动中值":
        out["F_smooth"] = pd.Series(f).rolling(window=window, center=True, min_periods=1).median().to_numpy()
    elif method == "Savitzky-Golay":
        try:
            from scipy.signal import savgol_filter
            poly = min(int(cfg.savgol_polyorder), window - 2)
            out["F_smooth"] = savgol_filter(f, window_length=window, polyorder=max(1, poly), mode="interp")
        except Exception:
            # 未安装 scipy 时退回移动平均
            out["F_smooth"] = pd.Series(f).rolling(window=window, center=True, min_periods=1).mean().to_numpy()
    elif method == "控制点插值":
        m = min(max(4, int(cfg.control_points)), n)
        ctrl_idx = np.unique(np.linspace(0, n - 1, m).astype(int))
        out["F_smooth"] = np.interp(np.arange(n), ctrl_idx, f[ctrl_idx])
    else:
        out["F_smooth"] = out["F"]

    return out


def preprocess(raw: pd.DataFrame, cfg: ProcessConfig) -> Tuple[pd.DataFrame, Dict[str, Any]]:
    df = raw.copy()
    before = len(df)

    if cfg.remove_pause:
        df = remove_pause_points(df, cfg.pause_dx, cfg.pause_dy)
    after_pause = len(df)

    df = downsample(df, cfg.downsample_step)
    after_down = len(df)

    df = smooth_force(df, cfg)
    df["PointID"] = np.arange(len(df))
    df = df[["PointID", "RawIndex", "D", "F", "F_smooth"]]

    log = {
        "原始点数": before,
        "停顿点修正后点数": after_pause,
        "隔行取数后点数": after_down,
        "最终点数": len(df),
        "删除点数": before - len(df),
        "力值列": "F_smooth",
    }
    return df, log


# -----------------------------
# 滞回环分解
# -----------------------------

def filled_sign(x: np.ndarray, eps: float = 1e-9) -> np.ndarray:
    x = np.asarray(x, dtype=float)
    s = np.zeros(len(x), dtype=int)
    s[x > eps] = 1
    s[x < -eps] = -1

    # forward fill
    last = 0
    for i in range(len(s)):
        if s[i] == 0:
            s[i] = last
        else:
            last = s[i]
    # backward fill leading zeros
    next_val = 0
    for i in range(len(s) - 1, -1, -1):
        if s[i] == 0:
            s[i] = next_val
        else:
            next_val = s[i]
    return s


def segment_by_zero_crossing(df: pd.DataFrame, cfg: ProcessConfig) -> List[Tuple[int, int]]:
    """零位移穿越法：每两个半周合并为一圈。"""
    d = df["D"].to_numpy()
    if len(d) < cfg.min_points_per_loop:
        return []

    s = filled_sign(d, cfg.zero_eps)
    cross = np.where((s[1:] != s[:-1]) & (s[1:] != 0) & (s[:-1] != 0))[0] + 1

    # 将首尾加入边界
    bounds = np.unique(np.r_[0, cross, len(df) - 1]).astype(int)
    bounds = bounds[np.r_[True, np.diff(bounds) > 2]]

    loops: List[Tuple[int, int]] = []
    for j in range(0, len(bounds) - 2, 2):
        a, b = int(bounds[j]), int(bounds[j + 2])
        if b - a + 1 >= cfg.min_points_per_loop:
            seg = df.iloc[a:b + 1]
            if seg["D"].max() > cfg.zero_eps and seg["D"].min() < -cfg.zero_eps:
                loops.append((a, b))
    return loops


def segment_by_turning_points(df: pd.DataFrame, cfg: ProcessConfig) -> List[Tuple[int, int]]:
    """位移转折点法：以正负峰值交替为依据，适用于数据未明显回零的情形。"""
    d = df["D"].to_numpy()
    turns = local_extrema_indices(d)
    if len(turns) < 4:
        return segment_by_zero_crossing(df, cfg)

    bounds = np.unique(np.r_[0, turns, len(df) - 1]).astype(int)
    loops: List[Tuple[int, int]] = []
    # 一个完整环通常包含两个相邻正/负峰值和回程，近似按 4 个转折点划分
    for j in range(0, len(bounds) - 4, 4):
        a, b = int(bounds[j]), int(bounds[j + 4])
        if b - a + 1 >= cfg.min_points_per_loop:
            seg = df.iloc[a:b + 1]
            if seg["D"].max() > cfg.zero_eps and seg["D"].min() < -cfg.zero_eps:
                loops.append((a, b))
    if not loops:
        return segment_by_zero_crossing(df, cfg)
    return loops



def parse_loading_levels(text: str) -> List[float]:
    """解析用户输入的目标加载位移级别，例如：5,10,15,20,25,30。"""
    if not text:
        return []
    vals: List[float] = []
    for part in re.split(r"[,，;；\s]+", str(text).strip()):
        if not part:
            continue
        try:
            vals.append(abs(float(part)))
        except Exception:
            pass
    vals = sorted(set(v for v in vals if np.isfinite(v) and v > 0))
    return vals


def classify_amplitude_level(amp: float, cfg: ProcessConfig) -> float:
    """将实际峰值位移归并到目标加载级别，避免 20.6 被误分到 21 级。"""
    amp = abs(float(amp))
    levels = parse_loading_levels(getattr(cfg, "loading_levels", ""))
    if levels:
        arr = np.asarray(levels, dtype=float)
        nearest = float(arr[np.argmin(np.abs(arr - amp))])
        tol = max(float(cfg.level_tolerance), 0.10 * nearest)
        if abs(nearest - amp) <= tol:
            return nearest
    if cfg.level_tolerance > 0:
        return round(amp / cfg.level_tolerance) * cfg.level_tolerance
    return amp


def significant_peak_events(df: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    """识别显著的正、负位移峰值，用于加载级位移峰值法分圈。"""
    d = df["D"].to_numpy(dtype=float)
    f = df["F_smooth"].to_numpy(dtype=float)
    if len(d) < 5:
        return pd.DataFrame(columns=["idx", "Type", "Sign", "D", "F", "AbsD", "AmplitudeLevel"])

    max_abs = float(np.nanmax(np.abs(d))) if len(d) else 0.0
    min_abs = float(cfg.min_peak_abs) if float(cfg.min_peak_abs) > 0 else max(0.10 * max_abs, 1e-9)
    prominence = float(cfg.peak_prominence) if float(cfg.peak_prominence) > 0 else max(0.03 * max_abs, 1e-9)
    distance = max(3, int(cfg.peak_distance))

    try:
        from scipy.signal import find_peaks
        pos_idx, _ = find_peaks(d, prominence=prominence, distance=distance)
        neg_idx, _ = find_peaks(-d, prominence=prominence, distance=distance)
    except Exception:
        turns = local_extrema_indices(d)
        pos_idx = turns[d[turns] > 0]
        neg_idx = turns[d[turns] < 0]

    rows = []
    for i in pos_idx:
        i = int(i)
        if abs(d[i]) >= min_abs:
            rows.append({
                "idx": i,
                "Type": "positive",
                "Sign": 1,
                "D": float(d[i]),
                "F": float(f[i]),
                "AbsD": float(abs(d[i])),
                "AmplitudeLevel": classify_amplitude_level(abs(d[i]), cfg),
            })
    for i in neg_idx:
        i = int(i)
        if abs(d[i]) >= min_abs:
            rows.append({
                "idx": i,
                "Type": "negative",
                "Sign": -1,
                "D": float(d[i]),
                "F": float(f[i]),
                "AbsD": float(abs(d[i])),
                "AmplitudeLevel": classify_amplitude_level(abs(d[i]), cfg),
            })

    if not rows:
        return pd.DataFrame(columns=["idx", "Type", "Sign", "D", "F", "AbsD", "AmplitudeLevel"])
    ev = pd.DataFrame(rows).sort_values("idx").reset_index(drop=True)

    # 若两个相邻峰值同号，保留绝对位移更大的峰值，减少噪声造成的多峰干扰。
    cleaned = []
    for _, row in ev.iterrows():
        if cleaned and int(cleaned[-1]["Sign"]) == int(row["Sign"]):
            if float(row["AbsD"]) >= float(cleaned[-1]["AbsD"]):
                cleaned[-1] = row.to_dict()
        else:
            cleaned.append(row.to_dict())
    return pd.DataFrame(cleaned).reset_index(drop=True)


def zero_boundary_candidates(d: np.ndarray, cfg: ProcessConfig) -> Tuple[np.ndarray, float]:
    """寻找接近零位移或发生正负号穿越的位置，作为滞回环边界候选。"""
    d = np.asarray(d, dtype=float)
    if len(d) == 0:
        return np.array([], dtype=int), 0.0
    max_abs = float(np.nanmax(np.abs(d)))
    zero_zone = max(float(cfg.zero_eps), 0.01 * max_abs, 0.05)

    near_zero = np.where(np.abs(d) <= zero_zone)[0]
    crossing = np.where((d[:-1] * d[1:] <= 0) | (np.abs(d[:-1]) <= zero_zone) | (np.abs(d[1:]) <= zero_zone))[0] + 1
    candidates = np.unique(np.r_[0, near_zero, crossing, len(d) - 1]).astype(int)
    return np.sort(candidates), zero_zone


def segment_by_peak_pairing(df: pd.DataFrame, cfg: ProcessConfig) -> List[Tuple[int, int]]:
    """
    加载级位移峰值法：先识别显著正/负位移峰值，再按相邻异号峰值成对分圈。
    适合 5、10、15、20、25、30 mm 这类分级加载且每级加载一圈或多圈的数据。
    """
    events = significant_peak_events(df, cfg)
    if len(events) < 2:
        return segment_by_zero_crossing(df, cfg)

    d = df["D"].to_numpy(dtype=float)
    zero_candidates, _ = zero_boundary_candidates(d, cfg)

    loops: List[Tuple[int, int]] = []
    last_end = 0
    i = 0
    while i < len(events) - 1:
        a = events.iloc[i]
        b = events.iloc[i + 1]

        if int(a["Sign"]) == int(b["Sign"]):
            i += 1
            continue

        ia, ib = int(a["idx"]), int(b["idx"])
        if ia > ib:
            ia, ib = ib, ia

        before = zero_candidates[zero_candidates < ia]
        after = zero_candidates[zero_candidates > ib]
        start = int(before[-1]) if len(before) else last_end
        end = int(after[0]) if len(after) else len(df) - 1

        # 避免相邻滞回环边界过度重叠。
        if start < last_end:
            start = last_end

        if end - start + 1 >= cfg.min_points_per_loop:
            seg = df.iloc[start:end + 1]
            has_pos = seg["D"].max() > cfg.zero_eps
            has_neg = seg["D"].min() < -cfg.zero_eps
            if has_pos and has_neg:
                loops.append((start, end))
                last_end = end

        i += 2

    if not loops:
        return segment_by_zero_crossing(df, cfg)
    return loops


def segment_loops(df: pd.DataFrame, cfg: ProcessConfig) -> List[Tuple[int, int]]:
    if cfg.loop_method == "加载级位移峰值法（推荐）":
        return segment_by_peak_pairing(df, cfg)
    if cfg.loop_method == "位移转折点法":
        return segment_by_turning_points(df, cfg)
    return segment_by_zero_crossing(df, cfg)


# -----------------------------
# 指标计算
# -----------------------------

def signed_trapz_area(d: np.ndarray, f: np.ndarray) -> float:
    if len(d) < 2:
        return 0.0
    return float(np.trapz(f, d))


def zero_force_residuals(d: np.ndarray, f: np.ndarray) -> Tuple[float, float, float]:
    """计算力为 0 时的残余位移。返回：正残余、负残余、最大绝对残余。"""
    residuals = []
    for i in range(1, len(f)):
        f1, f2 = f[i - 1], f[i]
        if not np.isfinite(f1) or not np.isfinite(f2):
            continue
        if f1 == 0:
            residuals.append(d[i - 1])
        elif f1 * f2 < 0:
            # 线性插值求 F=0 处的 D
            t = -f1 / (f2 - f1)
            residuals.append(d[i - 1] + t * (d[i] - d[i - 1]))
    if not residuals:
        return np.nan, np.nan, np.nan
    residuals = np.asarray(residuals)
    pos = residuals[residuals >= 0]
    neg = residuals[residuals < 0]
    rpos = float(pos.max()) if len(pos) else np.nan
    rneg = float(neg.min()) if len(neg) else np.nan
    rabs = float(residuals[np.argmax(np.abs(residuals))])
    return rpos, rneg, rabs


def compute_loop_area(d: np.ndarray, f: np.ndarray, cfg: ProcessConfig) -> Tuple[float, float]:
    """返回：带符号面积、用于耗能统计的正面积。"""
    method = getattr(cfg, "loop_area_method", "闭合环面积|trapz取绝对值")
    signed = signed_trapz_area(d, f)
    if "分段绝对积分" in method:
        # 对每个相邻区段的梯形面积取绝对值。适合曲线非常震荡时检查，但通常会放大耗能。
        area = float(np.sum(np.abs(0.5 * (f[1:] + f[:-1]) * np.diff(d)))) if len(d) >= 2 else 0.0
    elif "多边形" in method or "闭合" in method:
        # Shoelace 闭合多边形面积。对闭合滞回环更稳健。
        if len(d) >= 3:
            x = np.r_[d, d[0]]
            y = np.r_[f, f[0]]
            area = float(0.5 * np.abs(np.sum(x[:-1] * y[1:] - x[1:] * y[:-1])))
        else:
            area = abs(signed)
    else:
        # 默认：对 F-D 路径积分后取绝对值。
        area = abs(signed)
    return signed, area


def compute_loop_info(df: pd.DataFrame, loops: List[Tuple[int, int]], cfg: ProcessConfig) -> Tuple[pd.DataFrame, Dict[int, pd.DataFrame]]:
    rows = []
    loop_data: Dict[int, pd.DataFrame] = {}
    cumulative = 0.0

    for loop_id, (a, b) in enumerate(loops, start=1):
        seg = df.iloc[a:b + 1].copy().reset_index(drop=True)
        d = seg["D"].to_numpy()
        f = seg["F_smooth"].to_numpy()

        if len(seg) < 2:
            continue

        idx_dmax = int(np.argmax(d))
        idx_dmin = int(np.argmin(d))
        idx_fmax = int(np.argmax(f))
        idx_fmin = int(np.argmin(f))

        area_signed, area = compute_loop_area(d, f, cfg)
        cumulative += area

        Dpos = float(d[idx_dmax])
        F_at_Dpos = float(f[idx_dmax])
        Dneg = float(d[idx_dmin])
        F_at_Dneg = float(f[idx_dmin])
        Fmax = float(f[idx_fmax])
        D_at_Fmax = float(d[idx_fmax])
        Fmin = float(f[idx_fmin])
        D_at_Fmin = float(d[idx_fmin])

        # 用户可选择刚度和阻尼计算采用“位移峰值点”还是“荷载峰值点”。
        if getattr(cfg, "stiffness_peak_reference", "位移峰值点") == "荷载峰值点":
            Dpos_k, Fpos_k = D_at_Fmax, Fmax
            Dneg_k, Fneg_k = D_at_Fmin, Fmin
        else:
            Dpos_k, Fpos_k = Dpos, F_at_Dpos
            Dneg_k, Fneg_k = Dneg, F_at_Dneg

        denom_k = Dpos_k - Dneg_k
        K_sec = (Fpos_k - Fneg_k) / denom_k if abs(denom_k) > 1e-12 else np.nan
        K_pos = Fpos_k / Dpos_k if abs(Dpos_k) > 1e-12 else np.nan
        K_neg = abs(Fneg_k / Dneg_k) if abs(Dneg_k) > 1e-12 else np.nan

        elastic_energy = 0.5 * abs(Fpos_k * Dpos_k) + 0.5 * abs(Fneg_k * Dneg_k)
        if getattr(cfg, "damping_formula", "JGJ常用三角形能量法") == "不计算":
            energy_dissipation_coeff = np.nan
            heq = np.nan
        else:
            energy_dissipation_coeff = area / elastic_energy if elastic_energy > 1e-12 else np.nan
            heq = energy_dissipation_coeff / (2 * math.pi) if np.isfinite(energy_dissipation_coeff) else np.nan

        # 便于用户复核正、负向分支对阻尼的影响，不作为默认论文指标。
        heq_pos = area / (2 * math.pi * abs(Fpos_k * Dpos_k)) if abs(Fpos_k * Dpos_k) > 1e-12 else np.nan
        heq_neg = area / (2 * math.pi * abs(Fneg_k * Dneg_k)) if abs(Fneg_k * Dneg_k) > 1e-12 else np.nan

        rpos, rneg, rabs = zero_force_residuals(d, f)
        amp = max(abs(Dpos), abs(Dneg))
        level = classify_amplitude_level(amp, cfg)

        rows.append({
            "LoopID": loop_id,
            "StartPointID": int(df.iloc[a]["PointID"]),
            "EndPointID": int(df.iloc[b]["PointID"]),
            "PointCount": int(len(seg)),
            "AmplitudeLevel": level,
            "Dmax_pos": Dpos,
            "F_at_Dmax_pos": F_at_Dpos,
            "Dmin_neg": Dneg,
            "F_at_Dmin_neg": F_at_Dneg,
            "Fmax_pos": Fmax,
            "D_at_Fmax_pos": D_at_Fmax,
            "Fmin_neg": Fmin,
            "D_at_Fmin_neg": D_at_Fmin,
            "Area_signed": area_signed,
            "LoopArea_Eh": area,
            "CumulativeEnergy": cumulative,
            "SecantStiffness_K": K_sec,
            "SecantStiffness_pos": K_pos,
            "SecantStiffness_neg": K_neg,
            "StiffnessReference_Dpos": Dpos_k,
            "StiffnessReference_Fpos": Fpos_k,
            "StiffnessReference_Dneg": Dneg_k,
            "StiffnessReference_Fneg": Fneg_k,
            "ElasticEnergy_triangle": elastic_energy,
            "EnergyDissipationCoeff": energy_dissipation_coeff,
            "EquivalentViscousDamping_heq": heq,
            "EquivalentDamping_pos_check": heq_pos,
            "EquivalentDamping_neg_check": heq_neg,
            "ResidualDisp_pos": rpos,
            "ResidualDisp_neg": rneg,
            "ResidualDisp_absmax": rabs,
            "idx_Dmax_global": int(df.index[a + idx_dmax]),
            "idx_Dmin_global": int(df.index[a + idx_dmin]),
            "idx_Fmax_global": int(df.index[a + idx_fmax]),
            "idx_Fmin_global": int(df.index[a + idx_fmin]),
        })

        seg.insert(0, "LoopID", loop_id)
        loop_data[loop_id] = seg

    info = pd.DataFrame(rows)
    if not info.empty:
        info = add_degradation_columns(info)
    return info, loop_data


def add_degradation_columns(info: pd.DataFrame) -> pd.DataFrame:
    info = info.copy()
    info["StrengthDegradation_pos_to_prev"] = np.nan
    info["StrengthDegradation_neg_to_prev"] = np.nan
    info["StiffnessDegradation_to_prev"] = np.nan

    for level, g in info.groupby("AmplitudeLevel", sort=False):
        idxs = list(g.index)
        for k in range(1, len(idxs)):
            i, p = idxs[k], idxs[k - 1]
            fp, fpp = info.loc[i, "F_at_Dmax_pos"], info.loc[p, "F_at_Dmax_pos"]
            fn, fnp = abs(info.loc[i, "F_at_Dmin_neg"]), abs(info.loc[p, "F_at_Dmin_neg"])
            kk, kkp = info.loc[i, "SecantStiffness_K"], info.loc[p, "SecantStiffness_K"]
            if abs(fpp) > 1e-12:
                info.loc[i, "StrengthDegradation_pos_to_prev"] = fp / fpp
            if abs(fnp) > 1e-12:
                info.loc[i, "StrengthDegradation_neg_to_prev"] = fn / fnp
            if abs(kkp) > 1e-12:
                info.loc[i, "StiffnessDegradation_to_prev"] = kk / kkp
    return info


# -----------------------------
# 骨架曲线
# -----------------------------

def point_from_global_index(df: pd.DataFrame, idx: int, name: str, loop_id: int, level: float) -> Dict[str, Any]:
    r = df.loc[idx]
    return {
        "Branch": name,
        "LoopID": loop_id,
        "AmplitudeLevel": level,
        "PointID": int(r["PointID"]),
        "D": float(r["D"]),
        "F": float(r["F_smooth"]),
    }


def filter_first_cycle(info: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    if not cfg.skeleton_first_cycle_only or info.empty:
        return info
    return info.sort_values("LoopID").drop_duplicates("AmplitudeLevel", keep="first")


def skeleton_by_displacement_peak(df: pd.DataFrame, info: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    rows = [{"Branch": "origin", "LoopID": 0, "AmplitudeLevel": 0, "PointID": -1, "D": 0.0, "F": 0.0}]
    use = filter_first_cycle(info, cfg)
    for _, row in use.iterrows():
        rows.append(point_from_global_index(df, int(row["idx_Dmin_global"]), "negative", int(row["LoopID"]), row["AmplitudeLevel"]))
    for _, row in use.iterrows():
        rows.append(point_from_global_index(df, int(row["idx_Dmax_global"]), "positive", int(row["LoopID"]), row["AmplitudeLevel"]))
    sk = pd.DataFrame(rows)
    sk = sk.sort_values(["D", "LoopID"]).reset_index(drop=True)
    return sk


def skeleton_by_force_peak(df: pd.DataFrame, info: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    rows = [{"Branch": "origin", "LoopID": 0, "AmplitudeLevel": 0, "PointID": -1, "D": 0.0, "F": 0.0}]
    use = filter_first_cycle(info, cfg)
    for _, row in use.iterrows():
        rows.append(point_from_global_index(df, int(row["idx_Fmin_global"]), "negative", int(row["LoopID"]), row["AmplitudeLevel"]))
    for _, row in use.iterrows():
        rows.append(point_from_global_index(df, int(row["idx_Fmax_global"]), "positive", int(row["LoopID"]), row["AmplitudeLevel"]))
    sk = pd.DataFrame(rows)
    sk = sk.sort_values(["D", "LoopID"]).reset_index(drop=True)
    return sk


def skeleton_envelope(df: pd.DataFrame, bins: int = 120) -> pd.DataFrame:
    """简化包络线：正向按位移分箱取最大力，负向按位移分箱取最小力。"""
    d = df["D"].to_numpy()
    f = df["F_smooth"].to_numpy()
    rows = [{"Branch": "origin", "D": 0.0, "F": 0.0}]

    # 正向 envelope
    pos = pd.DataFrame({"D": d[d >= 0], "F": f[d >= 0]})
    if len(pos) > 0:
        pos = pos.sort_values("D")
        if pos["D"].max() > 0:
            pos["bin"] = pd.cut(pos["D"], bins=min(bins, max(5, len(pos)//5)), include_lowest=True)
            gp = pos.groupby("bin", observed=True).agg(D=("D", "max"), F=("F", "max")).dropna()
            gp = gp[gp["D"] >= 0]
            # 保证外包络单调不降低
            gp["F"] = gp["F"].cummax()
            for _, r in gp.iterrows():
                rows.append({"Branch": "positive_envelope", "D": float(r["D"]), "F": float(r["F"])})

    # 负向 envelope
    neg = pd.DataFrame({"D": d[d <= 0], "F": f[d <= 0]})
    if len(neg) > 0:
        neg = neg.sort_values("D")
        if neg["D"].min() < 0:
            neg["bin"] = pd.cut(neg["D"], bins=min(bins, max(5, len(neg)//5)), include_lowest=True)
            gn = neg.groupby("bin", observed=True).agg(D=("D", "min"), F=("F", "min")).dropna()
            gn = gn.sort_values("D")
            # 从 0 向负向位移发展时，包络力绝对值不降低
            gn_rev = gn.sort_values("D", ascending=False).copy()
            gn_rev["F"] = gn_rev["F"].cummin()
            gn = gn_rev.sort_values("D")
            for _, r in gn.iterrows():
                rows.append({"Branch": "negative_envelope", "D": float(r["D"]), "F": float(r["F"])})

    sk = pd.DataFrame(rows)
    sk = sk.drop_duplicates(subset=["Branch", "D", "F"]).sort_values("D").reset_index(drop=True)
    return sk


# -----------------------------
# 延性系数
# -----------------------------

def interpolate_x_at_y(x: np.ndarray, y: np.ndarray, target_y: float) -> Optional[float]:
    """在序列首次达到 target_y 处线性插值。"""
    for i in range(1, len(y)):
        if (y[i - 1] - target_y) * (y[i] - target_y) <= 0 and y[i] != y[i - 1]:
            t = (target_y - y[i - 1]) / (y[i] - y[i - 1])
            return float(x[i - 1] + t * (x[i] - x[i - 1]))
    return None


def interpolate_y_at_x(x: np.ndarray, y: np.ndarray, target_x: float) -> float:
    order = np.argsort(x)
    xs, ys = x[order], y[order]
    return float(np.interp(target_x, xs, ys))


def area_until_x(x: np.ndarray, y: np.ndarray, target_x: float) -> float:
    order = np.argsort(x)
    xs, ys = x[order], y[order]
    mask = xs <= target_x
    xs2 = xs[mask]
    ys2 = ys[mask]
    if len(xs2) == 0 or xs2[-1] < target_x:
        ytarget = interpolate_y_at_x(xs, ys, target_x)
        xs2 = np.r_[xs2, target_x]
        ys2 = np.r_[ys2, ytarget]
    if len(xs2) < 2:
        return 0.0
    return float(np.trapz(ys2, xs2))


def prepare_branch(skeleton: pd.DataFrame, direction: str) -> pd.DataFrame:
    sk = skeleton.copy()
    if direction == "positive":
        br = sk[(sk["D"] >= 0) & (sk["F"] >= 0)].copy()
        br["x"] = br["D"]
        br["y"] = br["F"]
    else:
        br = sk[(sk["D"] <= 0) & (sk["F"] <= 0)].copy()
        br["x"] = -br["D"]
        br["y"] = -br["F"]

    br = br[np.isfinite(br["x"]) & np.isfinite(br["y"])]
    br = br[br["x"] >= 0].sort_values("x")
    if br.empty:
        return br

    # 同一位移只保留最大承载力
    br = br.groupby("x", as_index=False).agg(y=("y", "max"))
    br = br.sort_values("x").reset_index(drop=True)
    return br


def estimate_initial_stiffness(x: np.ndarray, y: np.ndarray, ratio: float) -> float:
    if len(x) < 2:
        return np.nan
    fmax = np.nanmax(y)
    use = (y > 0) & (y <= ratio * fmax) & (x > 0)
    if use.sum() >= 2:
        coef = np.polyfit(x[use], y[use], 1)
        return float(coef[0])
    # fallback：第一个有效点的割线刚度
    valid = np.where((x > 1e-12) & (y > 0))[0]
    if len(valid) > 0:
        i = valid[0]
        return float(y[i] / x[i])
    return np.nan


def ultimate_displacement(x: np.ndarray, y: np.ndarray, ratio: float) -> Tuple[float, str]:
    imax = int(np.argmax(y))
    fmax = y[imax]
    target = ratio * fmax
    for i in range(imax + 1, len(y)):
        if y[i] <= target:
            if y[i] == y[i - 1]:
                return float(x[i]), f"下降至 {ratio:.2f}Fmax"
            t = (target - y[i - 1]) / (y[i] - y[i - 1])
            xu = x[i - 1] + t * (x[i] - x[i - 1])
            return float(xu), f"下降至 {ratio:.2f}Fmax"
    return float(x[-1]), f"未下降至 {ratio:.2f}Fmax，取最大骨架位移"


def ductility_for_branch(skeleton: pd.DataFrame, direction: str, cfg: ProcessConfig) -> pd.DataFrame:
    br = prepare_branch(skeleton, direction)
    rows = []
    if len(br) < 3:
        return pd.DataFrame([{
            "Direction": direction,
            "Method": "all",
            "Status": "骨架点不足，无法稳定计算延性系数",
        }])

    x = br["x"].to_numpy(dtype=float)
    y = br["y"].to_numpy(dtype=float)
    valid = (x >= 0) & (y >= 0)
    x, y = x[valid], y[valid]

    if len(x) < 3 or np.nanmax(y) <= 0:
        return pd.DataFrame([{
            "Direction": direction,
            "Method": "all",
            "Status": "骨架曲线无有效正向承载力",
        }])

    fmax = float(np.max(y))
    d_at_fmax = float(x[np.argmax(y)])
    if getattr(cfg, "init_stiffness_method", "低荷载线性拟合") == "首个有效骨架点割线":
        valid_k = np.where((x > 1e-12) & (y > 0))[0]
        k0 = float(y[valid_k[0]] / x[valid_k[0]]) if len(valid_k) else np.nan
    else:
        k0 = estimate_initial_stiffness(x, y, cfg.init_stiffness_force_ratio)
    du, u_status = ultimate_displacement(x, y, cfg.ultimate_strength_ratio)
    if "未下降" in str(u_status) and getattr(cfg, "ultimate_displacement_policy", "未下降则取最大骨架位移") == "未下降则不计算延性":
        du = np.nan
        u_status = u_status + "；按用户设置，延性系数不计算"

    def add_row(method: str, dy: Optional[float], fy: Optional[float], status: str):
        mu = du / dy if dy and dy > 1e-12 and np.isfinite(du) else np.nan
        rows.append({
            "Direction": direction,
            "Method": method,
            "K0": k0,
            "Fmax": fmax,
            "D_at_Fmax": d_at_fmax,
            "Dy": dy if dy is not None else np.nan,
            "Fy": fy if fy is not None else np.nan,
            "Du": du,
            "Ductility_mu": mu,
            "UltimateStatus": u_status,
            "Status": status,
        })

    # 方法1：割线刚度退化法，取 Ksec <= 0.75K0 的首点
    if np.isfinite(k0) and k0 > 0:
        ksec = np.divide(y, x, out=np.full_like(y, np.nan), where=x > 1e-12)
        idx = np.where((x > 0) & (ksec <= 0.75 * k0))[0]
        if len(idx) > 0:
            i = int(idx[0])
            add_row("割线刚度退化法(Ksec≤0.75K0)", float(x[i]), float(y[i]), "可计算")
        else:
            add_row("割线刚度退化法(Ksec≤0.75K0)", np.nan, np.nan, "未出现 Ksec≤0.75K0")
    else:
        add_row("割线刚度退化法(Ksec≤0.75K0)", np.nan, np.nan, "K0 无法计算")

    # 方法2：能量等效法 / EEEP 近似
    if np.isfinite(k0) and k0 > 0 and du > 0:
        area = area_until_x(x, y, du)
        # area = Fy*du - Fy^2/(2K0)
        # Fy^2/(2K0) - du*Fy + area = 0
        a = 1.0 / (2 * k0)
        b = -du
        c = area
        disc = b * b - 4 * a * c
        if disc >= 0:
            fy1 = (-b - math.sqrt(disc)) / (2 * a)
            fy2 = (-b + math.sqrt(disc)) / (2 * a)
            candidates = [fy for fy in [fy1, fy2] if fy > 0 and fy <= 1.2 * fmax]
            if candidates:
                fy = min(candidates, key=lambda z: abs(z - 0.8 * fmax))
                dy = fy / k0
                add_row("能量等效法(EEEP近似)", float(dy), float(fy), "可计算")
            else:
                add_row("能量等效法(EEEP近似)", np.nan, np.nan, "二次方程根不合理")
        else:
            add_row("能量等效法(EEEP近似)", np.nan, np.nan, "能量方程无实根")
    else:
        add_row("能量等效法(EEEP近似)", np.nan, np.nan, "K0 或 Du 无法计算")

    # 方法3：Park 近似：以 0.75Fmax 点割线外推至 Fmax
    d075 = interpolate_x_at_y(x, y, 0.75 * fmax)
    if d075 and d075 > 1e-12:
        k075 = 0.75 * fmax / d075
        dy = fmax / k075 if k075 > 0 else np.nan
        add_row("Park近似法(0.75Fmax割线外推)", float(dy), fmax, "可计算")
    else:
        add_row("Park近似法(0.75Fmax割线外推)", np.nan, np.nan, "无法插值得到0.75Fmax位移")

    return pd.DataFrame(rows)


def compute_ductility(skeleton: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    pos = ductility_for_branch(skeleton, "positive", cfg)
    neg = ductility_for_branch(skeleton, "negative", cfg)
    return pd.concat([pos, neg], ignore_index=True)


# -----------------------------
# 绘图与 GIF
# -----------------------------

def thin_for_plot(df: pd.DataFrame, max_points: int) -> pd.DataFrame:
    if len(df) <= max_points:
        return df
    idx = np.linspace(0, len(df) - 1, max_points).astype(int)
    return df.iloc[np.unique(idx)]


def add_zero_axes(cfg: ProcessConfig):
    if getattr(cfg, "show_axis_zero_line", True):
        plt.axhline(0, linewidth=0.8)
        plt.axvline(0, linewidth=0.8)


def save_figure(path: Path, cfg: ProcessConfig):
    kwargs = {"dpi": int(getattr(cfg, "figure_dpi", 160))}
    if path.suffix.lower() in [".jpg", ".jpeg"]:
        kwargs["pil_kwargs"] = {"quality": 95}
    plt.tight_layout()
    plt.savefig(path, **kwargs)
    plt.close()


def save_hysteresis_plot(raw: pd.DataFrame, processed: pd.DataFrame, out_path: Path, cfg: ProcessConfig):
    setup_matplotlib_style(cfg)
    p1 = thin_for_plot(raw, cfg.max_plot_points)
    p2 = thin_for_plot(processed, cfg.max_plot_points)

    plt.figure(figsize=(7.5, 5.5))
    plt.plot(p1["D"], p1["F"], linewidth=max(0.5, cfg.line_width * 0.8), alpha=0.55, label=T(cfg, "raw"))
    plt.plot(p2["D"], p2["F_smooth"], linewidth=cfg.line_width, label=T(cfg, "processed"))
    add_zero_axes(cfg)
    plt.xlabel(f"{T(cfg, 'disp')} / {cfg.disp_unit}")
    plt.ylabel(f"{T(cfg, 'force')} / {cfg.force_unit}")
    plt.title(T(cfg, "hys_title"))
    plt.legend()
    save_figure(out_path, cfg)


def save_loop_plot(loop_data: Dict[int, pd.DataFrame], out_path: Path, cfg: ProcessConfig):
    setup_matplotlib_style(cfg)
    plt.figure(figsize=(7.5, 5.5))
    for loop_id, seg in loop_data.items():
        p = thin_for_plot(seg, max(300, cfg.max_plot_points // max(1, len(loop_data))))
        plt.plot(p["D"], p["F_smooth"], linewidth=max(0.5, cfg.line_width * 0.8), alpha=0.7)
    add_zero_axes(cfg)
    plt.xlabel(f"{T(cfg, 'disp')} / {cfg.disp_unit}")
    plt.ylabel(f"{T(cfg, 'force')} / {cfg.force_unit}")
    plt.title(T(cfg, "loops_title"))
    save_figure(out_path, cfg)


def save_skeleton_plot(sk_disp: pd.DataFrame, sk_force: pd.DataFrame, sk_env: pd.DataFrame, out_path: Path, cfg: ProcessConfig):
    setup_matplotlib_style(cfg)
    plt.figure(figsize=(7.5, 5.5))
    if not sk_disp.empty:
        sd = sk_disp.sort_values("D")
        plt.plot(sd["D"], sd["F"], marker="o", linewidth=cfg.line_width, markersize=cfg.marker_size, label=T(cfg, "disp_skeleton"))
    if not sk_force.empty:
        sf = sk_force.sort_values("D")
        plt.plot(sf["D"], sf["F"], marker="s", linewidth=cfg.line_width, markersize=cfg.marker_size, label=T(cfg, "force_skeleton"))
    if not sk_env.empty:
        se = sk_env.sort_values("D")
        plt.plot(se["D"], se["F"], linewidth=cfg.line_width, label=T(cfg, "envelope_skeleton"))
    add_zero_axes(cfg)
    plt.xlabel(f"{T(cfg, 'disp')} / {cfg.disp_unit}")
    plt.ylabel(f"{T(cfg, 'force')} / {cfg.force_unit}")
    plt.title(T(cfg, "skeleton_title"))
    plt.legend()
    save_figure(out_path, cfg)


def save_indicator_plots(info: pd.DataFrame, out_dir: Path, cfg: ProcessConfig) -> List[Path]:
    setup_matplotlib_style(cfg)
    paths = []
    if info.empty:
        return paths

    if getattr(cfg, "plot_cumulative_energy", True):
        p = fig_path(out_dir, "cumulative_energy", cfg)
        plt.figure(figsize=(7.2, 4.8))
        plt.plot(info["LoopID"], info["CumulativeEnergy"], marker="o", linewidth=cfg.line_width, markersize=cfg.marker_size)
        plt.xlabel(T(cfg, "loop_id"))
        plt.ylabel(f"{T(cfg, 'cum_energy')} / {cfg.force_unit}·{cfg.disp_unit}")
        plt.title(T(cfg, "cum_energy_title"))
        save_figure(p, cfg)
        paths.append(p)

    if getattr(cfg, "plot_secant_stiffness", True):
        p = fig_path(out_dir, "secant_stiffness", cfg)
        plt.figure(figsize=(7.2, 4.8))
        plt.plot(info["LoopID"], info["SecantStiffness_K"], marker="o", linewidth=cfg.line_width, markersize=cfg.marker_size, label="K")
        plt.plot(info["LoopID"], info["SecantStiffness_pos"], marker="s", linewidth=cfg.line_width, markersize=cfg.marker_size, label="K+")
        plt.plot(info["LoopID"], info["SecantStiffness_neg"], marker="^", linewidth=cfg.line_width, markersize=cfg.marker_size, label="K-")
        plt.xlabel(T(cfg, "loop_id"))
        plt.ylabel(f"{T(cfg, 'stiffness')} / {cfg.force_unit}/{cfg.disp_unit}")
        plt.title(T(cfg, "stiffness_title"))
        plt.legend()
        save_figure(p, cfg)
        paths.append(p)

    if getattr(cfg, "plot_equivalent_damping", True):
        p = fig_path(out_dir, "equivalent_damping", cfg)
        plt.figure(figsize=(7.2, 4.8))
        plt.plot(info["LoopID"], info["EquivalentViscousDamping_heq"], marker="o", linewidth=cfg.line_width, markersize=cfg.marker_size)
        plt.xlabel(T(cfg, "loop_id"))
        plt.ylabel("h_eq")
        plt.title(T(cfg, "heq_title"))
        save_figure(p, cfg)
        paths.append(p)

    return paths


def save_gif(processed: pd.DataFrame, out_path: Path, cfg: ProcessConfig):
    if imageio is None:
        return
    if not cfg.make_gif or len(processed) < 5:
        return

    setup_matplotlib_style(cfg)
    d = processed["D"].to_numpy()
    f = processed["F_smooth"].to_numpy()
    step = max(2, int(cfg.gif_step))
    frames = []
    tmpdir = out_path.parent / "_gif_frames"
    tmpdir.mkdir(exist_ok=True)

    xmin, xmax = float(np.nanmin(d)), float(np.nanmax(d))
    ymin, ymax = float(np.nanmin(f)), float(np.nanmax(f))
    dx = max((xmax - xmin) * 0.08, 1e-6)
    dy = max((ymax - ymin) * 0.08, 1e-6)

    frame_paths = []
    for end in range(step, len(processed) + step, step):
        end = min(end, len(processed))
        plt.figure(figsize=(6.5, 4.8))
        plt.plot(d[:end], f[:end], linewidth=cfg.line_width)
        plt.scatter(d[end - 1], f[end - 1], s=max(8, cfg.marker_size * 3))
        add_zero_axes(cfg)
        plt.xlim(xmin - dx, xmax + dx)
        plt.ylim(ymin - dy, ymax + dy)
        plt.xlabel(f"{T(cfg, 'disp')} / {cfg.disp_unit}")
        plt.ylabel(f"{T(cfg, 'force')} / {cfg.force_unit}")
        plt.title(f"{T(cfg, 'animation_title')}  {end}/{len(processed)}")
        plt.tight_layout()
        fp = tmpdir / f"frame_{len(frame_paths):04d}.png"
        plt.savefig(fp, dpi=110)
        plt.close()
        frame_paths.append(fp)

        if len(frame_paths) > 250:
            break

    for fp in frame_paths:
        frames.append(imageio.imread(fp))
    imageio.mimsave(out_path, frames, duration=max(10, cfg.gif_interval_ms) / 1000.0)

    for fp in frame_paths:
        try:
            fp.unlink()
        except Exception:
            pass
    try:
        tmpdir.rmdir()
    except Exception:
        pass


# -----------------------------
# 输出 Excel
# -----------------------------

def autosize_excel_columns(writer, sheet_name: str, df: pd.DataFrame):
    try:
        ws = writer.sheets[sheet_name]
        for i, col in enumerate(df.columns):
            max_len = max([len(str(col))] + [len(str(v)) for v in df[col].head(200).values]) if not df.empty else len(str(col))
            ws.set_column(i, i, min(max(max_len + 2, 10), 32))
    except Exception:
        pass


def write_sheet(writer, df: pd.DataFrame, sheet_name: str, cfg: ProcessConfig, index: bool = False, **kwargs):
    out_df = maybe_translate_columns(df.copy(), cfg) if isinstance(df, pd.DataFrame) else df
    out_df.to_excel(writer, sheet_name=sheet_name, index=index, **kwargs)
    if isinstance(out_df, pd.DataFrame):
        autosize_excel_columns(writer, sheet_name, out_df)


def make_level_summary(info: pd.DataFrame) -> pd.DataFrame:
    if info.empty:
        return pd.DataFrame()
    return info.groupby("AmplitudeLevel", as_index=False).agg(
        loops=("LoopID", "count"),
        mean_area=("LoopArea_Eh", "mean"),
        cum_area=("LoopArea_Eh", "sum"),
        mean_K=("SecantStiffness_K", "mean"),
        mean_heq=("EquivalentViscousDamping_heq", "mean"),
        Fpos_max=("F_at_Dmax_pos", "max"),
        Fneg_absmax=("F_at_Dmin_neg", lambda s: np.max(np.abs(s))),
    )


def write_excel_report(
    out_path: Path,
    cfg: ProcessConfig,
    preprocess_log: Dict[str, Any],
    raw: pd.DataFrame,
    processed: pd.DataFrame,
    info: pd.DataFrame,
    loop_data: Dict[int, pd.DataFrame],
    sk_disp: pd.DataFrame,
    sk_force: pd.DataFrame,
    sk_env: pd.DataFrame,
    duct: pd.DataFrame,
):
    params_df = pd.DataFrame(list(asdict(cfg).items()), columns=["参数", "取值"])
    log_df = pd.DataFrame(list(preprocess_log.items()), columns=["项目", "取值"])
    peak_events = significant_peak_events(processed, cfg)
    loops_concat = pd.concat(loop_data.values(), ignore_index=True) if loop_data else pd.DataFrame()
    level_summary = make_level_summary(info)

    with pd.ExcelWriter(out_path, engine="xlsxwriter") as writer:
        params_df.to_excel(writer, sheet_name="Parameters", index=False, startrow=0)
        log_df.to_excel(writer, sheet_name="Parameters", index=False, startrow=len(params_df) + 3)
        autosize_excel_columns(writer, "Parameters", params_df)

        if getattr(cfg, "export_raw_data", True):
            write_sheet(writer, raw, "RawData", cfg)
        if getattr(cfg, "export_processed_data", True):
            write_sheet(writer, processed, "ProcessedData", cfg)
        if getattr(cfg, "export_detected_peaks", True):
            write_sheet(writer, peak_events, "DetectedPeaks", cfg)
        if getattr(cfg, "export_loop_info", True):
            write_sheet(writer, info, "LoopInfo", cfg)
        if getattr(cfg, "export_separated_loops", True):
            write_sheet(writer, loops_concat, "SeparatedLoops", cfg)
        if getattr(cfg, "export_skeleton_data", True):
            write_sheet(writer, sk_disp, "Skeleton_DispPeak", cfg)
            write_sheet(writer, sk_force, "Skeleton_ForcePeak", cfg)
            write_sheet(writer, sk_env, "Skeleton_Envelope", cfg)
        if getattr(cfg, "export_ductility", True):
            write_sheet(writer, duct, "Ductility", cfg)
        if getattr(cfg, "export_level_summary", True):
            write_sheet(writer, level_summary, "LevelSummary", cfg)

        workbook = writer.book
        readme = workbook.add_worksheet("README")
        if getattr(cfg, "output_language", "中文") == "English":
            lines = [
                "houlai_hysteresis analysis results",
                "",
                "1. LoopInfo includes peak points, hysteretic area, cumulative energy, secant stiffness, equivalent viscous damping, and residual deformation.",
                "2. DetectedPeaks lists the displacement peaks used by the loading-level peak-pairing algorithm.",
                "3. Skeleton_DispPeak uses displacement peak points; Skeleton_ForcePeak uses force peak points; Skeleton_Envelope gives an envelope curve.",
                "4. Ductility reports approximate ductility coefficients. If the curve does not drop to the selected strength ratio, the output follows the user-selected ultimate displacement policy.",
                "5. For timber Dougong or mortise-tenon tests, automatic loop separation should always be checked against the loading protocol.",
            ]
        else:
            lines = [
                "houlai_hysteresis 滞回曲线处理结果说明", "",
                "1. LoopInfo：每圈滞回环指标，包括峰值、面积、累计耗能、割线刚度、等效黏滞阻尼系数、残余变形等。",
                "2. DetectedPeaks：加载级位移峰值法识别到的正、负向峰值点，可用于检查自动分圈是否可靠。",
                "3. Skeleton_DispPeak：位移峰值点骨架曲线；Skeleton_ForcePeak：荷载峰值点骨架曲线；Skeleton_Envelope：滞回包络骨架曲线。",
                "4. Ductility：延性系数近似计算结果。若未下降至设定强度系数，程序会按用户选择处理。",
                "5. 木结构斗栱滞回曲线常有捏拢、滑移、间隙、摩擦和千斤顶震荡，自动识别结果应结合试验加载记录人工复核。",
            ]
        for r, text in enumerate(lines):
            readme.write(r, 0, text)
        readme.set_column(0, 0, 120)


# -----------------------------
# 主处理流程
# -----------------------------

def select_ductility_skeleton(sk_disp: pd.DataFrame, sk_force: pd.DataFrame, sk_env: pd.DataFrame, cfg: ProcessConfig) -> pd.DataFrame:
    source = getattr(cfg, "ductility_skeleton_source", "位移峰值骨架")
    if source == "荷载峰值骨架" and not sk_force.empty:
        return sk_force
    if source == "包络骨架" and not sk_env.empty:
        return sk_env
    if not sk_disp.empty:
        return sk_disp
    if not sk_force.empty:
        return sk_force
    return sk_env


def process_one_file(file_obj, cfg: ProcessConfig, out_root: Path) -> Dict[str, Any]:
    base = sanitize_filename(getattr(file_obj, "name", "hysteresis"))
    folder = out_root / base
    folder.mkdir(parents=True, exist_ok=True)

    df_input = read_input_table(file_obj, cfg)
    raw = extract_xy(df_input, cfg)
    processed, log = preprocess(raw, cfg)

    loops = segment_loops(processed, cfg)
    info, loop_data = compute_loop_info(processed, loops, cfg)

    sk_disp = skeleton_by_displacement_peak(processed, info, cfg) if not info.empty else pd.DataFrame()
    sk_force = skeleton_by_force_peak(processed, info, cfg) if not info.empty else pd.DataFrame()
    sk_env = skeleton_envelope(processed) if not processed.empty else pd.DataFrame()
    duct_source = select_ductility_skeleton(sk_disp, sk_force, sk_env, cfg)
    duct = compute_ductility(duct_source, cfg) if not duct_source.empty else pd.DataFrame()

    excel_path = folder / f"{base}_houlai_hysteresis_results.xlsx"
    write_excel_report(
        excel_path, cfg, log, raw, processed, info, loop_data,
        sk_disp, sk_force, sk_env, duct
    )

    figs = []
    if getattr(cfg, "plot_raw_processed", True):
        p = fig_path(folder, "hysteresis_raw_processed", cfg)
        save_hysteresis_plot(raw, processed, p, cfg)
        figs.append(p)

    if getattr(cfg, "plot_separated_loops", True):
        p = fig_path(folder, "separated_loops", cfg)
        save_loop_plot(loop_data, p, cfg)
        figs.append(p)

    if getattr(cfg, "plot_skeleton", True):
        p = fig_path(folder, "skeleton_curves", cfg)
        save_skeleton_plot(sk_disp, sk_force, sk_env, p, cfg)
        figs.append(p)

    figs.extend(save_indicator_plots(info, folder, cfg))

    gif_path = folder / "hysteresis_animation.gif"
    try:
        save_gif(processed, gif_path, cfg)
    except Exception:
        pass

    return {
        "base": base,
        "folder": folder,
        "excel": excel_path,
        "figures": figs,
        "gif": gif_path if gif_path.exists() else None,
        "raw": raw,
        "processed": processed,
        "loop_info": info,
        "sk_disp": sk_disp,
        "sk_force": sk_force,
        "sk_env": sk_env,
        "ductility": duct,
        "preprocess_log": log,
    }


def make_zip(root: Path) -> bytes:
    bio = io.BytesIO()
    with zipfile.ZipFile(bio, "w", zipfile.ZIP_DEFLATED) as zf:
        for p in root.rglob("*"):
            if p.is_file():
                zf.write(p, arcname=str(p.relative_to(root)))
    bio.seek(0)
    return bio.read()


# -----------------------------
# Streamlit 界面
# -----------------------------

# -----------------------------
# Streamlit 界面：v3.1 可视化增强版
# -----------------------------

def streamlit_app():
    st.set_page_config(page_title="houlai_hysteresis v1.0 在线增强版", layout="wide")

    # 顶部可见版本标识，防止误以为仍是旧版
    st.markdown(
        """
        <div style="padding:16px 20px;border-radius:14px;background:linear-gradient(90deg,#eef5ff,#f8fbff);border:1px solid #d8e7ff;margin-bottom:14px">
            <div style="font-size:26px;font-weight:800;line-height:1.35">houlai_hysteresis 滞回曲线处理与抗震评价程序 <span style="color:#2563eb">v1.0 在线增强版</span></div>
            <div style="font-size:14px;color:#475569;margin-top:6px">主界面已改为标签页：数据导入、预处理、分圈、指标计算、输出设置、运行导出。打开网页即可看到新版界面。</div>
        </div>
        """,
        unsafe_allow_html=True,
    )

    st.sidebar.success("当前版本：houlai_hysteresis v1.0 在线增强版")
    output_language = st.sidebar.radio("图表输出语言 / Figure language", ["中文", "English"], index=0, horizontal=True)
    preview_cfg = ProcessConfig(disp_col="", force_col="", output_language=output_language)
    st.caption(T(preview_cfg, "app_caption"))

    preset = st.sidebar.selectbox(
        "参数预设",
        ["斗栱分级加载：5,10,15,20,25,30", "通用低周反复加载", "自定义"],
        index=0,
    )
    if preset == "斗栱分级加载：5,10,15,20,25,30":
        preset_levels = "5,10,15,20,25,30"
        preset_loop_method = "加载级位移峰值法（推荐）"
        preset_level_tol = 2.0
        preset_prom = 1.0
        preset_dist = 20
        preset_min_peak = 2.0
    else:
        preset_levels = ""
        preset_loop_method = "加载级位移峰值法（推荐）"
        preset_level_tol = 1.0
        preset_prom = 0.5
        preset_dist = 20
        preset_min_peak = 0.0

    tab_data, tab_pre, tab_loop, tab_calc, tab_out, tab_run = st.tabs([
        "① 数据导入", "② 预处理", "③ 分圈设置", "④ 指标计算", "⑤ 输出设置", "⑥ 运行导出"
    ])

    # 默认值容器
    files = []
    has_header = True
    sheet_name = None
    disp_col = ""
    force_col = ""
    disp_unit = "mm"
    force_unit = "kN"
    zero_origin = False

    with tab_data:
        st.subheader("① 数据导入与列选择")
        st.info("新版 v3.1：这里可以直接选择文件、工作表、位移列和荷载/反力列。后面的标签页可以修改计算与输出决策。")
        files = st.file_uploader(
            "上传 Excel/CSV/TXT 文件，可多选",
            type=["xlsx", "xls", "xlsm", "csv", "txt", "dat"],
            accept_multiple_files=True,
        )
        c0, c1, c2 = st.columns([1, 1, 1])
        with c0:
            has_header = st.checkbox("文件首行为表头", value=True)
        with c1:
            disp_unit = st.text_input("位移单位", "mm")
        with c2:
            force_unit = st.text_input("力单位", "kN")
        zero_origin = st.checkbox("将首点平移到坐标原点", value=False)

        if files:
            first_file = files[0]
            suffix = Path(first_file.name).suffix.lower()
            if suffix in [".xlsx", ".xls", ".xlsm"]:
                sheets = list_excel_sheets(first_file)
                sheet_name = st.selectbox("Excel 工作表", sheets, index=0) if sheets else None
            temp_cfg_for_read = ProcessConfig(disp_col="", force_col="", sheet_name=sheet_name, has_header=has_header)
            first_df = read_input_table(first_file, temp_cfg_for_read)
            try:
                first_file.seek(0)
            except Exception:
                pass
            if first_df.shape[1] < 2:
                st.error("文件至少需要两列数据。")
                return
            columns = list(first_df.columns)
            dcol, fcol = st.columns(2)
            with dcol:
                disp_col = st.selectbox("位移列", columns, index=0)
            with fcol:
                force_col = st.selectbox("荷载/反力列", columns, index=1 if len(columns) > 1 else 0)
            st.dataframe(first_df.head(8), use_container_width=True)
        else:
            st.warning("请先上传数据文件；未上传前仍可先设置后续参数。")

    with tab_pre:
        st.subheader("② 预处理参数")
        p1, p2, p3 = st.columns(3)
        with p1:
            remove_pause = st.checkbox("启用停顿点修正", value=True)
            pause_dx = st.number_input("停顿点阈值 dx", value=0.02, min_value=0.0, step=0.01, format="%.6f")
            pause_dy = st.number_input("停顿点阈值 dy", value=0.02, min_value=0.0, step=0.01, format="%.6f")
        with p2:
            downsample_step = st.number_input("隔行取数步长", value=1, min_value=1, step=1)
            smooth_method = st.selectbox("力值平滑方法", ["不平滑", "移动平均", "移动中值", "Savitzky-Golay", "控制点插值"], index=0)
        with p3:
            smooth_window = st.number_input("平滑窗口/点数", value=11, min_value=3, step=2)
            savgol_polyorder = st.number_input("Savitzky-Golay 多项式阶数", value=3, min_value=1, max_value=6, step=1)
            control_points = st.number_input("控制点数量", value=80, min_value=4, step=5)
        st.caption("如果论文要求基于原始数据，建议保持‘不平滑’；平滑结果可作为对照分析。")

    with tab_loop:
        st.subheader("③ 滞回环分圈设置")
        l1, l2, l3 = st.columns(3)
        with l1:
            loop_method = st.selectbox("分圈方法", ["加载级位移峰值法（推荐）", "零位移穿越法", "位移转折点法"], index=["加载级位移峰值法（推荐）", "零位移穿越法", "位移转折点法"].index(preset_loop_method))
            loading_levels = st.text_input("目标加载位移级别", preset_levels)
            level_tolerance = st.number_input("位移级别合并容差", value=float(preset_level_tol), min_value=0.0001, step=0.5)
        with l2:
            zero_eps = st.number_input("零位移容差", value=1e-6, min_value=0.0, format="%.8f")
            min_points_per_loop = st.number_input("每圈最少点数", value=15, min_value=5, step=1)
        with l3:
            peak_prominence = st.number_input("峰值识别显著性 prominence", value=float(preset_prom), min_value=0.0, step=0.5)
            peak_distance = st.number_input("相邻峰值最小点距", value=int(preset_dist), min_value=3, step=5)
            min_peak_abs = st.number_input("最小峰值绝对位移", value=float(preset_min_peak), min_value=0.0, step=0.5)
        st.info("对 5/10/15/20/25/30 mm 这类分级加载，推荐使用‘加载级位移峰值法’。")

    with tab_calc:
        st.subheader("④ 指标计算决策")
        c1, c2, c3 = st.columns(3)
        with c1:
            skeleton_first_cycle_only = st.checkbox("骨架曲线仅取同级位移第一圈", value=True)
            loop_area_method = st.selectbox("滞回环面积算法", ["路径积分trapz取绝对值", "多边形闭合面积", "分段绝对积分"], index=0)
            stiffness_peak_reference = st.selectbox("割线刚度/阻尼采用的峰值点", ["位移峰值点", "荷载峰值点"], index=0)
        with c2:
            damping_formula = st.selectbox("等效黏滞阻尼计算", ["JGJ常用三角形能量法", "不计算"], index=0)
            ductility_skeleton_source = st.selectbox("延性系数采用的骨架曲线", ["位移峰值骨架", "荷载峰值骨架", "包络骨架"], index=0)
            init_stiffness_method = st.selectbox("初始刚度计算方法", ["低荷载线性拟合", "首个有效骨架点割线"], index=0)
        with c3:
            init_stiffness_force_ratio = st.slider("初始刚度拟合上限 F/Fmax", 0.1, 0.8, 0.4, 0.05)
            ultimate_strength_ratio = st.slider("极限强度系数", 0.5, 1.0, 0.85, 0.01)
            ultimate_displacement_policy = st.selectbox("骨架未下降至极限强度时", ["未下降则取最大骨架位移", "未下降则不计算延性"], index=0)

    with tab_out:
        st.subheader("⑤ 输出设置：图表语言、格式、工作表")
        o1, o2, o3 = st.columns(3)
        with o1:
            figure_format = st.selectbox("图片格式", ["png", "jpg", "svg", "pdf"], index=0)
            figure_dpi = st.number_input("图片 DPI", value=180, min_value=72, max_value=600, step=10)
            font_family = st.text_input("中文字体优先项", "Microsoft YaHei")
            excel_column_language = st.selectbox("Excel 列名", ["保留英文变量名", "中文列名"], index=0)
        with o2:
            line_width = st.number_input("线宽", value=1.0, min_value=0.2, max_value=5.0, step=0.1)
            marker_size = st.number_input("骨架点大小", value=3.0, min_value=0.0, max_value=12.0, step=0.5)
            show_axis_zero_line = st.checkbox("显示坐标零轴", value=True)
            max_plot_points = st.number_input("绘图最大点数", value=15000, min_value=1000, max_value=200000, step=1000)
        with o3:
            make_gif = st.checkbox("生成 GIF 动图", value=False)
            gif_step = st.number_input("GIF 动画增量步", value=30, min_value=2, step=5)
            gif_interval_ms = st.number_input("GIF 每帧间隔 ms", value=60, min_value=10, step=10)

        st.markdown("**选择输出图片**")
        g1, g2, g3 = st.columns(3)
        with g1:
            plot_raw_processed = st.checkbox("原始/处理后滞回曲线对比图", value=True)
            plot_separated_loops = st.checkbox("滞回环分解图", value=True)
        with g2:
            plot_skeleton = st.checkbox("骨架曲线图", value=True)
            plot_cumulative_energy = st.checkbox("累计耗能曲线", value=True)
        with g3:
            plot_secant_stiffness = st.checkbox("割线刚度退化曲线", value=True)
            plot_equivalent_damping = st.checkbox("等效黏滞阻尼曲线", value=True)

        st.markdown("**选择 Excel 工作表**")
        e1, e2, e3, e4 = st.columns(4)
        with e1:
            export_raw_data = st.checkbox("RawData 原始数据", value=True)
            export_processed_data = st.checkbox("ProcessedData 处理后数据", value=True)
        with e2:
            export_detected_peaks = st.checkbox("DetectedPeaks 峰值点", value=True)
            export_loop_info = st.checkbox("LoopInfo 滞回环指标", value=True)
        with e3:
            export_separated_loops = st.checkbox("SeparatedLoops 分圈数据", value=True)
            export_skeleton_data = st.checkbox("Skeleton 骨架曲线数据", value=True)
        with e4:
            export_ductility = st.checkbox("Ductility 延性系数", value=True)
            export_level_summary = st.checkbox("LevelSummary 分级摘要", value=True)

    # 如果还没上传，提供占位值，避免 cfg 构造失败
    if not files:
        disp_col = disp_col or "位移"
        force_col = force_col or "反力"

    cfg = ProcessConfig(
        disp_col=str(disp_col),
        force_col=str(force_col),
        sheet_name=sheet_name,
        has_header=has_header,
        disp_unit=disp_unit,
        force_unit=force_unit,
        zero_origin=zero_origin,
        remove_pause=remove_pause,
        pause_dx=float(pause_dx),
        pause_dy=float(pause_dy),
        downsample_step=int(downsample_step),
        smooth_method=smooth_method,
        smooth_window=int(smooth_window),
        savgol_polyorder=int(savgol_polyorder),
        control_points=int(control_points),
        loop_method=loop_method,
        zero_eps=float(zero_eps),
        min_points_per_loop=int(min_points_per_loop),
        level_tolerance=float(level_tolerance),
        loading_levels=str(loading_levels),
        peak_prominence=float(peak_prominence),
        peak_distance=int(peak_distance),
        min_peak_abs=float(min_peak_abs),
        skeleton_first_cycle_only=bool(skeleton_first_cycle_only),
        loop_area_method=loop_area_method,
        stiffness_peak_reference=stiffness_peak_reference,
        damping_formula=damping_formula,
        ductility_skeleton_source=ductility_skeleton_source,
        init_stiffness_method=init_stiffness_method,
        init_stiffness_force_ratio=float(init_stiffness_force_ratio),
        ultimate_strength_ratio=float(ultimate_strength_ratio),
        ultimate_displacement_policy=ultimate_displacement_policy,
        output_language=output_language,
        excel_column_language=excel_column_language,
        figure_format=figure_format,
        figure_dpi=int(figure_dpi),
        font_family=font_family,
        line_width=float(line_width),
        marker_size=float(marker_size),
        show_axis_zero_line=bool(show_axis_zero_line),
        max_plot_points=int(max_plot_points),
        plot_raw_processed=plot_raw_processed,
        plot_separated_loops=plot_separated_loops,
        plot_skeleton=plot_skeleton,
        plot_cumulative_energy=plot_cumulative_energy,
        plot_secant_stiffness=plot_secant_stiffness,
        plot_equivalent_damping=plot_equivalent_damping,
        export_raw_data=export_raw_data,
        export_processed_data=export_processed_data,
        export_detected_peaks=export_detected_peaks,
        export_loop_info=export_loop_info,
        export_separated_loops=export_separated_loops,
        export_skeleton_data=export_skeleton_data,
        export_ductility=export_ductility,
        export_level_summary=export_level_summary,
        make_gif=make_gif,
        gif_step=int(gif_step),
        gif_interval_ms=int(gif_interval_ms),
    )

    with tab_run:
        st.subheader("⑥ 运行、预览与导出")
        st.write("当前主要参数预览：")
        st.json({
            "版本": "v3.1 可视化增强版",
            "图表语言": output_language,
            "分圈方法": loop_method,
            "目标加载级别": loading_levels,
            "面积算法": loop_area_method,
            "骨架来源": ductility_skeleton_source,
            "输出图片格式": figure_format,
            "是否生成GIF": make_gif,
        })

        if not files:
            st.warning("请先在“① 数据导入”标签页上传文件。")
            return

        if st.button("开始处理并导出", type="primary", use_container_width=True):
            with st.spinner("正在处理滞回曲线并生成报告……"):
                with tempfile.TemporaryDirectory() as td:
                    out_root = Path(td) / "houlai_hysteresis_output"
                    out_root.mkdir()
                    results = []
                    for f in files:
                        try:
                            f.seek(0)
                        except Exception:
                            pass
                        res = process_one_file(f, cfg, out_root)
                        results.append(res)

                    zip_bytes = make_zip(out_root)
                    st.success(f"处理完成：共 {len(results)} 个文件。")
                    st.download_button(
                        "下载全部结果 ZIP",
                        data=zip_bytes,
                        file_name="houlai_hysteresis_results_v1_0.zip",
                        mime="application/zip",
                        use_container_width=True,
                    )

                    res0 = results[0]
                    st.subheader(f"预览：{res0['base']}")
                    m1, m2, m3, m4 = st.columns(4)
                    with m1:
                        st.metric("原始点数", res0["preprocess_log"].get("原始点数", 0))
                    with m2:
                        st.metric("最终点数", res0["preprocess_log"].get("最终点数", 0))
                    with m3:
                        st.metric("识别峰值点", len(res0.get("peak_info", [])))
                    with m4:
                        st.metric("识别滞回环数", len(res0["loop_info"]))

                    if not res0["loop_info"].empty:
                        st.dataframe(res0["loop_info"].head(20), use_container_width=True)
                    else:
                        st.warning("未识别到完整滞回环。可调整分圈参数或改用其他分圈方法。")

                    preview_images = []
                    for stem, cap in [
                        ("hysteresis_raw_processed", "原始曲线与处理后曲线对比"),
                        ("separated_loops", "滞回环分解"),
                        ("skeleton_curves", "骨架曲线"),
                        ("cumulative_energy", "累计耗能"),
                        ("secant_stiffness", "割线刚度退化"),
                        ("equivalent_damping", "等效黏滞阻尼"),
                    ]:
                        p = fig_path(res0["folder"], stem, cfg)
                        if p.exists() and p.suffix.lower() in [".png", ".jpg", ".jpeg"]:
                            preview_images.append((p, cap))
                    for p, cap in preview_images[:4]:
                        st.image(str(p), caption=cap)


if __name__ == "__main__":
    if st is None:
        print("请先安装 streamlit：pip install streamlit")
    else:
        streamlit_app()
