#!/usr/bin/env python3
"""
炉温曲线分析工具 v1.2
- 导入 GL840 CSV / Snap Cure TXT (GBK) 原始文件
- 解析多通道温度数据，自动检测采样间隔
- 时间-温度曲线图（勾选通道实时刷新，不同颜色）
- 数据间隔提取（线性插值，按秒级间隔提取数据，不可小于原始间隔）
- 导出 Excel（支持勾选转置为横向排列）
- ttkbootstrap flatly 主题，单文件绿色免安装
"""

# ===================== 启动前依赖检查 =====================
import sys
import subprocess
import importlib.util

REQUIRED_PACKAGES = {
    'ttkbootstrap': 'ttkbootstrap',
    'numpy': 'numpy',
    'matplotlib': 'matplotlib',
    'openpyxl': 'openpyxl',
}

def _check_dependencies():
    """检查依赖，缺失时弹出提示并返回缺失列表"""
    missing = []
    for pkg_name, import_name in REQUIRED_PACKAGES.items():
        if importlib.util.find_spec(import_name) is None:
            missing.append(pkg_name)

    if not missing:
        return True

    # 用 tkinter 弹窗提示（Python 自带，必定可用）
    import tkinter as tk
    from tkinter import messagebox
    root = tk.Tk()
    root.withdraw()
    msg = (
        f"缺少以下 Python 依赖包：\n\n"
        f"{' '.join(missing)}\n\n"
        f"请运行以下命令安装：\n"
        f"pip install {' '.join(missing)}\n\n"
        f"如果 pip 不在 PATH 中，请尝试：\n"
        f"python -m pip install {' '.join(missing)}"
    )
    messagebox.showerror("依赖缺失", msg)
    root.destroy()
    return False

if not _check_dependencies():
    sys.exit(1)

# ===================== 正式导入 =====================
import tkinter as tk
from tkinter import ttk, filedialog, messagebox
import ttkbootstrap as ttkb
from ttkbootstrap.constants import *
import os
import re
import math
from datetime import datetime
from dataclasses import dataclass, field
from typing import List, Dict, Tuple, Optional

import numpy as np
from matplotlib.figure import Figure
from matplotlib.backends.backend_tkagg import FigureCanvasTkAgg, NavigationToolbar2Tk
import matplotlib

matplotlib.rcParams['font.sans-serif'] = ['Microsoft YaHei', 'SimHei', 'DejaVu Sans']
matplotlib.rcParams['axes.unicode_minus'] = False

from openpyxl import Workbook
from openpyxl.utils import get_column_letter
from openpyxl.styles import Font, Alignment, PatternFill, Border, Side

# ===================== 常量 =====================
CHANNEL_COLORS = [
    '#1f77b4', '#ff7f0e', '#2ca02c', '#d62728', '#9467bd',
    '#8c564b', '#e377c2', '#7f7f7f', '#bcbd22', '#17becf',
    '#aec7e8', '#ffbb78', '#98df8a', '#ff9896', '#c5b0d5',
    '#c49c94', '#f7b6d2', '#c7c7c7', '#dbdb8d', '#9edae5'
]

WINDOW_WIDTH = 1400
WINDOW_HEIGHT = 900
MIN_WIDTH = 1200
MIN_HEIGHT = 800
LEFT_PANEL_WIDTH = 320


# ===================== 数据模型 =====================
@dataclass
class ProfileData:
    """统一数据模型"""
    file_path: str = ""
    file_type: str = ""  # "gl840" | "snapcure"
    channels: Dict[str, List[Tuple[float, float]]] = field(default_factory=dict)
    # channels: {name: [(time_sec, temperature), ...]}
    original_rate: float = 0.0  # 原始采样率（个/秒）
    metadata: Dict[str, str] = field(default_factory=dict)
    total_points: int = 0
    start_time: Optional[datetime] = None

    @property
    def channel_names(self) -> List[str]:
        return list(self.channels.keys())

    @property
    def time_range(self) -> Tuple[float, float]:
        """返回 (min_time, max_time)"""
        if not self.channels:
            return (0, 0)
        first_ch = next(iter(self.channels.values()))
        if not first_ch:
            return (0, 0)
        return (first_ch[0][0], first_ch[-1][0])

    @property
    def interval_seconds(self) -> float:
        """原始采样间隔（秒）"""
        return 1.0 / self.original_rate if self.original_rate > 0 else 0.0

    @property
    def rate_display(self) -> str:
        """采样间隔显示文本"""
        secs = self.interval_seconds
        if secs <= 0:
            return "-- 秒/次"
        if secs < 60:
            return f"{secs:.1f} 秒/次"
        elif secs < 3600:
            return f"{secs/60:.1f} 分钟/次 ({secs:.0f}秒)"
        else:
            return f"{secs/3600:.1f} 小时/次 ({secs:.0f}秒)"


# ===================== 文件解析器 =====================
def detect_file_type(filepath: str) -> str:
    """自动检测文件类型"""
    # 先尝试 UTF-8
    try:
        with open(filepath, 'r', encoding='utf-8') as f:
            head = f.read(500)
            if 'Vendor' in head and ('GL840' in head or 'GL' in head):
                return 'gl840'
            if '曲线开始时间' in head or 'Snap' in head:
                return 'snapcure'
    except UnicodeDecodeError:
        pass

    # 尝试 GBK（Snap Cure 文件）
    try:
        with open(filepath, 'r', encoding='gbk') as f:
            head = f.read(500)
            if '曲线开始时间' in head or 'Snap' in head or '产品名称' in head:
                return 'snapcure'
    except UnicodeDecodeError:
        pass

    # 回退：按扩展名
    ext = os.path.splitext(filepath)[1].lower()
    if ext == '.csv':
        return 'gl840'
    return 'snapcure'


def _parse_sampling_rate(interval_str: str) -> float:
    """解析采样间隔字符串，返回采样率（个/秒）"""
    interval_str = interval_str.strip().lower().replace(' ', '')
    if not interval_str:
        return 0.0

    # 匹配: "1min", "1s", "500ms", "2h" 等
    match = re.match(r'([\d.]+)\s*(min|ms|s|h|sec|m)?', interval_str)
    if not match:
        try:
            return 1.0 / float(interval_str)
        except ValueError:
            return 0.0

    value = float(match.group(1))
    unit = (match.group(2) or 's').lower()

    if unit in ('min', 'm'):
        secs = value * 60
    elif unit == 'ms':
        secs = value / 1000
    elif unit == 'h':
        secs = value * 3600
    else:  # s, sec
        secs = value

    return 1.0 / secs if secs > 0 else 0.0


def parse_gl840(filepath: str) -> ProfileData:
    """解析 GRAPHTEC GL840 CSV 文件"""
    data = ProfileData(file_path=filepath, file_type='gl840')
    channels: Dict[str, List[Tuple[float, float]]] = {}
    metadata = {}

    with open(filepath, 'r', encoding='utf-8') as f:
        content = f.read()

    lines = content.split('\n')

    # 解析头部元数据
    in_amp = False
    data_start = -1
    channel_count = 0

    for i, line in enumerate(lines):
        line = line.strip()
        if not line:
            continue

        if line.startswith('Sampling interval'):
            parts = line.split(',')
            if len(parts) >= 2:
                metadata['sampling_interval'] = parts[1].strip()
                data.original_rate = _parse_sampling_rate(parts[1].strip())

        elif line.startswith('Start time'):
            parts = line.split(',')
            if len(parts) >= 3:
                try:
                    time_str = f"{parts[1].strip()} {parts[2].strip()}"
                    data.start_time = datetime.strptime(time_str, '%Y-%m-%d %H:%M:%S')
                except ValueError:
                    pass

        elif line.startswith('Total data points'):
            parts = line.split(',')
            if len(parts) >= 2:
                try:
                    data.total_points = int(parts[1].strip())
                except ValueError:
                    pass

        elif line.startswith('AMP settings'):
            in_amp = True
            continue

        elif in_amp and line.startswith('CH'):
            # CH1," CH 1","M",TEMP,TC_K,... — 跳过表头行 "CH,Signal name,..."
            parts = [p.strip().strip('"') for p in line.split(',')]
            ch_name = parts[0].strip()
            if ch_name == 'CH':
                continue  # 跳过 AMP 表头
            if len(parts) >= 2:
                channels[ch_name] = []
                channel_count += 1

        elif line.startswith('Calc settings'):
            in_amp = False

        elif line.startswith('Data'):
            data_start = i
            break

    if data_start < 0:
        raise ValueError("未找到数据区域（Data 标记）")

    # 解析数据行
    if data.original_rate == 0.0 and data.total_points > 0:
        # 如果头部没有采样间隔，从数据中推断
        data.original_rate = 0.017  # 默认 1min

    # 找到数据标题行
    data_header_idx = -1
    for i in range(data_start + 1, min(data_start + 20, len(lines))):
        if lines[i].strip() and ('NO.' in lines[i] or 'Number' in lines[i]):
            data_header_idx = i
            break

    if data_header_idx < 0:
        data_header_idx = data_start + 1

    # 解析数据行
    ch_names = sorted(channels.keys(), key=lambda x: int(re.search(r'\d+', x).group()))
    first_time = None

    for i in range(data_header_idx + 1, len(lines)):
        line = lines[i].strip()
        if not line or line.startswith('*') or line.startswith('#'):
            continue

        parts = [p.strip() for p in line.split(',')]
        if len(parts) < 4 + channel_count:
            continue

        try:
            # 解析时间（Date&Time 是单列: "2026/06/26 11:36:54"）
            datetime_str = parts[1].strip()
            dt = datetime.strptime(datetime_str, '%Y/%m/%d %H:%M:%S')
            if first_time is None:
                first_time = dt
            elapsed = (dt - first_time).total_seconds()

            # 解析通道数据（parts[3] 开始是 CH1, CH2, ...）
            for j, ch_name in enumerate(ch_names):
                idx = 3 + j
                if idx < len(parts):
                    val_str = parts[idx].replace('+', '').replace(' ', '')
                    if val_str and val_str not in ('', '-', '--'):
                        temp = float(val_str)
                        channels[ch_name].append((elapsed, temp))

        except (ValueError, IndexError):
            continue

    # 过滤空通道
    data.channels = {k: v for k, v in channels.items() if v}
    data.metadata = metadata

    # 如果通道数量与数据不匹配，从实际数据行重新推断
    if not data.channels:
        actual_ch_count = 0
        for i in range(data_header_idx + 1, min(data_header_idx + 5, len(lines))):
            parts = [p.strip() for p in lines[i].split(',')]
            if len(parts) > 4:
                # 计算实际温度列数
                temp_count = 0
                for p in parts[3:]:
                    if p and ('+' in p or '-' in p) and not p.startswith('L') and not p.startswith('A'):
                        temp_count += 1
                    else:
                        break
                actual_ch_count = max(actual_ch_count, temp_count)

        if actual_ch_count > 0:
            data.channels = {}
            for j in range(actual_ch_count):
                ch_name = f"CH{j + 1}"
                data.channels[ch_name] = []

            first_time = None
            for i in range(data_header_idx + 1, len(lines)):
                line = lines[i].strip()
                if not line:
                    continue
                parts = [p.strip() for p in line.split(',')]
                if len(parts) < 4 + actual_ch_count:
                    continue
                try:
                    datetime_str = parts[1].strip()
                    dt = datetime.strptime(datetime_str, '%Y/%m/%d %H:%M:%S')
                    if first_time is None:
                        first_time = dt
                    elapsed = (dt - first_time).total_seconds()
                    for j in range(actual_ch_count):
                        idx = 3 + j
                        if idx < len(parts):
                            val_str = parts[idx].replace('+', '').replace(' ', '')
                            if val_str and val_str not in ('', '-', '--'):
                                data.channels[f"CH{j + 1}"].append((elapsed, float(val_str)))
                except (ValueError, IndexError):
                    continue

    data.total_points = max((len(v) for v in data.channels.values()), default=0)
    return data


def parse_snapcure(filepath: str) -> ProfileData:
    """解析 Snap Cure 回流焊 TXT 文件（GBK 编码）"""
    data = ProfileData(file_path=filepath, file_type='snapcure')
    metadata = {}

    with open(filepath, 'r', encoding='gbk', errors='replace') as f:
        content = f.read()

    lines = content.split('\n')

    # 解析头部元数据
    for line in lines[:80]:
        line = line.strip()
        if '产品名称' in line:
            parts = line.split('\t')
            metadata['product'] = parts[-1].strip() if len(parts) > 1 else ''
        elif '制程界限名称' in line:
            parts = line.split('\t')
            metadata['process_limit'] = parts[-1].strip() if len(parts) > 1 else ''
        elif '有效热电偶' in line:
            parts = line.split('\t')
            metadata['active_tc'] = parts[-1].strip() if len(parts) > 1 else ''

    # 查找数据区域："未经处理" 段落
    data_start = -1
    for i, line in enumerate(lines):
        if '未经处理' in line:
            data_start = i
            break

    if data_start < 0:
