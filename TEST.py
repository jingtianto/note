import openpyxl
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import re
import io
import warnings
warnings.filterwarnings('ignore')

# ---------------------- 配置区 ----------------------
INPUT_FILE = r"D:\PythonProject\自动分析测试结果\1.xlsx"
OUTPUT_FILE = r"D:\PythonProject\自动分析测试结果\2.xlsx"
TARGET_SHEET_NAME = "IPN"

# 正则
CELL_VALUE_PATTERN = re.compile(r'(~?)==\s*(-?\d+\.?\d*)[dD]?', re.DOTALL)
NOTE_MIN_PATTERN = re.compile(r'Minimum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_MAX_PATTERN = re.compile(r'Maximum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_AVG_PATTERN = re.compile(r'Average:\s*(-?\d+\.?\d*)', re.DOTALL)
TIME_FAILED_PATTERN = re.compile(r't=\[(\d+\.\d+),(\d+\.\d+)\]|LL\+(\d+\.\d+)s', re.DOTALL)

def get_color_from_indexed(cell):
    """indexed=10→红，indexed=11→绿"""
    fill = cell.fill
    if fill.patternType is None or fill.patternType == 'none':
        return 'black'
    fg_color = fill.fgColor
    if hasattr(fg_color, 'indexed'):
        idx = fg_color.indexed
        if idx == 10:
            return 'red'
        elif idx == 11:
            return 'green'
    return 'black'

def parse_cell_value(cell_text):
    if not cell_text:
        return None, False
    s = str(cell_text).strip().replace('\n', '').replace('\r', '')
    if s in ('me', 'None', ''):
        return None, False
    m = CELL_VALUE_PATTERN.search(s)
    if not m:
        return None, False
    is_dashed = bool(m.group(1))
    value_str = m.group(2)
    try:
        val = float(value_str)
    except:
        val = 0.0
    return val, is_dashed

def parse_note_data(note_text):
    if not note_text:
        return 0.0, 0.0, 0.0, [(0.0, 1.0)]
    s = note_text.replace('\n', ' ').replace('\r', ' ').strip()
    min_val = float(m.group(1)) if (m := NOTE_MIN_PATTERN.search(s)) else 0.0
    max_val = float(m.group(1)) if (m := NOTE_MAX_PATTERN.search(s)) else 0.0
    avg_val = float(m.group(1)) if (m := NOTE_AVG_PATTERN.search(s)) else 0.0
    times = []
    for match in TIME_FAILED_PATTERN.findall(s):
        if match[0] and match[1]:
            times.append((round(float(match[0]), 3), round(float(match[1]), 3)))
        elif match[2]:
            times.append((0.0, round(float(match[2]), 3)))
    if not times:
        times = [(0.0, 1.0)]
    return min_val, max_val, avg_val, times

def process_one_case(ws_in, row_idx, wb_out):
    test_id = str(ws_in.cell(row_idx, 1).value or '').strip()
    if not test_id:
        return
    sheet_name = f"TC_{test_id}"
    if sheet_name in wb_out.sheetnames:
        del wb_out[sheet_name]
    ws_out = wb_out.create_sheet(sheet_name)
    signals = []
    max_col = min(ws_in.max_column, 1000)

    for col in range(3, max_col+1):
        signal_name = ws_in.cell(1, col).value
        if not signal_name:
            continue
        data_cell = ws_in.cell(row_idx, col)
        cell_val = data_cell.value
        target_value, is_dashed = parse_cell_value(cell_val)
        if target_value is None:
            continue
        signal_color = get_color_from_indexed(data_cell)
        note_text = data_cell.comment.text if (data_cell.comment and data_cell.comment.text) else ""
        min_val, max_val, avg_val, times = parse_note_data(note_text)
        signals.append({
            "name": signal_name,
            "color": signal_color,
            "target": target_value,
            "dashed": is_dashed,
            "min": min_val,
            "max": max_val,
            "avg": avg_val,
            "times": times
        })

    if not signals:
        return

    n = len(signals)
    fig, axes = plt.subplots(n, 1, figsize=(18, 3*n), sharex=True)
    if n == 1:
        axes = [axes]

    for ax, sig in zip(axes, signals):
        ax.set_title(sig['name'], fontsize=10, loc='left', color=sig['color'], weight='bold')
        ax.grid(True, alpha=0.3)
        ys = [sig["min"], sig["max"], sig["avg"], sig["target"]]
        y_min, y_max = min(ys) - 0.5, max(ys) + 0.5
        ax.set_ylim(y_min, y_max)
        
        # ========== 核心修改：给每条线加标签，让图例自动绑定颜色 ==========
        for t1, t2 in sig["times"]:
            # 实际值（金色）：加label，只在第一次循环显示
            ax.hlines(sig["min"], t1, t2, color='gold', lw=2, label=f"Actual: {sig['min']}~{sig['max']}" if t1==0 else "")
            ax.hlines(sig["max"], t1, t2, color='gold', lw=2)
            ax.fill_between([t1, t2], sig["min"], sig["max"], color='gold', alpha=0.4)
            # 期望值（蓝色）：加label，只在第一次循环显示
            ax.hlines(sig["target"], t1, t2, color='blue',
                      linestyle='--' if sig["dashed"] else '-', lw=1.8, label=f"Exp: {sig['target']}" if t1==0 else "")
            ax.plot((t1+t2)/2, sig["avg"], 'o', color='gold', ms=9)

        # ========== 只显示一次图例，自动匹配颜色 ==========
        ax.legend(fontsize=7, loc='upper left')
        
    axes[-1].set_xlabel("Time (s)")
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    ws_out.add_image(Image(buf), "A1")
    plt.close('all')

# ---------------------- 主程序 ----------------------
if __name__ == "__main__":
    wb_in = openpyxl.load_workbook(INPUT_FILE, data_only=False)
    if TARGET_SHEET_NAME not in wb_in.sheetnames:
        exit(1)
    ws_in = wb_in[TARGET_SHEET_NAME]
    wb_out = openpyxl.Workbook()
    processed = set()
    max_row = ws_in.max_row  # 处理全部行
    
    for r in range(2, max_row+1):
        # 跳过隐藏行
        if ws_in.row_dimensions[r].hidden:
            continue
        
        cid = str(ws_in.cell(r, 1).value or '').strip()
        if not cid or cid in processed:
            continue
        
        process_one_case(ws_in, r, wb_out)
        processed.add(cid)
    
    if 'Sheet' in wb_out.sheetnames:
        del wb_out['Sheet']
    wb_out.save(OUTPUT_FILE)
    print("✅ 处理完成！图例颜色+文字完全匹配（Actual=金色，Exp=蓝色）")
