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

def get_exact_color_name(cell):
    """
    🔥 兼容所有红绿颜色的识别逻辑
    只要R值远大于G/B → 判定为红色
    只要G值远大于R/B → 判定为绿色
    其他 → 黑色
    """
    fill = cell.fill
    if fill.patternType is None or fill.patternType == 'none':
        print("[颜色识别] 无填充 → 黑色")
        return 'black'
    
    # 获取原始RGB值（处理所有格式）
    r, g, b = 0, 0, 0
    if hasattr(fill.fgColor, 'rgb'):
        rgb_val = fill.fgColor.rgb
        print(f"[颜色识别] 原始RGB值: {rgb_val}")
        
        # 处理ARGB字符串（如 FFFF0000 → R=255, G=0, B=0）
        if isinstance(rgb_val, str):
            if len(rgb_val) == 8:  # AARRGGBB
                r = int(rgb_val[2:4], 16)
                g = int(rgb_val[4:6], 16)
                b = int(rgb_val[6:8], 16)
            elif len(rgb_val) == 6:  # RRGGBB
                r = int(rgb_val[0:2], 16)
                g = int(rgb_val[2:4], 16)
                b = int(rgb_val[4:6], 16)
        
        # 处理RGB元组
        elif isinstance(rgb_val, tuple):
            r, g, b = rgb_val[:3]
    
    print(f"[颜色识别] 解析后 R={r}, G={g}, B={b}")
    
    # 判定颜色（容错率极高）
    if r > 150 and g < 100 and b < 100:  # 红色系（只要R足够大）
        print("[颜色识别] → 红色")
        return 'red'
    elif g > 150 and r < 100 and b < 100:  # 绿色系（只要G足够大）
        print("[颜色识别] → 绿色")
        return 'green'
    else:
        print("[颜色识别] → 黑色")
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
        print(f"\n===== 信号: {signal_name[:30]} =====")
        
        data_cell = ws_in.cell(row_idx, col)
        cell_val = data_cell.value
        target_value, is_dashed = parse_cell_value(cell_val)
        if target_value is None:
            continue
        
        # 识别颜色
        signal_color = get_exact_color_name(data_cell)
        
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

        for t1, t2 in sig["times"]:
            ax.hlines(sig["min"], t1, t2, color='gold', lw=2)
            ax.hlines(sig["max"], t1, t2, color='gold', lw=2)
            ax.fill_between([t1, t2], sig["min"], sig["max"], color='gold', alpha=0.4)
            ax.hlines(sig["target"], t1, t2, color='blue',
                      linestyle='--' if sig["dashed"] else '-', lw=1.8)
            ax.plot((t1+t2)/2, sig["avg"], 'o', color='gold', ms=9)

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
    max_row = min(ws_in.max_row, 200)

    for r in range(2, max_row+1):
        cid = str(ws_in.cell(r, 1).value or '').strip()
        if not cid or cid in processed:
            continue
        print(f"\n==================== 处理用例: {cid} ====================")
        process_one_case(ws_in, r, wb_out)
        processed.add(cid)
    
    if 'Sheet' in wb_out.sheetnames:
        del wb_out['Sheet']
    wb_out.save(OUTPUT_FILE)
    print("\n✅ 完成！输出文件：", OUTPUT_FILE)
