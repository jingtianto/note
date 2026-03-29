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

def get_cell_fill_color_hex(cell):
    fill = cell.fill
    print(f"[颜色调试] patternType: {fill.patternType}")
    if fill.patternType is None or fill.patternType == 'none':
        print(f"[颜色调试] 无填充 → 返回黑色 #000000")
        return '#000000'

    if hasattr(fill.fgColor, 'rgb'):
        rgb_val = fill.fgColor.rgb
        print(f"[颜色调试] fgColor.rgb = {rgb_val}")
        if isinstance(rgb_val, str):
            if len(rgb_val) == 8:
                res = f'#{rgb_val[2:]}'
                print(f"[颜色调试] ARGB → {res}")
                return res
            elif len(rgb_val) == 6:
                res = f'#{rgb_val}'
                print(f"[颜色调试] RGB → {res}")
                return res
        elif isinstance(rgb_val, tuple):
            res = f'#{rgb_val[0]:02x}{rgb_val[1]:02x}{rgb_val[2]:02x}'
            print(f"[颜色调试] Tuple → {res}")
            return res

    print(f"[颜色调试] 未识别 → 返回黑色 #000000")
    return '#000000'

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
            times.append((round(float(match[0]),3), round(float(match[1]),3)))
        elif match[2]:
            times.append((0.0, round(float(match[2]),3)))
    if not times:
        times = [(0.0, 1.0)]
    return min_val, max_val, avg_val, times

def process_one_case(ws_in, row_idx, wb_out):
    test_id = ws_in.cell(row_idx,1).value
    if not test_id:
        return
    test_id = str(test_id).strip()
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

        # ==================================================
        # 打印关键信息，方便排查
        # ==================================================
        print("-" * 80)
        print(f"[信号] {signal_name}")
        print(f"[单元格内容] {cell_val}")

        target_value, is_dashed = parse_cell_value(cell_val)
        print(f"[识别期望值] {target_value}")

        if target_value is None:
            print(f"[跳过] 无期望值")
            continue

        # 读取颜色并打印
        signal_color = get_cell_fill_color_hex(data_cell)
        print(f"[最终使用颜色] {signal_color}")

        note_text = data_cell.comment.text if (data_cell.comment and data_cell.comment.text) else ""
        min_val, max_val, avg_val, times = parse_note_data(note_text)
        print(f"[实际值] min={min_val} max={max_val} avg={avg_val}")

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
        print(f"⚠️ 用例 {test_id} 无有效信号")
        return

    n = len(signals)
    fig, axes = plt.subplots(n, 1, figsize=(18, 3*n), sharex=True)
    if n == 1:
        axes = [axes]

    for ax, sig in zip(axes, signals):
        # 这里强制把颜色打出来
        print(f"\n[绘图] 信号名: {sig['name']}, 使用颜色: {sig['color']}")
        ax.set_title(sig["name"], fontsize=9, loc='left', color=sig["color"], weight='bold')
        ax.grid(True, alpha=0.3)

        ys = [sig["min"], sig["max"], sig["avg"], sig["target"]]
        y_min, y_max = min(ys)-0.5, max(ys)+0.5
        ax.set_ylim(y_min, y_max)

        for t1, t2 in sig["times"]:
            ax.hlines(sig["min"], t1, t2, color='gold', lw=2)
            ax.hlines(sig["max"], t1, t2, color='gold', lw=2)
            ax.fill_between([t1,t2], sig["min"], sig["max"], color='gold', alpha=0.4)
            ax.hlines(sig["target"], t1, t2, color='blue',
                      linestyle='--' if sig["dashed"] else '-', lw=1.8)
            ax.plot((t1+t2)/2, sig["avg"], 'o', color='gold', ms=9)

        ax.legend([f"Exp:{sig['target']}", f"Actual:{sig['min']}~{sig['max']}"],
                  loc='upper left', fontsize=7)

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
        print(f"不存在sheet: {TARGET_SHEET_NAME}")
        exit(1)

    ws_in = wb_in[TARGET_SHEET_NAME]
    wb_out = openpyxl.Workbook()
    processed = set()
    max_row = min(ws_in.max_row, 200)

    for r in range(2, max_row+1):
        cid = ws_in.cell(r, 1).value
        if not cid:
            continue
        cid = str(cid).strip()
        if cid in processed:
            continue
        print("\n" + "="*80)
        print(f"🚀 处理用例: {cid}")
        process_one_case(ws_in, r, wb_out)
        processed.add(cid)

    if 'Sheet' in wb_out.sheetnames:
        del wb_out['Sheet']
    wb_out.save(OUTPUT_FILE)
    print("\n✅ 全部完成")
