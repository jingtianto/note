import openpyxl
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import re
import io
from openpyxl.styles import PatternFill

# ---------------------- 配置区 ----------------------
INPUT_FILE = r"D:\PythonProject\自动分析测试结果\1.xlsx"
OUTPUT_FILE = r"D:\PythonProject\自动分析测试结果\2.xlsx"
TARGET_SHEET_NAME = "IPN"

# 正则表达式（适配你的批注格式）
CELL_VALUE_PATTERN = re.compile(r'(~?)==?\s*(-?\d+\.?\d*)[dD]', re.DOTALL)
NOTE_MIN_PATTERN = re.compile(r'Minimum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_MAX_PATTERN = re.compile(r'Maximum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_AVG_PATTERN = re.compile(r'Average:\s*(-?\d+\.?\d*)', re.DOTALL)
TIME_FAILED_PATTERN = re.compile(r't=\[(\d+\.\d+),(\d+\.\d+)\]', re.DOTALL)  # 适配 t=[7.001,7.110]

def get_cell_fill_color(cell):
    fill = cell.fill
    if fill.patternType is None or fill.patternType == 'none':
        return '#000000'
    if hasattr(fill.fgColor, 'rgb') and fill.fgColor.rgb:
        rgb_str = fill.fgColor.rgb
        if isinstance(rgb_str, str) and len(rgb_str) == 8:
            return f'#{rgb_str[2:]}'
        elif isinstance(rgb_str, str) and len(rgb_str) == 6:
            return f'#{rgb_str}'
    return '#000000'

def parse_cell_value(cell_text):
    if not cell_text:
        return None, False
    cell_text_str = str(cell_text).replace('\n','').replace('\r','').strip()
    if cell_text_str in ('me', 'None', ''):
        return None, False
    match = CELL_VALUE_PATTERN.search(cell_text_str)
    if match:
        is_dashed = bool(match.group(1))
        value = float(match.group(2))
        return value, is_dashed
    return None, False

def parse_note_data(note_text):
    if not note_text:
        return 0.0, 0.0, 0.0, [(0.0, 1.0)]  # 无批注时给默认时间
    note_text = note_text.replace('\n',' ').replace('\r',' ').strip()
    min_match = NOTE_MIN_PATTERN.search(note_text)
    min_val = float(min_match.group(1)) if min_match else 0.0
    max_match = NOTE_MAX_PATTERN.search(note_text)
    max_val = float(max_match.group(1)) if max_match else 0.0
    avg_match = NOTE_AVG_PATTERN.search(note_text)
    avg_val = float(avg_match.group(1)) if avg_match else 0.0
    time_matches = TIME_FAILED_PATTERN.findall(note_text)
    failed_intervals = []
    if time_matches:
        for t_start_str, t_end_str in time_matches:
            t_start = round(float(t_start_str), 3)
            t_end = round(float(t_end_str), 3)
            failed_intervals.append((t_start, t_end))
    else:
        failed_intervals = [(0.0, 1.0)]  # 兜底时间区间
    return min_val, max_val, avg_val, failed_intervals

def process_test_case(ws_in, row_idx, wb_out):
    test_id = ws_in.cell(row_idx, 1).value 
    if not test_id:
        return
    test_id_str = str(test_id).strip()
    sheet_name = f"TC_{test_id_str}"
    if sheet_name in wb_out.sheetnames:
        del wb_out[sheet_name]
    ws_out = wb_out.create_sheet(title=sheet_name)
    
    valid_signals = []
    for col_idx in range(3, ws_in.max_column + 1):
        signal_name = ws_in.cell(1, col_idx).value or f"Signal_Col{col_idx}"
        data_cell = ws_in.cell(row_idx, col_idx)
        signal_color = get_cell_fill_color(data_cell)
        cell_val = data_cell.value
        note_text = data_cell.comment.text if (data_cell.comment and data_cell.comment.text) else ""
        
        # 修复：保留0值，只过滤真正无效值
        cell_val_str = str(cell_val).strip() if cell_val is not None else ""
        if cell_val is None or cell_val_str in ('me', 'None', ''):
            continue
        
        target_value, is_dashed = parse_cell_value(cell_val)
        min_val, max_val, avg_val, failed_intervals = parse_note_data(note_text)
        
        # 修复：即使target_value是0，也保留
        if target_value is None and not failed_intervals:
            continue
        if target_value is None:
            target_value = 0.0  # 兜底为0
        
        valid_signals.append({
            "name": signal_name,
            "color": signal_color,
            "target": target_value,
            "is_dashed": is_dashed,
            "min": min_val,
            "max": max_val,
            "avg": avg_val,
            "times": failed_intervals
        })
    
    if not valid_signals:
        return
    
    n = len(valid_signals)
    fig, axes = plt.subplots(n, 1, figsize=(22, 4*n), sharex=True)
    if n == 1:
        axes = [axes]
    for i, sig in enumerate(valid_signals):
        ax = axes[i]
        ax.set_title(sig["name"], fontweight='bold', fontsize=9, loc='left', pad=10, color=sig["color"])
        ax.grid(True, alpha=0.3)
        ax.set_ylabel("Value", fontsize=10)
        
        all_y_values = [sig["min"], sig["max"], sig["avg"], sig["target"]]
        y_min = min(all_y_values) - 0.5
        y_max = max(all_y_values) + 0.5
        ax.set_ylim(y_min, y_max)
        
        # 绘制Actual区间（0值也会画）
        for t_start, t_end in sig["times"]:
            ax.hlines(sig["min"], t_start, t_end, color='gold', linewidth=3, zorder=1)
            ax.hlines(sig["max"], t_start, t_end, color='gold', linewidth=3, zorder=1)
            ax.fill_between([t_start, t_end], sig["min"], sig["max"], color='gold', alpha=0.5, zorder=1)
        
        # 绘制Expected线（0值也会画）
        for t_start, t_end in sig["times"]:
            ax.hlines(sig["target"], t_start, t_end, color='blue', 
                      linestyle='--' if sig["is_dashed"] else '-', linewidth=2, zorder=2)
        
        # 绘制Average点（0值也会画）
        for t_start, t_end in sig["times"]:
            ax.plot((t_start + t_end)/2, sig["avg"], 'yo', markersize=12, zorder=3)
        
        ax.legend(["Actual Min Value", "Actual Max Value", "Actual Range", 
                   f"Expected: {sig['target']}", f"Actual Average: {sig['avg']}"],
                  loc='upper left', bbox_to_anchor=(1.01, 1.0), fontsize=8)
        ax.set_position([ax.get_position().x0, ax.get_position().y0, 
                        ax.get_position().width*0.85, ax.get_position().height])
    axes[-1].set_xlabel("Time (s)", fontsize=12)
    fig.suptitle(f"Test Case {test_id_str} - Signal Analysis", fontsize=16, y=0.98)
    plt.tight_layout(rect=[0, 0, 0.88, 0.96])
    buf = io.BytesIO()
    plt.savefig(buf, format='png', bbox_inches='tight', dpi=300)
    buf.seek(0)
    ws_out.add_image(Image(buf), "A1")
    plt.close()

if __name__ == "__main__":
    wb_in = openpyxl.load_workbook(INPUT_FILE, data_only=False)
    if TARGET_SHEET_NAME not in wb_in.sheetnames:
        print(f"❌ 未找到Sheet: {TARGET_SHEET_NAME}")
        exit(1)
    ws_in = wb_in[TARGET_SHEET_NAME]
    wb_out = openpyxl.Workbook()
    processed_ids = set()
    for row_idx in range(2, ws_in.max_row + 1):
        current_id = ws_in.cell(row_idx, 1).value
        if not current_id:
            continue
        current_id_str = str(current_id).strip()
        if current_id_str in processed_ids:
            continue
        process_test_case(ws_in, row_idx, wb_out)
        processed_ids.add(current_id_str)
    if "Sheet" in wb_out.sheetnames:
        del wb_out["Sheet"]
    wb_out.save(OUTPUT_FILE)
    print(f"✅ 处理完成！生成Sheet: {list(wb_out.sheetnames)}")
