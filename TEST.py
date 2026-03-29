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

# 🔥 修复正则：匹配 ==/~== 同时兼容数值（可选）
CELL_VALUE_PATTERN = re.compile(r'(~?)==\s*(-?\d+\.?\d*)?[dD]?', re.DOTALL)
NOTE_MIN_PATTERN = re.compile(r'Minimum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_MAX_PATTERN = re.compile(r'Maximum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_AVG_PATTERN = re.compile(r'Average:\s*(-?\d+\.?\d*)', re.DOTALL)
TIME_FAILED_PATTERN = re.compile(r't=\[(\d+\.\d+),(\d+\.\d+)\]|LL\+(\d+\.\d+)s', re.DOTALL)

def get_cell_fill_color_hex(cell):
    """精准获取单元格填充色（十六进制）"""
    fill = cell.fill
    if fill.patternType is None or fill.patternType == 'none':
        return '#000000'
    
    if hasattr(fill.fgColor, 'rgb') and fill.fgColor.rgb:
        rgb_val = fill.fgColor.rgb
        if isinstance(rgb_val, str):
            if len(rgb_val) == 8:
                return f'#{rgb_val[2:]}'
            elif len(rgb_val) == 6:
                return f'#{rgb_val}'
        elif isinstance(rgb_val, tuple):
            return f'#{rgb_val[0]:02x}{rgb_val[1]:02x}{rgb_val[2]:02x}'
    
    return '#000000'

def parse_cell_value(cell_text):
    """
    🔥 修复：兼容只有 ==/~== 无数值的情况
    返回：(期望值, 是否虚线)，无期望值返回 (None, False)
    """
    if not cell_text:
        return None, False
    
    s = str(cell_text).strip().replace('\n', '').replace('\r', '')
    if s in ('me', 'None', ''):
        return None, False
    
    m = CELL_VALUE_PATTERN.search(s)
    if not m:
        return None, False  # 没有 ==/~==，直接跳过
    
    # 处理分组：group1=~（可选），group2=数值（可选）
    is_dashed = bool(m.group(1))  # 有~就是虚线
    value_str = m.group(2)        # 数值部分
    
    # 兼容无数值/数值为空的情况
    if not value_str or not value_str.replace('.','').isdigit():
        target_value = 0.0  # 兜底为0
    else:
        target_value = float(value_str)
    
    return target_value, is_dashed

def parse_note_data(note_text):
    """解析批注中的Min/Max/Average/Time"""
    if not note_text:
        return 0.0, 0.0, 0.0, [(0.0, 1.0)]
    
    s = note_text.replace('\n', ' ').replace('\r', ' ').strip()
    min_val = float(m.group(1)) if (m := NOTE_MIN_PATTERN.search(s)) else 0.0
    max_val = float(m.group(1)) if (m := NOTE_MAX_PATTERN.search(s)) else 0.0
    avg_val = float(m.group(1)) if (m := NOTE_AVG_PATTERN.search(s)) else 0.0

    times = []
    # 解析 t=[7.001,7.110]
    for match in TIME_FAILED_PATTERN.findall(s):
        if match[0] and match[1]:  # t=[x,y] 格式
            times.append((round(float(match[0]), 3), round(float(match[1]), 3)))
        elif match[2]:  # LL+0.11s 格式
            times.append((0.0, round(float(match[2]), 3)))
    
    if not times:
        times = [(0.0, 1.0)]
    return min_val, max_val, avg_val, times

def process_one_case(ws_in, row_idx, wb_out):
    """处理单个测试用例"""
    test_id = ws_in.cell(row_idx, 1).value
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
        
        # 只处理有期望值（==/~==）的信号
        target_value, is_dashed = parse_cell_value(cell_val)
        if target_value is None:
            continue
        
        # 获取单元格填充色
        signal_color = get_cell_fill_color_hex(data_cell)
        
        # 解析批注
        note_text = data_cell.comment.text if (data_cell.comment and data_cell.comment.text) else ""
        min_val, max_val, avg_val, time_intervals = parse_note_data(note_text)

        signals.append({
            'name': signal_name,
            'color': signal_color,
            'target': target_value,
            'dashed': is_dashed,
            'min': min_val,
            'max': max_val,
            'avg': avg_val,
            'times': time_intervals
        })

    if not signals:
        print(f"⚠️  用例 {test_id} 无有效信号（无==/~==列）")
        return

    # 绘图
    n = len(signals)
    fig, axes = plt.subplots(n, 1, figsize=(18, 3*n), sharex=True)
    if n == 1:
        axes = [axes]

    for ax, sig in zip(axes, signals):
        ax.set_title(sig['name'], fontsize=9, loc='left', color=sig['color'], weight='bold')
        ax.grid(True, alpha=0.3)
        ax.set_ylabel("Value", fontsize=10)

        # Y轴自适应
        ys = [sig['min'], sig['max'], sig['avg'], sig['target']]
        y_min, y_max = min(ys) - 0.2, max(ys) + 0.2
        ax.set_ylim(y_min, y_max)

        # 绘制Actual/Expected/Average
        for t1, t2 in sig['times']:
            ax.hlines(sig['min'], t1, t2, color='gold', lw=2, zorder=1)
            ax.hlines(sig['max'], t1, t2, color='gold', lw=2, zorder=1)
            ax.fill_between([t1, t2], sig['min'], sig['max'], color='gold', alpha=0.4, zorder=1)
            ax.hlines(sig['target'], t1, t2, color='blue',
                      linestyle='--' if sig['dashed'] else '-', lw=1.8, zorder=2)
            ax.plot((t1+t2)/2, sig['avg'], 'o', color='gold', ms=9, zorder=3)

        # 图例
        ax.legend([f"Exp: {sig['target']}", f"Actual: {sig['min']}-{sig['max']}"],
                  loc='upper left', fontsize=7, frameon=False)
        ax.set_position([ax.get_position().x0, ax.get_position().y0, 0.85, ax.get_position().height])

    axes[-1].set_xlabel("Time (s)", fontsize=12)
    fig.suptitle(f"Test Case {test_id} - Signal Analysis", fontsize=16, y=0.98)
    
    # 保存图片
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    ws_out.add_image(Image(buf), 'A1')
    plt.close('all')

# ---------------------- 主程序 ----------------------
if __name__ == "__main__":
    try:
        wb_in = openpyxl.load_workbook(INPUT_FILE, data_only=False)
        if TARGET_SHEET_NAME not in wb_in.sheetnames:
            print(f"❌ 错误：未找到Sheet页「{TARGET_SHEET_NAME}」")
            print(f"   可用Sheet：{wb_in.sheetnames}")
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
            
            print(f"🚀 处理中: {cid}")
            process_one_case(ws_in, r, wb_out)
            processed.add(cid)

        if 'Sheet' in wb_out.sheetnames:
            del wb_out['Sheet']
        
        wb_out.save(OUTPUT_FILE)
        print(f"\n✅ 全部完成！输出文件：{OUTPUT_FILE}")
        print(f"📋 共处理 {len(processed)} 个用例：{list(processed)}")
    
    except Exception as e:
        print(f"\n❌ 运行出错：{type(e).__name__} - {e}")
        # 出错时保存已处理内容
        wb_out.save(OUTPUT_FILE)
        print(f"⚠️  已保存部分结果到：{OUTPUT_FILE}")
