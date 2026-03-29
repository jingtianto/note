import openpyxl
from openpyxl.drawing.image import Image
import matplotlib.pyplot as plt
import re
import io
import warnings
warnings.filterwarnings('ignore')

# ---------------------- 配置区 ----------------------
INPUT_FILE = r"D:\PythonProject\自动分析测试结果\1.xlsx"
OUTPUT_FILE = r"D:\PythonProject\自动分析测试结果\2.xlsx"  # 输出文件路径
TARGET_SHEET_NAME = "IPN"

# 正则表达式
CELL_VALUE_PATTERN = re.compile(r'(~?)==?\s*(-?\d+\.?\d*)[dD]', re.DOTALL)
NOTE_MIN_PATTERN = re.compile(r'Minimum value:\s*(-?\d+\.?\d*)', re.DOTALL)
NOTE_MAX_PATTERN = re.compile(r'Maximum value:\s*(-?\d+\.?\d*)', re.DOTALL)  # 修复：之前写错成Minimum
NOTE_AVG_PATTERN = re.compile(r'Average:\s*(-?\d+\.?\d*)', re.DOTALL)
TIME_FAILED_PATTERN = re.compile(r't=\[(\d+\.\d+),(\d+\.\d+)\]', re.DOTALL)

def get_cell_fill_color(cell):
    """获取单元格填充色（处理Excel ARGB格式）"""
    fill = cell.fill
    if fill.patternType is None or fill.patternType == 'none':
        return '#000000'
    if hasattr(fill.fgColor, 'rgb') and fill.fgColor.rgb:
        rgb_str = fill.fgColor.rgb
        if isinstance(rgb_str, str) and len(rgb_str) == 8:
            return f'#{rgb_str[2:]}'
    return '#000000'

def parse_cell_value(cell_text):
    """解析单元格期望值"""
    if not cell_text:
        return None, False
    s = str(cell_text).strip().replace('\n', '').replace('\r', '')
    if s in ('me', 'None', ''):
        return None, False
    m = CELL_VALUE_PATTERN.search(s)
    if m:
        return float(m.group(2)), bool(m.group(1))
    return None, False

def parse_note_data(note_text):
    """解析批注中的min/max/avg/时间"""
    if not note_text:
        return 0.0, 0.0, 0.0, [(0.0, 1.0)]
    s = note_text.replace('\n', ' ').replace('\r', ' ').strip()
    min_val = float(m.group(1)) if (m := NOTE_MIN_PATTERN.search(s)) else 0.0
    max_val = float(m.group(1)) if (m := NOTE_MAX_PATTERN.search(s)) else 0.0
    avg_val = float(m.group(1)) if (m := NOTE_AVG_PATTERN.search(s)) else 0.0

    times = []
    for t1, t2 in TIME_FAILED_PATTERN.findall(s):
        times.append((round(float(t1),3), round(float(t2),3)))
    if not times:
        times = [(0.0, 1.0)]
    return min_val, max_val, avg_val, times

def process_one_case(ws_in, row_idx, wb_out):
    """处理单个测试用例"""
    test_id = ws_in.cell(row_idx,1).value
    if not test_id:
        return
    test_id = str(test_id).strip()
    sheet_name = f"TC_{test_id}"

    # 删除已存在的Sheet（避免重复）
    if sheet_name in wb_out.sheetnames:
        del wb_out[sheet_name]
    ws_out = wb_out.create_sheet(sheet_name)

    signals = []
    # 限制最大列数，避免卡死
    max_col = min(ws_in.max_column, 1000)

    for col in range(3, max_col+1):
        header = ws_in.cell(1, col).value
        if not header:  # 跳过空表头列
            continue
        cell = ws_in.cell(row_idx, col)
        color = get_cell_fill_color(cell)
        val = cell.value
        note = cell.comment.text if (cell.comment and cell.comment.text) else ""

        if val is None:
            continue
        s_val = str(val).strip()
        if s_val in ('me', 'None', ''):
            continue

        target, dashed = parse_cell_value(val)
        mn, mx, avg, ts = parse_note_data(note)

        if target is None:
            target = 0.0

        signals.append({
            'name': header,
            'color': color,
            'target': target,
            'dashed': dashed,
            'min': mn,
            'max': mx,
            'avg': avg,
            'times': ts
        })

    if not signals:
        print(f"⚠️  测试用例{test_id}无有效信号")
        return

    # 绘图（简化版，更快）
    n = len(signals)
    fig, axes = plt.subplots(n, 1, figsize=(16, 3*n), sharex=True)
    if n == 1:
        axes = [axes]

    for ax, sig in zip(axes, signals):
        ax.set_title(sig['name'], fontsize=9, loc='left', color=sig['color'], weight='bold')
        ax.grid(True, alpha=0.3)

        # Y轴自适应
        ys = [sig['min'], sig['max'], sig['avg'], sig['target']]
        y0, y1 = min(ys)-0.2, max(ys)+0.2
        ax.set_ylim(y0, y1)

        # 绘制Actual/Expected/Average
        for t1, t2 in sig['times']:
            ax.hlines(sig['min'], t1, t2, color='gold', lw=2, label='Actual Min')
            ax.hlines(sig['max'], t1, t2, color='gold', lw=2, label='Actual Max')
            ax.fill_between([t1,t2], sig['min'], sig['max'], color='gold', alpha=0.4, label='Actual Range')
            ax.hlines(sig['target'], t1, t2, color='blue',
                      linestyle='--' if sig['dashed'] else '-', lw=1.5, label=f'Expected: {sig["target"]}')
            ax.plot((t1+t2)/2, sig['avg'], 'o', c='gold', ms=8, label=f'Average: {sig["avg"]}')

        # 去重图例
        handles, labels = ax.get_legend_handles_labels()
        by_label = dict(zip(labels, handles))
        ax.legend(by_label.values(), by_label.keys(), loc='upper left', fontsize=7)

    plt.tight_layout(pad=1.0)
    # 保存图片（低dpi加快速度）
    buf = io.BytesIO()
    plt.savefig(buf, format='png', dpi=150, bbox_inches='tight')
    buf.seek(0)
    ws_out.add_image(Image(buf), 'A1')
    plt.close()  # 及时关闭画布，释放内存

# ---------------------- 主程序 ----------------------
if __name__ == "__main__":
    try:
        # 加载Excel（保留格式）
        wb_in = openpyxl.load_workbook(INPUT_FILE, data_only=False)
        if TARGET_SHEET_NAME not in wb_in.sheetnames:
            print(f"❌ 错误：未找到Sheet页「{TARGET_SHEET_NAME}」")
            print(f"   可用Sheet：{wb_in.sheetnames}")
            exit(1)
        ws_in = wb_in[TARGET_SHEET_NAME]
        wb_out = openpyxl.Workbook()

        done = set()
        # 限制最大行数，避免卡死
        max_row = min(ws_in.max_row, 200)

        # 遍历测试用例
        for r in range(2, max_row+1):
            cid = ws_in.cell(r,1).value
            if not cid:
                continue
            cid = str(cid).strip()
            if cid in done:
                continue
            print(f"✅ 正在处理用例：{cid}")
            process_one_case(ws_in, r, wb_out)
            done.add(cid)

        # 清理默认Sheet
        if 'Sheet' in wb_out.sheetnames:
            del wb_out['Sheet']
        
        # 🔥 修复：使用正确的变量名 OUTPUT_FILE
        wb_out.save(OUTPUT_FILE)
        print(f"\n🎉 处理完成！文件已保存到：{OUTPUT_FILE}")
        print(f"📋 共处理 {len(done)} 个测试用例：{list(done)}")
        
    except Exception as e:
        print(f"\n❌ 运行出错：{type(e).__name__} - {e}")
        # 出错时也保存已处理的内容
        wb_out.save(OUTPUT_FILE)
        print(f"⚠️  已保存部分结果到：{OUTPUT_FILE}")
