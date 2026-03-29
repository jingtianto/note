import openpyxl

# ---------------------- 配置区 ----------------------
INPUT_FILE = r"D:\PythonProject\自动分析测试结果\1.xlsx"
TARGET_SHEET_NAME = "IPN"
# 要测试的单元格：AR7 和 AS7
TEST_CELLS = ['AR7', 'AS7']

def check_cell_color(cell):
    """打印单元格所有颜色相关信息"""
    print(f"\n=== 单元格 {cell.coordinate} ===")
    fill = cell.fill
    
    # 1. 打印填充类型
    print(f"1. 填充类型 (patternType): {fill.patternType}")
    
    # 2. 打印前景色所有属性
    fg_color = fill.fgColor
    print(f"2. 前景色对象: {fg_color}")
    print(f"   - type: {type(fg_color)}")
    print(f"   - rgb: {getattr(fg_color, 'rgb', '无此属性')}")
    print(f"   - theme: {getattr(fg_color, 'themeColor', '无此属性')}")
    print(f"   - tint: {getattr(fg_color, 'tint', '无此属性')}")
    print(f"   - indexed: {getattr(fg_color, 'indexed', '无此属性')}")
    
    # 3. 尝试解析RGB（兼容所有格式）
    try:
        rgb_val = fg_color.rgb
        if isinstance(rgb_val, str):
            if len(rgb_val) == 8:  # AARRGGBB
                r = int(rgb_val[2:4], 16)
                g = int(rgb_val[4:6], 16)
                b = int(rgb_val[6:8], 16)
                print(f"4. 解析RGB: R={r}, G={g}, B={b}")
            elif len(rgb_val) == 6:  # RRGGBB
                r = int(rgb_val[0:2], 16)
                g = int(rgb_val[2:4], 16)
                b = int(rgb_val[4:6], 16)
                print(f"4. 解析RGB: R={r}, G={g}, B={b}")
        elif isinstance(rgb_val, tuple):
            r, g, b = rgb_val[:3]
            print(f"4. 解析RGB: R={r}, G={g}, B={b}")
    except Exception as e:
        print(f"4. 解析RGB失败: {e}")

# ---------------------- 主程序 ----------------------
if __name__ == "__main__":
    try:
        # 打开Excel（只读模式，避免修改）
        wb = openpyxl.load_workbook(INPUT_FILE, data_only=False, read_only=False)
        
        # 检查工作表是否存在
        if TARGET_SHEET_NAME not in wb.sheetnames:
            print(f"❌ 找不到工作表: {TARGET_SHEET_NAME}")
            print(f"   可用工作表: {wb.sheetnames}")
            wb.close()
            exit(1)
        
        ws = wb[TARGET_SHEET_NAME]
        
        # 逐个检查指定单元格
        for cell_addr in TEST_CELLS:
            try:
                cell = ws[cell_addr]
                check_cell_color(cell)
            except Exception as e:
                print(f"❌ 读取单元格 {cell_addr} 失败: {e}")
        
        wb.close()
        print("\n✅ 检测完成！")
    
    except Exception as e:
        print(f"\n❌ 程序运行失败: {type(e).__name__} - {e}")
