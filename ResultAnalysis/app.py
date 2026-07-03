# app.py - 完整版，覆盖所有边界情况
# 运行方式: python app.py
# 依赖安装: pip install openpyxl flask flask-cors

from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
import io
import traceback

app = Flask(__name__)
CORS(app)

COLOR_MAP = {
    10: {'name': '红色', 'meaning': 'failed'},
    35: {'name': '浅蓝色', 'meaning': 'signal missing'},
    52: {'name': '橙色', 'meaning': 'trigger not occurred'},
    13: {'name': '黄色', 'meaning': 'warning'},
    11: {'name': '绿色', 'meaning': 'passed'},
    9:  {'name': 'blank', 'meaning': 'blank'}
}


def safe_int(value):
    if value is None:
        return None
    # openpyxl 对象（有 value 属性）
    if hasattr(value, 'value'):
        val = value.value
        if val is None:
            return None
        try:
            return int(val)
        except (TypeError, ValueError):
            return None
    # 普通值
    try:
        return int(value)
    except (TypeError, ValueError):
        return None


def safe_float(value):
    if value is None:
        return None
    if hasattr(value, 'value'):
        val = value.value
        if val is None:
            return None
        try:
            return float(val)
        except (TypeError, ValueError):
            return None
    try:
        return float(value)
    except (TypeError, ValueError):
        return None


def safe_str(value):
    if value is None:
        return None
    if hasattr(value, 'value'):
        val = value.value
        if val is None:
            return None
        try:
            return str(val)
        except Exception:
            return None
    try:
        return str(value)
    except Exception:
        return None


@app.route('/api/read_excel', methods=['POST'])
def read_excel():
    try:
        file = request.files.get('file')
        if not file:
            return jsonify({'success': False, 'error': '未上传文件'}), 400

        sheet_name = request.form.get('sheet_name', 'IPN')
        cell_address = request.form.get('cell_address', 'A24').upper()

        file_bytes = file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

        if sheet_name not in wb.sheetnames:
            return jsonify({
                'success': False,
                'error': f'工作表 "{sheet_name}" 不存在',
                'available_sheets': wb.sheetnames
            })

        sheet = wb[sheet_name]
        cell = sheet[cell_address]

        cell_value = cell.value

        # --- 读取颜色信息 ---
        color_indexed = None
        color_rgb = None
        color_theme = None
        color_tint = None

        fill = cell.fill

        # 安全获取 fill
        if fill is not None:
            # 检查 fgColor
            if hasattr(fill, 'fgColor'):
                fg = fill.fgColor
                if fg is not None:
                    if hasattr(fg, 'indexed'):
                        idx = fg.indexed
                        if idx is not None and idx != -1:
                            color_indexed = safe_int(idx)
                    if hasattr(fg, 'rgb'):
                        rgb_val = fg.rgb
                        if rgb_val is not None:
                            color_rgb = safe_str(rgb_val)
                    if hasattr(fg, 'theme'):
                        theme_val = fg.theme
                        if theme_val is not None:
                            color_theme = safe_int(theme_val)
                    if hasattr(fg, 'tint'):
                        tint_val = fg.tint
                        if tint_val is not None:
                            color_tint = safe_float(tint_val)

            # 如果 fgColor 没有颜色，尝试 bgColor
            if color_indexed is None and hasattr(fill, 'bgColor'):
                bg = fill.bgColor
                if bg is not None:
                    if hasattr(bg, 'indexed'):
                        idx = bg.indexed
                        if idx is not None and idx != -1:
                            color_indexed = safe_int(idx)
                    if hasattr(bg, 'rgb') and color_rgb is None:
                        rgb_val = bg.rgb
                        if rgb_val is not None:
                            color_rgb = safe_str(rgb_val)

        # 映射颜色
        color_info = None
        if color_indexed is not None:
            color_info = COLOR_MAP.get(color_indexed, None)

        return jsonify({
            'success': True,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'cell_value': str(cell_value) if cell_value is not None else None,
            'color_indexed': color_indexed,
            'color_rgb': color_rgb,
            'color_theme': color_theme,
            'color_tint': color_tint,
            'color_info': color_info,
            'all_sheets': wb.sheetnames,
            'fill_available': fill is not None,
            'has_fgColor': fill is not None and hasattr(fill, 'fgColor') and fill.fgColor is not None,
            'has_bgColor': fill is not None and hasattr(fill, 'bgColor') and fill.bgColor is not None,
        })

    except openpyxl.utils.exceptions.InvalidFileException:
        return jsonify({
            'success': False,
            'error': '无效的 Excel 文件，请确认文件未损坏且为 .xlsx 格式'
        }), 400
    except Exception as e:
        return jsonify({
            'success': False,
            'error': str(e),
            'trace': traceback.format_exc()
        }), 500


@app.route('/api/health', methods=['GET'])
def health():
    return jsonify({'status': 'ok'})


if __name__ == '__main__':
    print('🚀 启动 Excel 颜色读取服务...')
    print('📡 http://127.0.0.1:5000')
    print('📋 健康检查: http://127.0.0.1:5000/api/health')
    print('📤 上传接口: http://127.0.0.1:5000/api/read_excel')
    print('按 Ctrl+C 停止服务')
    app.run(debug=False, host='127.0.0.1', port=5000)