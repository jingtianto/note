# app.py - 完全修复 JSON 序列化问题
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

        # --- 读取颜色信息（全部转为 Python 原生类型） ---
        color_indexed = None
        color_rgb = None
        color_theme = None
        color_tint = None

        fill = cell.fill
        if fill and hasattr(fill, 'fgColor') and fill.fgColor:
            fg = fill.fgColor
            # indexed
            if hasattr(fg, 'indexed'):
                idx = fg.indexed
                if idx is not None and idx != -1:
                    color_indexed = int(idx)  # 转为 int
            # rgb
            if hasattr(fg, 'rgb'):
                rgb_val = fg.rgb
                if rgb_val is not None:
                    color_rgb = str(rgb_val)  # 转为 str
            # theme
            if hasattr(fg, 'theme'):
                theme_val = fg.theme
                if theme_val is not None:
                    color_theme = int(theme_val)  # 转为 int
            # tint
            if hasattr(fg, 'tint'):
                tint_val = fg.tint
                if tint_val is not None:
                    color_tint = float(tint_val)  # 转为 float

        # 如果 fgColor 没有颜色，尝试 bgColor
        if color_indexed is None and fill and hasattr(fill, 'bgColor') and fill.bgColor:
            bg = fill.bgColor
            if hasattr(bg, 'indexed'):
                idx = bg.indexed
                if idx is not None and idx != -1:
                    color_indexed = int(idx)
            if hasattr(bg, 'rgb') and not color_rgb:
                rgb_val = bg.rgb
                if rgb_val is not None:
                    color_rgb = str(rgb_val)

        # 映射颜色
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
            'has_fgColor': fill and hasattr(fill, 'fgColor') and fill.fgColor is not None,
            'has_bgColor': fill and hasattr(fill, 'bgColor') and fill.bgColor is not None,
        })

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
    print('🚀 启动服务...')
    print('📡 http://127.0.0.1:5000')
    app.run(debug=False, host='127.0.0.1', port=5000)