# app.py - Excel 颜色读取后端
# 运行方式: python app.py
# 依赖安装: pip install openpyxl flask flask-cors

from flask import Flask, request, jsonify
from flask_cors import CORS
import openpyxl
import io
import traceback

app = Flask(__name__)
CORS(app)  # 允许前端跨域访问

# 颜色映射表（与前端一致）
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
    """接收 Excel 文件，读取指定单元格的颜色信息"""
    try:
        # 获取上传的文件
        file = request.files.get('file')
        if not file:
            return jsonify({'success': False, 'error': '未上传文件'}), 400

        # 获取参数
        sheet_name = request.form.get('sheet_name', 'IPN')
        cell_address = request.form.get('cell_address', 'A24').upper()

        # 用 openpyxl 加载文件
        file_bytes = file.read()
        wb = openpyxl.load_workbook(io.BytesIO(file_bytes), data_only=True)

        # 检查工作表是否存在
        if sheet_name not in wb.sheetnames:
            return jsonify({
                'success': False,
                'error': f'工作表 "{sheet_name}" 不存在',
                'available_sheets': wb.sheetnames
            })

        sheet = wb[sheet_name]

        # 检查单元格是否存在
        if cell_address not in sheet:
            return jsonify({
                'success': False,
                'error': f'单元格 "{cell_address}" 不存在'
            })

        cell = sheet[cell_address]
        cell_value = cell.value

        # 读取颜色索引
        color_index = None
        fill = cell.fill
        if fill and hasattr(fill, 'fgColor') and fill.fgColor:
            color_index = fill.fgColor.indexed

        # 如果 indexed 为空，尝试读取 rgb
        rgb_value = None
        if fill and hasattr(fill, 'fgColor') and fill.fgColor:
            if hasattr(fill.fgColor, 'rgb'):
                rgb_value = fill.fgColor.rgb

        # 映射颜色
        color_info = COLOR_MAP.get(color_index, None) if color_index is not None else None

        # 获取所有工作表名称（用于调试）
        all_sheets = wb.sheetnames

        return jsonify({
            'success': True,
            'sheet_name': sheet_name,
            'cell_address': cell_address,
            'cell_value': str(cell_value) if cell_value is not None else None,
            'color_indexed': color_index,
            'color_rgb': rgb_value,
            'color_info': color_info,
            'all_sheets': all_sheets
        })

    except Exception as e:
        error_trace = traceback.format_exc()
        return jsonify({
            'success': False,
            'error': str(e),
            'trace': error_trace
        }), 500


@app.route('/api/health', methods=['GET'])
def health():
    """健康检查接口"""
    return jsonify({'status': 'ok', 'service': 'Excel 颜色读取服务'})


if __name__ == '__main__':
    print('🚀 启动 Excel 颜色读取服务...')
    print('📡 访问地址: http://127.0.0.1:5000')
    print('📋 健康检查: http://127.0.0.1:5000/api/health')
    print('📤 上传接口: http://127.0.0.1:5000/api/read_excel')
    print('按 Ctrl+C 停止服务')
    app.run(debug=False, host='127.0.0.1', port=5000)