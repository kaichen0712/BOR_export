# app.py
# ---------------------------------------------------------
# Flask 後端 - BOR 排班系統網頁版
# 無狀態設計：不儲存使用者檔案，直接在記憶體中處理
# ---------------------------------------------------------

import io
import json
from pathlib import Path
from datetime import datetime
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename

import pandas as pd
from parser import parse_source_excel_from_bytes, load_identity_map_from_bytes, _find_sheet_names
from scheduler import expand_to_daily
from writer import write_schedule_to_excel_memory

app = Flask(__name__)
app.config['SECRET_KEY'] = 'bor-schedule-secret-key-2024'
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 最大 16MB

BASE_DIR = Path(__file__).resolve().parent

ALLOWED_EXTENSIONS = {'xlsx', 'xls', 'xlsm'}


def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS


@app.route('/')
def index():
    """首頁"""
    return render_template('index.html')


@app.route('/api/holidays', methods=['GET'])
def get_holidays():
    """取得假日資料"""
    try:
        with open(BASE_DIR / 'holidays.json', 'r', encoding='utf-8') as f:
            holidays = json.load(f)
        return jsonify({'success': True, 'holidays': holidays})
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)})


@app.route('/api/preview', methods=['POST'])
def preview_file():
    """預覽上傳的 Excel 檔案（只讀取人員名單，不儲存檔案）"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '未選擇檔案'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '未選擇檔案'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': '不支援的檔案格式，請上傳 .xlsx 或 .xlsm 檔案'})
    
    try:
        # 直接在記憶體中讀取，不儲存到磁碟
        file_bytes = file.read()
        
        # 找出頁籤名稱
        import io
        file_io = io.BytesIO(file_bytes)
        main_sheet, sub_sheet, identity_sheet = _find_sheet_names(file_io)
        
        # 讀取身分表的人員名單
        identity_map = load_identity_map_from_bytes(file_bytes, sheet_name=identity_sheet)
        staff_list = list(identity_map.keys())
        
        return jsonify({
            'success': True,
            'filename': file.filename,
            'staff_list': staff_list,
            'staff_count': len(staff_list),
            'identity_map': identity_map,
            'sheets': {
                'main': main_sheet,
                'sub': sub_sheet,
                'identity': identity_sheet
            }
        })
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'檔案處理錯誤: {str(e)}'})


@app.route('/api/generate', methods=['POST'])
def generate_schedule():
    """產生排班 Excel（不儲存檔案，直接回傳）"""
    if 'file' not in request.files:
        return jsonify({'success': False, 'error': '未選擇檔案'})
    
    file = request.files['file']
    if file.filename == '':
        return jsonify({'success': False, 'error': '未選擇檔案'})
    
    if not allowed_file(file.filename):
        return jsonify({'success': False, 'error': '不支援的檔案格式'})
    
    try:
        # 取得參數
        year = int(request.form.get('year', datetime.now().year))
        month = int(request.form.get('month', datetime.now().month))
        staff_order = request.form.get('staff_order', '')
        
        # 直接在記憶體中讀取檔案
        file_bytes = file.read()
        
        print(f"=== 網頁排班處理開始 ===")
        print(f"目標年月: {year}年{month}月")
        print(f"人員排序輸入: {staff_order[:100] if staff_order else '(無)'}...")
        
        # 找出頁籤名稱
        file_io = io.BytesIO(file_bytes)
        main_sheet, sub_sheet, identity_sheet = _find_sheet_names(file_io)
        
        # 解析來源 Excel（從記憶體）
        schedule, monthly_rules = parse_source_excel_from_bytes(file_bytes, year, month)
        identity_map = load_identity_map_from_bytes(file_bytes, sheet_name=identity_sheet)
        
        # 展開成每日排班
        daily_schedule = expand_to_daily(schedule, year, month, identity_map)
        
        # 如果有提供人員排序，按照使用者指定的順序輸出（不管有沒有排班資料）
        if staff_order.strip():
            # 處理各種格式：
            # - 「姓名 標記」如「江載仁 HN」、「陳宜汶 N2」
            # - 「姓名職稱」如「謝沛錞行助」、「林怡靜行助」
            # - 姓名有空格如「孫  華」、「羅  婕」
            import re
            raw_lines = [line.strip() for line in staff_order.strip().split('\n') if line.strip()]
            ordered_names = []
            # 職稱清單（這些單獨出現時要跳過）
            titles = ['行助', '護理師', '護士', '專師', '副護理長', '護理長', '組長', '主任', 'N', 'N1', 'N2', 'N3', 'N4', 'HN', 'AHN']
            for line in raw_lines:
                # 如果整行只是職稱或標記，跳過
                if line.strip() in titles:
                    continue
                # 移除結尾的標記（英文字母和數字的組合）
                name = re.sub(r'\s+[A-Za-z]+\d*\s*$', '', line)
                # 移除結尾的職稱
                name = re.sub(r'(行助|護理師|護士|專師|副護理長|護理長|組長|主任)$', '', name)
                # 移除姓名中的所有空白
                name = name.replace(' ', '').replace('\u3000', '').replace('\t', '')
                if name and len(name) >= 2:  # 姓名至少2個字
                    ordered_names.append(name)
            
            print(f"[DEBUG] 使用者指定排序: {ordered_names[:10]}... (共 {len(ordered_names)} 人)")
            if ordered_names:
                # 按照使用者給的順序，建立新的排班字典
                # 即使沒有排班資料，也要包含這個人（空白資料）
                ordered_schedule = {}
                for name in ordered_names:
                    if name in daily_schedule:
                        ordered_schedule[name] = daily_schedule[name]
                    else:
                        # 沒有排班資料的人，給空字典
                        ordered_schedule[name] = {}
                daily_schedule = ordered_schedule
                print(f"[DEBUG] 排序後順序: {list(daily_schedule.keys())[:10]}...")
        
        # 產生 Excel 到記憶體
        output_buffer = write_schedule_to_excel_memory(
            daily_schedule, year, month, monthly_rules, identity_map
        )
        
        # 產生檔名
        output_filename = f"BOR_{year}{month:02d}_排班表.xlsx"
        
        # 直接回傳檔案（不儲存到磁碟）
        output_buffer.seek(0)
        return send_file(
            output_buffer,
            mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            as_attachment=True,
            download_name=output_filename
        )
        
    except Exception as e:
        import traceback
        traceback.print_exc()
        return jsonify({'success': False, 'error': f'產生排班失敗: {str(e)}'})


@app.route('/api/months', methods=['GET'])
def get_available_months():
    """取得可用的月份選項"""
    current_year = datetime.now().year
    months = []
    for year in range(current_year - 1, current_year + 2):
        for month in range(1, 13):
            months.append({
                'value': f"{year}-{month:02d}",
                'label': f"{year}年{month}月",
                'year': year,
                'month': month
            })
    return jsonify({'success': True, 'months': months})


if __name__ == '__main__':
    print("=" * 50)
    print("BOR 排班系統網頁版（無狀態模式）")
    print(f"請在瀏覽器開啟: http://localhost:5000")
    print("=" * 50)
    app.run(debug=True, host='0.0.0.0', port=5000)
