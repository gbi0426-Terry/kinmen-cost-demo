import sys, os
sys.path.insert(0, os.path.join(os.path.dirname(__file__), '..', 'engine'))

from flask import Flask, render_template, request, send_file, jsonify
import shutil, tempfile
from calculate import calculate

app = Flask(__name__)

DATA_DIR  = os.path.join(os.path.dirname(__file__), '..', 'data')
OUTPUT_DIR = os.path.join(os.path.dirname(__file__), '..', 'output')

EXPECTED_FILES = {
    '01_薪資統計表.xlsx':   '薪資統計表',
    '02_工時比例表.xlsx':   '工時比例表',
    '03_半成品盤存表.xlsx': '半成品盤存表',
    '04_包裝廠生產月報.xlsx': '包裝廠生產月報',
    '05_物料盤存表.xlsx':   '物料盤存表',
}

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/run-demo', methods=['POST'])
def run_demo():
    try:
        result = calculate(DATA_DIR, OUTPUT_DIR)
        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500

@app.route('/upload', methods=['POST'])
def upload():
    tmpdir = tempfile.mkdtemp()
    try:
        for key in EXPECTED_FILES:
            f = request.files.get(key)
            dest = os.path.join(tmpdir, key)
            if f:
                f.save(dest)
            else:
                # 缺少的檔案用虛擬資料補上
                shutil.copy(os.path.join(DATA_DIR, key), dest)

        out_dir = tempfile.mkdtemp()
        result = calculate(tmpdir, out_dir)

        # 複製輸出到固定目錄供下載
        for fname in os.listdir(out_dir):
            shutil.copy(os.path.join(out_dir, fname), os.path.join(OUTPUT_DIR, fname))
        shutil.rmtree(out_dir, ignore_errors=True)

        return jsonify(result)
    except Exception as e:
        return jsonify({'success': False, 'error': str(e)}), 500
    finally:
        shutil.rmtree(tmpdir, ignore_errors=True)

@app.route('/download')
def download():
    month = '2026年05月'
    path = os.path.join(OUTPUT_DIR, f'產品成本表_{month}.xlsx')
    if not os.path.exists(path):
        return '檔案不存在，請先執行計算', 404
    return send_file(path, as_attachment=True,
                     download_name=f'產品成本表_{month}.xlsx')

if __name__ == '__main__':
    app.run(debug=False, port=5001)
