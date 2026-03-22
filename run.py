#!/usr/bin/env python3
"""
動支及黏存單自動產生系統 - 本地伺服器
執行方式：python run.py
然後開啟瀏覽器至 http://localhost:8000
"""

import os
import sys
import json
import tempfile
import threading
import webbrowser
from http.server import HTTPServer, SimpleHTTPRequestHandler
from socketserver import ThreadingMixIn
from urllib.parse import urlparse, parse_qs

# 確保工作目錄正確
os.chdir(os.path.dirname(os.path.abspath(__file__)))

try:
    from fill_excel import process as fill_excel
    PYTHON_MODE = True
except ImportError:
    PYTHON_MODE = False
    print('警告：找不到 fill_excel.py，將使用前端模式（格式可能不完整）')


class ThreadedHTTPServer(ThreadingMixIn, HTTPServer):
    """多執行緒 HTTP 伺服器，避免單一請求阻塞"""
    daemon_threads = True


class BudgetHandler(SimpleHTTPRequestHandler):
    def log_message(self, format, *args):
        # 只顯示 POST 請求
        if args and str(args[1]) == '200' and self.command == 'POST':
            print(f'[{self.command}] {self.path}')

    def do_OPTIONS(self):
        self.send_response(200)
        self._send_cors()
        self.send_header('Content-Length', '0')
        self.end_headers()

    def _send_cors(self):
        self.send_header('Access-Control-Allow-Origin', '*')
        self.send_header('Access-Control-Allow-Methods', 'GET, POST, OPTIONS')
        self.send_header('Access-Control-Allow-Headers', 'Content-Type')
        self.send_header('Connection', 'keep-alive')

    def do_GET(self):
        if self.path == '/api/ping':
            self.send_response(200)
            self._send_cors()
            self.send_header('Content-Type', 'application/json')
            body = b'{"ok":true,"mode":"python"}'
            self.send_header('Content-Length', str(len(body)))
            self.end_headers()
            self.wfile.write(body)
        else:
            # Serve static files for everything else
            super().do_GET()

    def do_POST(self):
        if self.path == '/api/generate':
            self.handle_generate()
        else:
            self.send_error(404)

    def handle_generate(self):
        content_length = int(self.headers['Content-Length'])
        body = self.rfile.read(content_length)

        try:
            data = json.loads(body.decode('utf-8'))
        except json.JSONDecodeError as e:
            self.send_error(400, f'JSON 格式錯誤: {e}')
            return

        if not PYTHON_MODE:
            self.send_error(503, '未安裝 openpyxl，請執行 pip install openpyxl')
            return

        try:
            # 產生輸出檔案
            with tempfile.NamedTemporaryFile(
                suffix='.xlsx', delete=False, prefix='動支單_'
            ) as tmp:
                output_path = tmp.name

            template_path = os.path.join(
                os.path.dirname(os.path.abspath(__file__)),
                'template', 'template.xlsx'
            )

            fill_excel(template_path, data, output_path)

            # 讀取並回傳檔案
            with open(output_path, 'rb') as f:
                xlsx_data = f.read()

            os.unlink(output_path)

            # 組合檔名（RFC 5987 UTF-8 編碼）
            from urllib.parse import quote
            sheet_type = data.get('templateType', 'budget')
            year = data.get('year', '')
            month = str(data.get('month', '')).zfill(2)
            day = str(data.get('day', '')).zfill(2)
            type_ascii = 'budget-in' if sheet_type == '預算內' else 'agency'
            filename_zh = f'動支及黏存單_{sheet_type}_{year}{month}{day}.xlsx'
            filename_safe = f'budget_{type_ascii}_{year}{month}{day}.xlsx'
            filename_encoded = quote(filename_zh, safe='')

            self.send_response(200)
            self.send_header('Content-Type',
                'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
            self.send_header('Content-Disposition',
                f"attachment; filename=\"{filename_safe}\"; filename*=UTF-8''{filename_encoded}")
            self.send_header('Content-Length', str(len(xlsx_data)))
            self._send_cors()
            self.end_headers()
            self.wfile.write(xlsx_data)

        except Exception as e:
            import traceback
            traceback.print_exc()
            self.send_error(500, str(e))


def check_openpyxl():
    try:
        import openpyxl
        return True
    except ImportError:
        return False


def main():
    port = 8000
    # 避免與已有伺服器衝突
    for p in [8000, 8001, 8002, 8080, 8888]:
        try:
            server = ThreadedHTTPServer(('', p), BudgetHandler)
            port = p
            break
        except OSError:
            continue
    else:
        print('Error: Cannot find an available port. Please free up 8000-8888.')
        sys.exit(1)

    if not check_openpyxl():
        print('Installing openpyxl...')
        os.system(f'{sys.executable} -m pip install openpyxl -q')

    url = f'http://localhost:{port}'
    print(f'[OK] Server started: {url}')
    print(f'     Press Ctrl+C to stop.')

    # 自動開啟瀏覽器
    threading.Timer(1.0, lambda: webbrowser.open(url)).start()

    try:
        server.serve_forever()
    except KeyboardInterrupt:
        print('\nServer stopped.')


if __name__ == '__main__':
    main()
