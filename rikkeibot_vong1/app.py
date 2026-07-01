import os
import sys
import io
import json
import threading
import queue
from flask import Flask, request, render_template, Response, jsonify, send_from_directory

app = Flask(__name__)

# Configure stdout/stderr encoding to UTF-8 to prevent CP1252/Windows encoding crashes
for stream in [sys.stdout, sys.stderr]:
    if hasattr(stream, 'reconfigure'):
        try:
            stream.reconfigure(encoding='utf-8')
        except Exception:
            pass

# Queue for real-time logs
log_queue = queue.Queue()

class StreamCapture:
    def __init__(self, original_stdout):
        self.original_stdout = original_stdout

    def write(self, s):
        if s:
            if isinstance(s, bytes):
                try:
                    s = s.decode('utf-8', errors='replace')
                except Exception:
                    s = str(s)
            try:
                self.original_stdout.write(s)
            except TypeError:
                try:
                    if isinstance(s, str):
                        self.original_stdout.write(s.encode('utf-8', errors='replace'))
                    else:
                        self.original_stdout.write(str(s))
                except Exception:
                    pass
            except UnicodeEncodeError:
                try:
                    enc = getattr(self.original_stdout, 'encoding', None) or 'utf-8'
                    self.original_stdout.write(s.encode(enc, errors='replace').decode(enc, errors='replace'))
                except Exception:
                    pass
            except Exception:
                pass

            try:
                self.original_stdout.flush()
            except Exception:
                pass

            if s.strip():
                log_queue.put(s)

    def flush(self):
        try:
            self.original_stdout.flush()
        except Exception:
            pass

    def __getattr__(self, attr):
        return getattr(self.original_stdout, attr)

# Replace global stdout to capture all prints
original_stdout = sys.stdout
sys.stdout = StreamCapture(original_stdout)

def run_script(mode):
    try:
        # Check if running as packaged PyInstaller EXE
        if getattr(sys, 'frozen', False):
            log_queue.put(">>> Khởi chạy Pipeline trong luồng xử lý (Chế độ đóng gói EXE)...\n")
            if mode == 'overwrite':
                import full_automation_lark_app
                import importlib
                importlib.reload(full_automation_lark_app)
                full_automation_lark_app.run_automation()
            elif mode == 'update':
                import auto_update_new_app
                import importlib
                importlib.reload(auto_update_new_app)
                auto_update_new_app.run_automation()
            else:
                log_queue.put("LỖI: Chế độ chạy không hợp lệ.\n")
            log_queue.put("===DONE===")
        else:
            log_queue.put(">>> Khởi chạy Pipeline trong tiến trình cô lập để tối ưu RAM...\n")
            script_name = 'full_automation_lark_app.py' if mode == 'overwrite' else 'auto_update_new_app.py'
            
            # Đảm bảo đường dẫn chạy chính xác cho script con
            base_dir = os.path.dirname(os.path.abspath(__file__))
            script_path = os.path.join(base_dir, script_name)
            
            import subprocess
            process = subprocess.Popen(
                [sys.executable, script_path],
                stdout=subprocess.PIPE,
                stderr=subprocess.STDOUT,
                text=True,
                bufsize=1, # Line-buffered
                encoding='utf-8',
                errors='replace',
                cwd=base_dir
            )
            
            while True:
                line = process.stdout.readline()
                if not line and process.poll() is not None:
                    break
                if line:
                    log_queue.put(line)
                    
            process.stdout.close()
            return_code = process.wait()
            
            if return_code == 0:
                log_queue.put("===DONE===")
            else:
                log_queue.put(f"\n[LỖI HỆ THỐNG]: Tiến trình con kết thúc đột ngột với mã lỗi {return_code}\n")
                log_queue.put("===DONE===")
    except Exception as e:
        log_queue.put(f"\n[LỖI HỆ THỐNG]: {e}\n")
        import traceback
        log_queue.put(traceback.format_exc())
        log_queue.put("===DONE===")

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/api/run', methods=['POST'])
def run():
    data = request.json
    config_data = data.get('config', {})
    mode = data.get('mode', 'overwrite')
    
    # Save API config
    config_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'config.json')
    with open(config_path, 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=4)
        
    # Clear old results files (JSON & ZIP)
    base_dir = os.path.dirname(os.path.abspath(__file__))
    for f_name in ['processed_students.json', 'PDF_TrungTuyen_CNTT.zip', 'PDF_TrungTuyen_QTKD.zip']:
        f_path = os.path.join(base_dir, f_name)
        if os.path.exists(f_path):
            try:
                os.remove(f_path)
            except Exception:
                pass
    
    # Clear old queue
    while not log_queue.empty():
        log_queue.get()
        
    # Start thread
    threading.Thread(target=run_script, args=(mode,)).start()
    return jsonify({"status": "started"})

@app.route('/api/records')
def get_records():
    json_path = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'processed_students.json')
    if os.path.exists(json_path):
        try:
            with open(json_path, 'r', encoding='utf-8') as f:
                data = json.load(f)
            return jsonify(data)
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    return jsonify([])

@app.route('/api/download/<filename>')
def download_file(filename):
    if filename in ["PDF_TrungTuyen_CNTT.zip", "PDF_TrungTuyen_QTKD.zip"]:
        base_dir = os.path.dirname(os.path.abspath(__file__))
        return send_from_directory(base_dir, filename, as_attachment=True)
    return "File not found", 404

@app.route('/api/stream')
def stream():
    def event_stream():
        while True:
            try:
                msg = log_queue.get(timeout=15.0)
                if msg == "===DONE===":
                    yield f"data: {json.dumps({'type': 'done'})}\n\n"
                    break
                else:
                    yield f"data: {json.dumps({'type': 'log', 'message': msg})}\n\n"
            except queue.Empty:
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"
    return Response(event_stream(), mimetype="text/event-stream")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    print(f"Khởi động Web Server tại http://0.0.0.0:{port}")
    app.run(debug=True, host='0.0.0.0', port=port, threaded=True, use_reloader=False)
