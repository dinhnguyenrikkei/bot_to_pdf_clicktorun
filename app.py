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
                s = s.decode('utf-8', errors='replace')
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
                    enc = self.original_stdout.encoding or 'utf-8'
                    self.original_stdout.write(s.encode(enc, errors='replace').decode(enc, errors='replace'))
                except Exception:
                    pass
            self.original_stdout.flush()
            if s.strip():
                log_queue.put(s)

    def flush(self):
        self.original_stdout.flush()

    def __getattr__(self, attr):
        return getattr(self.original_stdout, attr)

# Replace global stdout to capture all prints
original_stdout = sys.stdout
sys.stdout = StreamCapture(original_stdout)

def run_script(mode):
    try:
        log_queue.put(">>> Khởi chạy Pipeline trong tiến trình cô lập để tối ưu RAM...\n")
        
        script_name = 'full_automation_lark.py' if mode == 'overwrite' else 'auto_update_new.py'
        
        # Chạy tiến trình python độc lập để giải phóng RAM triệt để khi kết thúc
        import subprocess
        process = subprocess.Popen(
            [sys.executable, script_name],
            stdout=subprocess.PIPE,
            stderr=subprocess.STDOUT,
            text=True,
            bufsize=1, # Line-buffered để có log thời gian thực ngay lập tức
            encoding='utf-8',
            errors='replace'
        )
        
        # Đọc trực tiếp luồng stdout/stderr của tiến trình con và chuyển tiếp vào log_queue
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
        log_queue.put(f"\n[LỖI KHỞI CHẠY TIẾN TRÌNH CON]: {e}\n")
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
    with open('config.json', 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=4)
        
    # Clear old results files (JSON & ZIP) to ensure clean pipeline runs
    for f_path in ['processed_students.json', 'PDF_TrungTuyen_CNTT.zip', 'PDF_TrungTuyen_QTKD.zip']:
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
    if os.path.exists('processed_students.json'):
        try:
            with open('processed_students.json', 'r', encoding='utf-8') as f:
                data = json.load(f)
            return jsonify(data)
        except Exception as e:
            return jsonify({"error": str(e)}), 500
    return jsonify([])

@app.route('/api/download/<filename>')
def download_file(filename):
    if filename in ["PDF_TrungTuyen_CNTT.zip", "PDF_TrungTuyen_QTKD.zip"]:
        return send_from_directory(os.getcwd(), filename, as_attachment=True)
    return "File not found", 404

@app.route('/api/stream')
def stream():
    def event_stream():
        while True:
            try:
                # Đợi tối đa 15 giây, nếu trống thì gửi gói tin ping keep-alive
                msg = log_queue.get(timeout=15.0)
                if msg == "===DONE===":
                    yield f"data: {json.dumps({'type': 'done'})}\n\n"
                    break
                else:
                    yield f"data: {json.dumps({'type': 'log', 'message': msg})}\n\n"
            except queue.Empty:
                # Gửi sự kiện ping để tránh Render/Cloudflare/Browser ngắt kết nối do quá lâu không có dữ liệu
                yield f"data: {json.dumps({'type': 'ping'})}\n\n"
    return Response(event_stream(), mimetype="text/event-stream")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    print(f"Khởi động Web Server tại http://0.0.0.0:{port}")
    # Tắt chế độ use_reloader vì nó có thể chạy 2 lần và lỗi stdout redirection
    app.run(debug=True, host='0.0.0.0', port=port, threaded=True, use_reloader=False)
