import os
import sys
import io
import json
import threading
import queue
from flask import Flask, request, render_template, Response, jsonify

app = Flask(__name__)

# Queue for real-time logs
log_queue = queue.Queue()

class StreamCapture(io.StringIO):
    def __init__(self, original_stdout):
        super().__init__()
        self.original_stdout = original_stdout

    def write(self, s):
        if s:
            self.original_stdout.write(s)
            self.original_stdout.flush()
            if s.strip():
                log_queue.put(s)

    def flush(self):
        self.original_stdout.flush()

# Replace global stdout to capture all prints
original_stdout = sys.stdout
sys.stdout = StreamCapture(original_stdout)

def run_script(mode):
    try:
        log_queue.put(">>> Bắt đầu chạy Pipeline...\n")
        
        # Save current working directory
        cwd = os.getcwd()
        
        if mode == 'overwrite':
            import full_automation_lark
            import importlib
            importlib.reload(full_automation_lark)
            full_automation_lark.run_automation()
        elif mode == 'update':
            import auto_update_new
            import importlib
            importlib.reload(auto_update_new)
            auto_update_new.run_automation()
        else:
            log_queue.put("LỖI: Chế độ chạy không hợp lệ.\n")
            
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
    with open('config.json', 'w', encoding='utf-8') as f:
        json.dump(config_data, f, indent=4)
    
    # Clear old queue
    while not log_queue.empty():
        log_queue.get()
        
    # Start thread
    threading.Thread(target=run_script, args=(mode,)).start()
    return jsonify({"status": "started"})

@app.route('/api/stream')
def stream():
    def event_stream():
        while True:
            msg = log_queue.get()
            if msg == "===DONE===":
                yield f"data: {json.dumps({'type': 'done'})}\n\n"
                break
            else:
                yield f"data: {json.dumps({'type': 'log', 'message': msg})}\n\n"
    return Response(event_stream(), mimetype="text/event-stream")

if __name__ == '__main__':
    port = int(os.environ.get("PORT", 5000))
    print(f"Khởi động Web Server tại http://0.0.0.0:{port}")
    # Tắt chế độ use_reloader vì nó có thể chạy 2 lần và lỗi stdout redirection
    app.run(debug=True, host='0.0.0.0', port=port, threaded=True, use_reloader=False)
