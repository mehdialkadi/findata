import os
import subprocess
import uuid
import shutil
import webbrowser
from flask import Flask, request, jsonify, render_template, send_from_directory
from script_dependencies.script import run_extraction
import threading
import time
import sys

last_heartbeat = time.time()

# Determine base path for bundled app or script
if getattr(sys, 'frozen', False):
    BASE_DIR = os.path.dirname(sys.executable)
else:
    BASE_DIR = os.path.dirname(os.path.abspath(__file__))

UPLOAD_FOLDER = os.path.join(BASE_DIR, 'uploads')
OUTPUT_FOLDER = os.path.join(BASE_DIR, 'outputs')
SCRIPT_DEP_FOLDER = os.path.join(BASE_DIR, 'script_dependencies')

app = Flask(__name__)
TASKS = {}

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(OUTPUT_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/health')
def health():
    return jsonify(status='ok')

@app.route('/heartbeat')
def heartbeat():
    global last_heartbeat
    last_heartbeat = time.time()
    return '', 204

@app.route('/upload', methods=['POST'])
def upload_file():
    file = request.files.get('pdf_file')
    if not file or not file.filename.endswith('.pdf'):
        return jsonify(status='error', message='Invalid file type. Please upload a PDF.'), 400

    task_id = str(uuid.uuid4())
    upload_path = os.path.join(UPLOAD_FOLDER, f"{task_id}.pdf")
    output_dir = os.path.join(OUTPUT_FOLDER, task_id)
    os.makedirs(output_dir, exist_ok=True)

    file.save(upload_path)

    TASKS[task_id] = {'status': 'processing', 'files': []}

    try:
        temp_path = os.path.join(SCRIPT_DEP_FOLDER, 'bilan.pdf')
        shutil.copy(upload_path, temp_path)
        run_extraction(temp_path)
        if os.path.exists(temp_path):
            os.remove(temp_path)

        output_files = []
        for f in os.listdir(SCRIPT_DEP_FOLDER):
            if f.endswith('.xlsx') and ('modèle' in f.lower() or 'extration' in f.lower()):
                src = os.path.join(SCRIPT_DEP_FOLDER, f)
                dst = os.path.join(output_dir, f)
                shutil.move(src, dst)
                output_files.append(f)

        TASKS[task_id]['status'] = 'completed'
        TASKS[task_id]['files'] = output_files
    except Exception as e:
        TASKS[task_id]['status'] = 'failed'
        TASKS[task_id]['error'] = str(e)

    return jsonify(status='success', task_id=task_id)

@app.route('/status/<task_id>')
def get_status(task_id):
    task = TASKS.get(task_id)
    if not task:
        return jsonify(status='error', message='Invalid task ID'), 404

    if task['status'] == 'completed':
        return jsonify(status='completed', files=task['files'])
    elif task['status'] == 'failed':
        return jsonify(status='failed', error=task.get('error', 'Unknown error'))
    else:
        return jsonify(status='processing')

@app.route('/download/<task_id>/<filename>')
def download_file(task_id, filename):
    task_output_dir = os.path.join(OUTPUT_FOLDER, task_id)
    if not os.path.exists(os.path.join(task_output_dir, filename)):
        return 'File not found', 404
    return send_from_directory(task_output_dir, filename, as_attachment=True)

def monitor_heartbeat():
    global last_heartbeat
    while True:
        time.sleep(5)
        if time.time() - last_heartbeat > 15:
            print("❌ No heartbeat in 15s. Shutting down...")
            os._exit(0)

def open_browser():
    time.sleep(1)
    webbrowser.open("http://127.0.0.1:5000/")

if __name__ == '__main__':
    threading.Thread(target=monitor_heartbeat, daemon=True).start()
    threading.Thread(target=open_browser, daemon=True).start()
    app.run(debug=False, use_reloader=False)
