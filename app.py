#!/usr/bin/env python3
"""
TUM ICI Results Processor — Web Interface
Run:  python app.py
Then open http://localhost:5000
"""

import os
import sys
import json
import uuid
import queue
import threading
from pathlib import Path
from flask import (Flask, request, jsonify, send_file,
                   render_template, Response, stream_with_context)

# ── Paths ─────────────────────────────────────────────────────
BASE_DIR     = Path(__file__).parent
UPLOADS_DIR  = BASE_DIR / 'uploads'
OUTPUTS_DIR  = BASE_DIR / 'outputs'
TEMPLATE_PATH = BASE_DIR / 'template.docx'

UPLOADS_DIR.mkdir(exist_ok=True)
OUTPUTS_DIR.mkdir(exist_ok=True)

# ── In-memory job store ───────────────────────────────────────
jobs: dict = {}
job_queues: dict = {}

app = Flask(__name__, template_folder='templates')
app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024   # 200 MB

# ═══════════════════════════════════════════════════════════════
#  ROUTES
# ═══════════════════════════════════════════════════════════════

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload():
    if 'pdf' not in request.files:
        return jsonify({'error': 'No file part'}), 400

    f = request.files['pdf']
    if not f.filename or not f.filename.lower().endswith('.pdf'):
        return jsonify({'error': 'Please upload a PDF file'}), 400

    exam_type = request.form.get('exam_type', 'ordinary')

    job_id   = str(uuid.uuid4())
    pdf_name = f'input_{job_id}.pdf'
    pdf_path = UPLOADS_DIR / pdf_name
    f.save(str(pdf_path))

    jobs[job_id] = {
        'status': 'queued', 'progress': 0, 'message': 'Job queued…',
        'output_file': None, 'error': None, 'exam_type': exam_type,
    }
    job_queues[job_id] = queue.Queue()

    t = threading.Thread(target=run_job, args=(job_id, pdf_path, exam_type), daemon=True)
    t.start()

    return jsonify({'job_id': job_id})

@app.route('/progress/<job_id>')
def progress(job_id):
    if job_id not in jobs: return jsonify({'error': 'Unknown job'}), 404

    def event_stream():
        q = job_queues.get(job_id)
        if q is None:
            yield _sse({'status': 'error', 'message': 'Job queue not found'})
            return
        while True:
            try:
                msg = q.get(timeout=30)
                yield _sse(msg)
                if msg.get('status') in ('done', 'error'): break
            except queue.Empty:
                yield ': ping\n\n'

    return Response(stream_with_context(event_stream()), mimetype='text/event-stream', headers={'Cache-Control': 'no-cache'})

@app.route('/status/<job_id>')
def status(job_id):
    if job_id not in jobs: return jsonify({'error': 'Unknown job'}), 404
    return jsonify(jobs[job_id])

@app.route('/download/<job_id>')
def download(job_id):
    if job_id not in jobs: return jsonify({'error': 'Unknown job'}), 404
    job = jobs[job_id]
    if job['status'] != 'done' or not job['output_file']:
        return jsonify({'error': 'File not ready'}), 400
    
    out_path = OUTPUTS_DIR / job['output_file']
    if not out_path.exists(): return jsonify({'error': 'Output file missing'}), 404
    
    mode_str = "Special_Supp" if job.get('exam_type') == 'supp' else "Ordinary"
    
    return send_file(
        str(out_path),
        as_attachment=True,
        download_name=f'ICI_Results_{mode_str}.docx',
        mimetype='application/vnd.openxmlformats-officedocument.wordprocessingml.document',
    )

# ═══════════════════════════════════════════════════════════════
#  BACKGROUND PROCESSING
# ═══════════════════════════════════════════════════════════════

def push(job_id: str, status: str, progress: int, message: str, output_file: str = None, error: str = None):
    jobs[job_id].update({'status': status, 'progress': progress, 'message': message, 'output_file': output_file, 'error': error})
    payload = {'status': status, 'progress': progress, 'message': message}
    if output_file: payload['output_file'] = output_file
    if error: payload['error'] = error
    q = job_queues.get(job_id)
    if q: q.put(payload)

def run_job(job_id: str, pdf_path: Path, exam_type: str):
    try:
        push(job_id, 'running', 5, 'Ensuring dependencies are installed…')
        _ensure_deps()

        # Import standalone scripts dynamically based on mode
        sys.path.insert(0, str(BASE_DIR))
        
        if exam_type == 'supp':
            from supp.extract import extract_pdf_data
            from supp.generate import generate as supp_generate
            generate_func = lambda j, o: supp_generate(str(TEMPLATE_PATH), j, o)
            out_prefix = 'supp'
        else:
            from extract import extract_pdf_data
            from generate import generate as ord_generate
            generate_func = lambda j, o: ord_generate(str(TEMPLATE_PATH), j, o)
            out_prefix = 'ordinary'

        # 1. Extract Data from PDF
        push(job_id, 'running', 20, 'Opening PDF and extracting records (this may take a minute)…')
        data = extract_pdf_data(pdf_path)
        record_count = len(data.get('records', []))
        push(job_id, 'running', 60, f'Found {record_count} records. Aggregating…')

        # 2. Save Temporary JSON for generate to read
        json_path = OUTPUTS_DIR / f'temp_{job_id}.json'
        with open(json_path, 'w') as f:
            json.dump(data, f)

        # 3. Generate the Word Document
        push(job_id, 'running', 80, 'Building the Word document from template…')
        out_name = f'{out_prefix}_results_{job_id}.docx'
        out_path = OUTPUTS_DIR / out_name
        
        generate_func(str(json_path), str(out_path))

        # Cleanup temporary JSON
        json_path.unlink(missing_ok=True)

        push(job_id, 'done', 100, 'Document ready!', output_file=out_name)

    except Exception as exc:
        import traceback
        tb = traceback.format_exc()
        push(job_id, 'error', 0, f'Error: {exc}', error=tb)
    finally:
        try: pdf_path.unlink(missing_ok=True)
        except Exception: pass

def _ensure_deps():
    import subprocess
    try:
        import fitz  # PyMuPDF
        import flask
    except ImportError:
        subprocess.run([sys.executable, '-m', 'pip', 'install', 'pymupdf', 'flask', '-q'], check=True)

def _sse(data: dict) -> str:
    return f'data: {json.dumps(data)}\n\n'

if __name__ == '__main__':
    port = int(os.environ.get('PORT', 5000))
    print(f'Starting TUM Results Processor on http://localhost:{port}')
    app.run(debug=False, host='0.0.0.0', port=port, threaded=True)