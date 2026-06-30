import os
import io
import uuid
import tempfile
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from converter import process_ppt, write_excel
from werkzeug.datastructures import FileStorage

app = Flask(__name__)
app.secret_key = 'dev-secret-key-change-me'

app.config['MAX_CONTENT_LENGTH'] = 200 * 1024 * 1024

@app.errorhandler(413)
def too_large(e):
    flash('That file is too large. Please upload .pptx file(s) totaling under 200 MB.')
    return redirect(url_for('index'))

ALLOWED_EXT = {'.pptx'}


def allowed_file(filename):
    return os.path.splitext(filename.lower())[1] in ALLOWED_EXT


@app.errorhandler(413)
def too_large(e):
    flash('That file is too large. Please upload a smaller .pptx file (max 500 MB).')
    return redirect(url_for('index'))


@app.route('/', methods=['GET'])
def index():
    return render_template('index.html')


import threading

jobs = {}  # job_id -> {"status": "pending"|"done"|"error", "result": BytesIO or None, "error": str or None}

def run_conversion(job_id, filepaths, filenames):
    try:
        all_records = []
        for path, name in zip(filepaths, filenames):
            all_records.extend(process_ppt(path))
        if not all_records:
            jobs[job_id] = {"status": "error", "error": "No data extracted."}
            return
        buf = io.BytesIO()
        write_excel(all_records, buf)
        buf.seek(0)
        jobs[job_id] = {"status": "done", "result": buf}
    except Exception as e:
        jobs[job_id] = {"status": "error", "error": str(e)}
    finally:
        for p in filepaths:
            try:
                os.remove(p)
            except OSError:
                pass

@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('ppt_files') or []
    files = [f for f in files if f and f.filename]
    if not files:
        flash('Please choose at least one .pptx file.')
        return redirect(url_for('index'))
    bad = [f.filename for f in files if not allowed_file(f.filename)]
    if bad:
        flash(f"Only .pptx files are allowed. Rejected: {', '.join(bad)}")
        return redirect(url_for('index'))

    job_id = uuid.uuid4().hex
    persist_dir = tempfile.mkdtemp()  # NOT auto-cleaned TemporaryDirectory, since job outlives this request
    saved_paths, names = [], []
    for f in files:
        safe_name = f"{uuid.uuid4().hex}_{os.path.basename(f.filename)}"
        path = os.path.join(persist_dir, safe_name)
        f.save(path)
        saved_paths.append(path)
        names.append(f.filename)

    jobs[job_id] = {"status": "pending"}
    threading.Thread(target=run_conversion, args=(job_id, saved_paths, names), daemon=True).start()
    return redirect(url_for('job_status', job_id=job_id))

@app.route('/status/<job_id>')
def job_status(job_id):
    job = jobs.get(job_id)
    if not job:
        flash('Unknown job.')
        return redirect(url_for('index'))
    if job["status"] == "pending":
        return render_template('processing.html', job_id=job_id)  # auto-refresh page
    if job["status"] == "error":
        flash(f"Conversion failed: {job['error']}")
        del jobs[job_id]
        return redirect(url_for('index'))
    buf = job["result"]
    del jobs[job_id]
    return send_file(buf, as_attachment=True, download_name='converted.xlsx',
                      mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')


if __name__ == "__main__":
    app.run(debug=True)
