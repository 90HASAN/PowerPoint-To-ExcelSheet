import os
import io
import uuid
import tempfile
from flask import Flask, request, render_template, send_file, flash, redirect, url_for
from converter import process_ppt, write_excel

app = Flask(__name__)
app.secret_key = 'dev-secret-key-change-me'
app.config['MAX_CONTENT_LENGTH'] = 500 * 1024 * 1024  # 500 MB upload limit

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


@app.route('/convert', methods=['POST'])
def convert():
    files = request.files.getlist('ppt_files')
    files = [f for f in files if f and f.filename]

    if not files:
        flash('Please choose at least one .pptx file.')
        return redirect(url_for('index'))

    bad = [f.filename for f in files if not allowed_file(f.filename)]
    if bad:
        flash(f"Only .pptx files are allowed. Rejected: {', '.join(bad)}")
        return redirect(url_for('index'))

    with tempfile.TemporaryDirectory() as tmpdir:
        all_records = []
        for f in files:
            safe_name = f"{uuid.uuid4().hex}_{os.path.basename(f.filename)}"
            saved_path = os.path.join(tmpdir, safe_name)
            f.save(saved_path)
            try:
                records = process_ppt(saved_path)
            except Exception as e:
                flash(f"Failed to process '{f.filename}': {e}")
                return redirect(url_for('index'))
            all_records.extend(records)

        if not all_records:
            flash('No data could be extracted from the uploaded file(s).')
            return redirect(url_for('index'))

        excel_buffer = io.BytesIO()
        write_excel(all_records, excel_buffer)
        excel_buffer.seek(0)

    return send_file(
        excel_buffer,
        as_attachment=True,
        download_name='converted.xlsx',
        mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
    )


if __name__ == "__main__":
    app.run(debug=True)
