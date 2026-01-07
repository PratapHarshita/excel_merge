from flask import Blueprint, render_template, request, send_file
from flask import after_this_request
from werkzeug.utils import secure_filename
import os
from .utils import merge_files_flexible, cleanup_files
import config

excel_merger_bp = Blueprint('excel_merger', __name__, template_folder='templates', static_folder='static')

@excel_merger_bp.route('/', methods=['GET', 'POST'])
def index():
    return render_template('index.html')


# -------------------
# Merge route
# -------------------
@excel_merger_bp.route('/merge', methods=['POST'])
def merge():
    files = request.files.getlist('files')
    merge_options = request.form.getlist('merge_option')
    col_values = request.form.getlist('col_value')  # number input per file

    if not files or len(files) < 2:
        return "Upload at least 2 Excel files", 400

    if len(merge_options) != len(files):
        return "Mismatch between files and options", 400

    saved_files = []
    for file in files:
        if file.filename != '' and file.filename.split('.')[-1].lower() in config.ALLOWED_EXTENSIONS:
            filename = secure_filename(file.filename)
            filepath = os.path.join(config.UPLOAD_FOLDER, filename)
            file.save(filepath)
            saved_files.append(filepath)

    # Convert column values to int or None
    col_values = [int(c) if c else None for c in col_values]

    output_path = merge_files_flexible(saved_files, merge_options, col_values)

    @after_this_request
    def cleanup(response):
        cleanup_files(saved_files + [output_path])
        return response
    return send_file(output_path, as_attachment=True)


# -------------------
# Split route
# -------------------
from flask import request, jsonify, send_file
import pandas as pd
import zipfile
import os
import config

@excel_merger_bp.route('/read_headers', methods=['POST'])
def read_headers():
    """Return Excel headers for preview"""
    file = request.files.get('file')
    if not file:
        return jsonify({'columns': []})
    filepath = os.path.join(config.UPLOAD_FOLDER, secure_filename(file.filename))
    file.save(filepath)
    df = pd.read_excel(filepath)
    return jsonify({'columns': list(df.columns)})

@excel_merger_bp.route('/split', methods=['POST'])
def split():
    file = request.files.get('file')
    filename = secure_filename(file.filename)
    filepath = os.path.join(config.UPLOAD_FOLDER, filename)
    file.save(filepath)

    repeat_cols = [int(x) for x in request.form.getlist('repeat_cols')]
    columns_per_file = []

    i = 0
    while True:
        cols = request.form.getlist(f'columns_file_{i}')
        if not cols:
            break
        columns_per_file.append([int(c) for c in cols])
        i += 1

    import pandas as pd, zipfile

    df = pd.read_excel(filepath)
    output_files = []

    for idx, cols in enumerate(columns_per_file):
        final_cols = sorted(set(cols + repeat_cols))
        out_path = os.path.join(
            config.UPLOAD_FOLDER,
            f"split_part_{idx+1}.xlsx"
        )
        df.iloc[:, final_cols].to_excel(out_path, index=False)
        output_files.append(out_path)

    zip_path = os.path.join(config.UPLOAD_FOLDER, "split_files.zip")
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for f in output_files:
            zipf.write(f, os.path.basename(f))

    @after_this_request
    def cleanup(response):
        cleanup_files([filepath, zip_path] + output_files)
        return response

    return send_file(zip_path, as_attachment=True)
