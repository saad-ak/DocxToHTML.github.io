from flask import Flask, request, render_template, redirect, send_file, flash
import os
from werkzeug.utils import secure_filename
import shutil
import zipfile

# Import your bulk_convert_docx_to_html function
from conversion.googleDocToHTMLBulk import bulk_convert_docx_to_html

app = Flask(__name__)
app.config['UPLOAD_FOLDER'] = 'uploads'
app.config['OUTPUT_FOLDER'] = 'outputs'
app.secret_key = os.urandom(24)  # Random 24-byte key

# Ensure folders exist
os.makedirs(app.config['UPLOAD_FOLDER'], exist_ok=True)
os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

# Allowed file extensions
ALLOWED_EXTENSIONS = {'docx'}

def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

def clean_up_folders(upload_folder, output_folder):
    """Delete all files in the upload and output folders."""
    # Clean up the upload folder
    for filename in os.listdir(upload_folder):
        file_path = os.path.join(upload_folder, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)  # Remove the file
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Remove the directory and its contents
        except Exception as e:
            print(f"Error cleaning up {file_path}: {e}")
    
    # Clean up the output folder
    for filename in os.listdir(output_folder):
        file_path = os.path.join(output_folder, filename)
        try:
            if os.path.isfile(file_path):
                os.remove(file_path)  # Remove the file
            elif os.path.isdir(file_path):
                shutil.rmtree(file_path)  # Remove the directory and its contents
        except Exception as e:
            print(f"Error cleaning up {file_path}: {e}")

@app.route('/')
def index():
    """Render the main page."""
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads."""
    if 'files[]' not in request.files:
        flash('No files part')
        return redirect('/')
    
    # Clean up previous uploads before accepting new files
    clean_up_folders(app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'])
    
    files = request.files.getlist('files[]')
    for file in files:
        if file and allowed_file(file.filename):
            filename = secure_filename(file.filename)
            file.save(os.path.join(app.config['UPLOAD_FOLDER'], filename))
    
    flash('Files uploaded successfully!')
    return redirect('/')

@app.route('/convert', methods=['POST'])
def convert_files():
    """Run the conversion script."""
    try:
        # Clear output folder before conversion
        shutil.rmtree(app.config['OUTPUT_FOLDER'])
        os.makedirs(app.config['OUTPUT_FOLDER'], exist_ok=True)

        # Run the conversion
        bulk_convert_docx_to_html(app.config['UPLOAD_FOLDER'], app.config['OUTPUT_FOLDER'])
        flash('Files converted successfully!')
    except Exception as e:
        flash(f'Error during conversion: {e}')
    
    return redirect('/')

@app.route('/download')
def download_files():
    """Zip and allow download of all HTML files."""
    zip_path = os.path.join(app.config['OUTPUT_FOLDER'], 'converted_files.zip')
    with zipfile.ZipFile(zip_path, 'w') as zipf:
        for root, _, files in os.walk(app.config['OUTPUT_FOLDER']):
            for file in files:
                if file.endswith('.html'):
                    file_path = os.path.join(root, file)
                    zipf.write(file_path, os.path.relpath(file_path, app.config['OUTPUT_FOLDER']))
    
    return send_file(zip_path, as_attachment=True)

if __name__ == '__main__':
    app.run(debug=True)
