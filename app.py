import os
from flask import Flask, request, send_from_directory, render_template
import subprocess

app = Flask(__name__, template_folder='web/templates', static_folder='web/static')
UPLOAD_FOLDER = 'uploads'
DOWNLOAD_FOLDER = 'downloads'
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['DOWNLOAD_FOLDER'] = DOWNLOAD_FOLDER

os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(DOWNLOAD_FOLDER, exist_ok=True)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return 'No file part'
    file = request.files['file']
    if file.filename == '':
        return 'No selected file'
    if file:
        input_filepath = os.path.join(app.config['UPLOAD_FOLDER'], file.filename)
        output_filepath = os.path.join(app.config['DOWNLOAD_FOLDER'], 'output.pptx')
        file.save(input_filepath)

        # Command to run the pptx-tool with the provided options
        command = [
            'poetry', 'run', 'pptx-tool', 'fix-font',
            '--theme=themes/interdisplay-pretendard.json', input_filepath, output_filepath
        ]
        
        try:
            subprocess.run(command, check=True)
        except subprocess.CalledProcessError as e:
            return f'Error processing file: {str(e)}'

        return send_from_directory(app.config['DOWNLOAD_FOLDER'], 'output.pptx', as_attachment=True)

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080, debug=True)
