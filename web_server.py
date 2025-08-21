import os
import json
import traceback
import time
from flask import Flask, request, jsonify, render_template, send_from_directory
from werkzeug.utils import secure_filename
import agent
import processor

# --- Configuration ---
UPLOAD_FOLDER = 'uploads'
STATIC_FOLDER = 'static'
ALLOWED_EXTENSIONS = {'xlsx', 'xls'}

app = Flask(__name__, template_folder='templates', static_folder=STATIC_FOLDER)
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER

# Ensure all necessary directories exist
os.makedirs(UPLOAD_FOLDER, exist_ok=True)
os.makedirs(os.path.join(STATIC_FOLDER, 'charts'), exist_ok=True)
os.makedirs(os.path.join(STATIC_FOLDER, 'outputs'), exist_ok=True)

# --- Helper Functions ---
def allowed_file(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS

# --- App Routes ---

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/static/<path:path>')
def send_static(path):
    return send_from_directory(STATIC_FOLDER, path)

@app.route('/api/upload', methods=['POST'])
def upload_file():
    if 'file' not in request.files:
        return jsonify({"error": "No file part"}), 400
    file = request.files['file']
    if file.filename == '':
        return jsonify({"error": "No selected file"}), 400
    if file and allowed_file(file.filename):
        filename = secure_filename(file.filename)
        file_path = os.path.join(app.config['UPLOAD_FOLDER'], filename)
        file.save(file_path)
        return jsonify({"success": True, "file_path": file_path})
    else:
        return jsonify({"error": "File type not allowed"}), 400

@app.route('/api/preview', methods=['POST'])
def preview_file():
    data = request.get_json()
    file_path = data.get('file_path')
    if not file_path or not os.path.exists(file_path):
        return jsonify({"error": "File not found or path is invalid."}), 400
    try:
        preview_data = processor.read_rows(file_path=file_path, limit=10)
        return jsonify({"success": True, "data": preview_data})
    except Exception as e:
        return jsonify({"error": str(e)}), 500

@app.route('/api/agent', methods=['POST'])
def run_agent():
    data = request.get_json()
    user_input = data.get('prompt')
    file_path = data.get('file_path')

    if not user_input or not file_path:
        return jsonify({"error": "Prompt and file_path are required."}), 400

    try:
        # The web server is responsible for controlling output paths.
        result = agent.run_agent_task(
            user_input,
            file_path,
            # Pass the correct, absolute paths for the output directories
            chart_output_dir=os.path.abspath(os.path.join(STATIC_FOLDER, 'charts')),
            file_output_dir=os.path.abspath(os.path.join(STATIC_FOLDER, 'outputs'))
        )

        # Post-process observations to create web-accessible artifacts
        artifacts = []
        for observation in result.get('observations', []):
            if not isinstance(observation, dict):
                continue

            if observation.get('success') and 'chart_path' in observation:
                chart_filename = os.path.basename(observation['chart_path'])
                artifacts.append({
                    "type": "chart",
                    "url": f'/static/charts/{chart_filename}'
                })

            if observation.get('success') and 'output_file' in observation:
                output_filename = os.path.basename(observation['output_file'])
                artifacts.append({
                    "type": "file",
                    "url": f'/static/outputs/{output_filename}',
                    "filename": output_filename
                })

        result['artifacts'] = artifacts
        return jsonify(result)

    except Exception as e:
        traceback.print_exc()
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True, port=5001)
