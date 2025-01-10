import os
import logging
from flask import Flask, render_template, request, jsonify, send_from_directory
from flask_socketio import SocketIO
from werkzeug.utils import secure_filename
from script_manager import ScriptManager
import config

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Initialize Flask app
app = Flask(__name__)
app.config['SECRET_KEY'] = os.urandom(24)
app.config['UPLOAD_FOLDER'] = config.UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = config.MAX_CONTENT_LENGTH
socketio = SocketIO(app)

# Initialize script manager
script_manager = ScriptManager(socketio)

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/download/<filename>')
def download_file(filename):
    """Handle file downloads from the output folder"""
    return send_from_directory(config.OUTPUT_FOLDER, filename, as_attachment=True)

@app.route('/upload', methods=['POST'])
def upload_files():
    """Handle file uploads"""
    try:
        if 'baseSheet' not in request.files or 'newDataSheet' not in request.files:
            return jsonify({'error': 'Both files are required'}), 400

        base_sheet = request.files['baseSheet']
        new_data_sheet = request.files['newDataSheet']

        if base_sheet.filename == '' or new_data_sheet.filename == '':
            return jsonify({'error': 'No selected files'}), 400

        if not all(config.allowed_file(f.filename) for f in [base_sheet, new_data_sheet]):
            return jsonify({'error': 'Invalid file type. Only Excel files (.xlsx, .xls) are allowed'}), 400

        # Save base sheet
        base_filename = secure_filename(base_sheet.filename)
        base_path = os.path.join(app.config['UPLOAD_FOLDER'], base_filename)
        base_sheet.save(base_path)

        # Save new data sheet
        new_data_filename = secure_filename(new_data_sheet.filename)
        new_data_path = os.path.join(app.config['UPLOAD_FOLDER'], new_data_filename)
        new_data_sheet.save(new_data_path)

        # Update script manager with new file paths
        script_manager.update_file_paths(base_path, new_data_path)

        return jsonify({
            'message': 'Files uploaded successfully',
            'base_sheet': base_filename,
            'new_data_sheet': new_data_filename
        })
    except Exception as e:
        logger.error(f"Error uploading files: {str(e)}")
        return jsonify({'error': str(e)}), 500

@socketio.on('start_pipeline')
def handle_pipeline_start(data):
    """Handle pipeline start request from client"""
    logger.debug(f"Starting pipeline execution with config: {data}")
    if data and 'steps' in data:
        script_manager.execute_pipeline(data['steps'])
    else:
        script_manager.execute_pipeline()

@socketio.on('update_pipeline_config')
def handle_pipeline_config_update(data):
    """Handle pipeline configuration update from client"""
    logger.debug(f"Updating pipeline configuration: {data}")
    if data and 'steps' in data:
        success = script_manager.update_pipeline_config(data['steps'])
        socketio.emit('pipeline_config_updated', {'success': success})

@socketio.on('connect')
def handle_connect():
    """Handle client connection"""
    logger.debug("Client connected")

@socketio.on('disconnect')
def handle_disconnect():
    """Handle client disconnection"""
    logger.debug("Client disconnected")