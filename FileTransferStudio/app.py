import os
from flask import Flask, render_template, request, jsonify, send_file
from werkzeug.utils import secure_filename
import tempfile
import logging
from utils import transfer_data_to_powerpoint, transfer_stickers_to_powerpoint # Added import for sticker function

# Configure logging
logging.basicConfig(level=logging.DEBUG)
logger = logging.getLogger(__name__)

# Create the Flask app
app = Flask(__name__)
app.secret_key = os.environ.get("FLASK_SECRET_KEY", "your-secret-key-here")

# Configure upload settings
ALLOWED_EXTENSIONS_EXCEL = {'xlsx', 'xls'}
ALLOWED_EXTENSIONS_PPT = {'pptx'}
UPLOAD_FOLDER = tempfile.gettempdir()  # Use system temp directory directly
app.config['UPLOAD_FOLDER'] = UPLOAD_FOLDER
app.config['MAX_CONTENT_LENGTH'] = 16 * 1024 * 1024  # 16MB max file size

def allowed_file_excel(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_EXCEL

def allowed_file_ppt(filename):
    return '.' in filename and filename.rsplit('.', 1)[1].lower() in ALLOWED_EXTENSIONS_PPT

@app.route('/')
def index():
    return render_template('index.html')

@app.route('/upload-excel', methods=['POST'])
def upload_excel():
    try:
        logger.debug("Excel upload request received")

        if 'excel_file' not in request.files:
            logger.error("No excel_file in request.files")
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['excel_file']
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file_excel(file.filename):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Invalid file type. Please upload an Excel file (.xlsx or .xls)'}), 400

        try:
            filename = f"excel_{secure_filename(file.filename)}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            logger.debug(f"Saving Excel file to: {filepath}")

            # Remove old file if it exists
            if os.path.exists(filepath):
                os.remove(filepath)

            file.save(filepath)
            app.config['EXCEL_PATH'] = filepath
            logger.debug(f"Excel file saved successfully at: {filepath}")
            return jsonify({'success': 'Excel file uploaded successfully'})

        except Exception as save_error:
            logger.error(f"Error saving Excel file: {str(save_error)}")
            return jsonify({'error': f'Error saving file: {str(save_error)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error in upload_excel: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

@app.route('/upload-powerpoint', methods=['POST'])
def upload_powerpoint():
    try:
        logger.debug("PowerPoint upload request received")

        if 'ppt_file' not in request.files:
            logger.error("No ppt_file in request.files")
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['ppt_file']
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file_ppt(file.filename):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Invalid file type. Please upload a PowerPoint file (.pptx)'}), 400

        try:
            filename = f"ppt_{secure_filename(file.filename)}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            logger.debug(f"Saving PowerPoint file to: {filepath}")

            # Remove old file if it exists
            if os.path.exists(filepath):
                os.remove(filepath)

            file.save(filepath)
            app.config['PPT_PATH'] = filepath
            logger.debug(f"PowerPoint file saved successfully at: {filepath}")
            return jsonify({'success': 'PowerPoint file uploaded successfully'})

        except Exception as save_error:
            logger.error(f"Error saving PowerPoint file: {str(save_error)}")
            return jsonify({'error': f'Error saving file: {str(save_error)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error in upload_powerpoint: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

@app.route('/transfer', methods=['POST'])
def transfer_data():
    try:
        excel_path = app.config.get('EXCEL_PATH')
        ppt_path = app.config.get('PPT_PATH')

        logger.debug(f"Transfer requested with Excel path: {excel_path} and PPT path: {ppt_path}")

        if not excel_path or not ppt_path:
            logger.error("Missing file paths in config")
            return jsonify({'error': 'Please upload both Excel and PowerPoint files first'}), 400

        if not os.path.exists(excel_path) or not os.path.exists(ppt_path):
            logger.error(f"Files missing - Excel exists: {os.path.exists(excel_path)}, PPT exists: {os.path.exists(ppt_path)}")
            return jsonify({'error': 'One or both uploaded files are missing. Please upload them again.'}), 400

        # Generate output path in temp directory
        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_presentation.pptx')
        logger.debug(f"Output will be saved to: {output_path}")

        # Remove old output file if it exists
        if os.path.exists(output_path):
            os.remove(output_path)

        # Transfer data
        transfer_data_to_powerpoint(excel_path, ppt_path, output_path)
        logger.debug("Data transfer completed successfully")

        return send_file(
            output_path,
            as_attachment=True,
            download_name='modified_presentation.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        logger.error(f"Error during transfer: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.route('/stickers')
def stickers():
    return render_template('stickers.html')

@app.route('/upload-excel-stickers', methods=['POST'])
def upload_excel_stickers():
    try:
        logger.debug("Excel stickers upload request received")

        if 'excel_file' not in request.files:
            logger.error("No excel_file in request.files")
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['excel_file']
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file_excel(file.filename):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Invalid file type. Please upload an Excel file (.xlsx or .xls)'}), 400

        try:
            filename = f"excel_stickers_{secure_filename(file.filename)}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            logger.debug(f"Saving Excel stickers file to: {filepath}")

            if os.path.exists(filepath):
                os.remove(filepath)

            file.save(filepath)
            app.config['EXCEL_STICKERS_PATH'] = filepath
            logger.debug(f"Excel stickers file saved successfully at: {filepath}")
            return jsonify({'success': 'Excel file uploaded successfully'})

        except Exception as save_error:
            logger.error(f"Error saving Excel stickers file: {str(save_error)}")
            return jsonify({'error': f'Error saving file: {str(save_error)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error in upload_excel_stickers: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

@app.route('/upload-powerpoint-stickers', methods=['POST'])
def upload_powerpoint_stickers():
    try:
        logger.debug("PowerPoint stickers upload request received")

        if 'ppt_file' not in request.files:
            logger.error("No ppt_file in request.files")
            return jsonify({'error': 'No file uploaded'}), 400

        file = request.files['ppt_file']
        if file.filename == '':
            logger.error("Empty filename")
            return jsonify({'error': 'No file selected'}), 400

        if not allowed_file_ppt(file.filename):
            logger.error(f"Invalid file type: {file.filename}")
            return jsonify({'error': 'Invalid file type. Please upload a PowerPoint file (.pptx)'}), 400

        try:
            filename = f"ppt_stickers_{secure_filename(file.filename)}"
            filepath = os.path.join(app.config['UPLOAD_FOLDER'], filename)
            logger.debug(f"Saving PowerPoint stickers file to: {filepath}")

            if os.path.exists(filepath):
                os.remove(filepath)

            file.save(filepath)
            app.config['PPT_STICKERS_PATH'] = filepath
            logger.debug(f"PowerPoint stickers file saved successfully at: {filepath}")
            return jsonify({'success': 'PowerPoint file uploaded successfully'})

        except Exception as save_error:
            logger.error(f"Error saving PowerPoint stickers file: {str(save_error)}")
            return jsonify({'error': f'Error saving file: {str(save_error)}'}), 500

    except Exception as e:
        logger.error(f"Unexpected error in upload_powerpoint_stickers: {str(e)}")
        return jsonify({'error': 'An unexpected error occurred'}), 500

@app.route('/transfer-stickers', methods=['POST'])
def transfer_stickers():
    try:
        excel_path = app.config.get('EXCEL_STICKERS_PATH')
        ppt_path = app.config.get('PPT_STICKERS_PATH')

        logger.debug(f"Stickers transfer requested with Excel path: {excel_path} and PPT path: {ppt_path}")

        if not excel_path or not ppt_path:
            logger.error("Missing file paths in config")
            return jsonify({'error': 'Please upload both Excel and PowerPoint files first'}), 400

        if not os.path.exists(excel_path) or not os.path.exists(ppt_path):
            logger.error(f"Files missing - Excel exists: {os.path.exists(excel_path)}, PPT exists: {os.path.exists(ppt_path)}")
            return jsonify({'error': 'One or both uploaded files are missing. Please upload them again.'}), 400

        output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_presentation_stickers.pptx')
        logger.debug(f"Output will be saved to: {output_path}")

        if os.path.exists(output_path):
            os.remove(output_path)

        # Use the sticker-specific transfer function
        transfer_stickers_to_powerpoint(excel_path, ppt_path, output_path)
        logger.debug("Stickers transfer completed successfully")

        return send_file(
            output_path,
            as_attachment=True,
            download_name='modified_presentation_stickers.pptx',
            mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation'
        )

    except Exception as e:
        logger.error(f"Error during stickers transfer: {str(e)}")
        return jsonify({'error': str(e)}), 500

@app.teardown_request
def cleanup_temp_files(exception=None):
    try:
        # Only clean up files after successful transfer
        if request.endpoint == 'transfer' and exception is None:
            excel_path = app.config.get('EXCEL_PATH')
            ppt_path = app.config.get('PPT_PATH')
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_presentation.pptx')

            for path in [excel_path, ppt_path, output_path]:
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                        logger.debug(f"Cleaned up file: {path}")
                    except Exception as e:
                        logger.error(f"Error cleaning up file {path}: {str(e)}")
        elif request.endpoint == 'transfer-stickers' and exception is None: #Added cleanup for stickers
            excel_path = app.config.get('EXCEL_STICKERS_PATH')
            ppt_path = app.config.get('PPT_STICKERS_PATH')
            output_path = os.path.join(app.config['UPLOAD_FOLDER'], 'modified_presentation_stickers.pptx')

            for path in [excel_path, ppt_path, output_path]:
                if path and os.path.exists(path):
                    try:
                        os.remove(path)
                        logger.debug(f"Cleaned up file: {path}")
                    except Exception as e:
                        logger.error(f"Error cleaning up file {path}: {str(e)}")

    except Exception as e:
        logger.error(f"Error in cleanup: {str(e)}")

if __name__ == '__main__':
    # Log the upload directory
    logger.info(f"Upload directory: {UPLOAD_FOLDER}")
    logger.info(f"Upload directory exists: {os.path.exists(UPLOAD_FOLDER)}")
    logger.info(f"Upload directory is writable: {os.access(UPLOAD_FOLDER, os.W_OK)}")

    app.run(host='0.0.0.0', port=5000, debug=True)