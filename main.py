import os
import json
import logging
from flask import Flask, request, jsonify
import requests
from ppt.generator import create_powerpoint
from googleapiclient.discovery import build
from google.auth import default
from googleapiclient.http import MediaIoBaseUpload
import io

app = Flask(__name__)

DRIVE_FOLDER_ID = os.getenv("DRIVE_FOLDER_ID")
API_ENDPOINT_URL = os.getenv("API_ENDPOINT_URL")

# Set up logging
logging.basicConfig(level=logging.INFO)

# Google Drive authentication
def get_drive_service():
    credentials, project = default(scopes=["https://www.googleapis.com/auth/drive.file"])
    return build('drive', 'v3', credentials=credentials)

@app.route('/generate', methods=['POST'])
def generate():
    """
    Endpoint to generate a PowerPoint presentation based on the given quarter and year.
    """
    try:
        data = request.get_json()
        quarter_no = data.get('quarter_no')
        year_no = data.get('year_no')

        if not quarter_no or not year_no:
            logging.error("quarter_no and year_no are required")
            return jsonify({'error': 'quarter_no and year_no are required'}), 400

        # Call external API
        response = requests.post(API_ENDPOINT_URL, json={
            'quarter_no': quarter_no,
            'year_no': year_no
        }, timeout=600)

        if response.status_code != 200:
            logging.error("Error calling external API")
            return jsonify({'error': 'Error calling external API'}), 500

        api_response = response.json()
        pptx_file = create_powerpoint(api_response)

        # Upload to Google Drive
        drive_service = get_drive_service()
        file_metadata = {
            'name': f'Presentation_{quarter_no}_{year_no}.pptx',
            'parents': [DRIVE_FOLDER_ID]
        }
        media = MediaIoBaseUpload(io.BytesIO(pptx_file), mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

        file_id = file.get('id')
        file_link = f"https://drive.google.com/file/d/{file_id}/view"

        logging.info(f"File successfully uploaded with ID: {file_id}")

        return jsonify({
            'file_link': file_link,
            'file_id': file_id,
            'api_response': api_response,
            'parameters': data
        })
    except Exception as e:
        logging.error(f"An error occurred: {e}")
        return jsonify({'error': str(e)}), 500

if __name__ == '__main__':
    app.run(debug=True)
