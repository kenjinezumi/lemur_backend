import os
import logging
import json
import requests
from flask import Flask, request, jsonify
from google.auth import default
from googleapiclient.discovery import build
import time
import threading
from pptx import Presentation
from googleapiclient.http import MediaFileUpload

# Flask app
app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Google Drive API setup
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations']
creds, project = default(scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

@app.route('/')
def index():
    return 'Lemur Service'

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy"}), 200

def log_elapsed_time(start_time, stop_event):
    while not stop_event.is_set():
        elapsed_time = time.time() - start_time
        logger.info(f"Elapsed time: {elapsed_time:.2f} seconds")
        time.sleep(10)

@app.route('/generate', methods=['POST'])
def generate():
    """
    Endpoint to receive a request to generate a presentation.
    """
    try:
        data = request.get_json()
        logger.info(f"Received request data: {data}")

        # Start a separate thread to log elapsed time
        start_time = time.time()
        stop_event = threading.Event()
        elapsed_time_logger = threading.Thread(target=log_elapsed_time, args=(start_time, stop_event))
        elapsed_time_logger.start()

        slide_numbers = [11]
        slide_data = {}

        # Fetch slide data for each slide number
        api_url = 'http://34.90.192.243/insight_slide'
        for slide_no in slide_numbers:
            response = requests.post(api_url, json={"slide_no": str(slide_no)}, timeout=3600)
            logger.info(f"API response status code for slide {slide_no}: {response.status_code}")
            logger.info(f"API response content for slide {slide_no}: {response.text}")

            response.raise_for_status()  # Raise an exception for HTTP errors

            slide_data[slide_no] = response.json()
            logger.info(f"Received slide data from API for slide {slide_no}: {slide_data[slide_no]}")

        # Generate the presentation
        presentation_link = create_presentation(slide_data, data['file_id'])
        logger.info(f"Generated presentation link: {presentation_link}")

        response_data = {
            "original_parameters": data,
            "presentation_link": presentation_link
        }
        return jsonify(response_data), 200
    except requests.exceptions.RequestException as e:
        logger.error(f"API request error: {e}")
        return jsonify({"error": "API request failed", "details": str(e)}), 500
    except Exception as e:
        logger.error(f"Error generating presentation: {e}")
        return jsonify({"error": str(e)}), 500
    finally:
        stop_event.set()
        elapsed_time_logger.join(1)

def create_presentation(data, file_id):
    """
    Create a presentation and populate it with data.
    """
    try:
        # Load the template presentation
        template_path = 'template.pptx'
        prs = Presentation(template_path)
        
        # Populate the presentation with data
        for slide_no, content in data.items():
            if slide_no - 1 < len(prs.slides):
                slide = prs.slides[slide_no - 1]  # Adjust index since slides are 0-indexed
                populate_slide(slide, content)
            else:
                logger.error(f"Slide number {slide_no} is out of range for the presentation")

        # Save the modified presentation
        output_path = f'/tmp/{file_id}.pptx'
        prs.save(output_path)
        
        # Upload the presentation to Google Drive
        file_metadata = {
            'name': f'Generated Presentation {file_id}',
            'parents': [os.getenv('DRIVE_FOLDER_ID')]
        }
        media = MediaFileUpload(output_path, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()

        logger.info(f"Uploaded presentation with ID: {uploaded_file.get('id')}")
        return f"https://drive.google.com/file/d/{uploaded_file.get('id')}/view"
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")

def populate_slide(slide, content):
    """
    Populate a slide with the given content.
    """
    try:
        # Update the slide content here based on the content structure
        if 'data' in content:
            # Log the data content for debugging
            logger.info(f"Populating slide with data: {content['data']}")

            # Find and populate the table
            for shape in slide.shapes:
                if shape.has_table:
                    table = shape.table
                    for row_idx, row_data in enumerate(content['data']):
                        for col_idx, cell_data in enumerate(row_data):
                            table.cell(row_idx, col_idx).text = str(cell_data)

        if 'insights' in content:
            # Log the insights content for debugging
            logger.info(f"Populating slide with insights: {content['insights']}")

            # Find and populate the text box
            for shape in slide.shapes:
                if shape.has_text_frame:
                    text_frame = shape.text_frame
                    text_frame.clear()  # Clear existing content
                    for paragraph in content['insights']:
                        p = text_frame.add_paragraph()
                        p.text = paragraph

        logger.info(f"Slide populated with content: {content}")
    except Exception as e:
        logger.error(f"Error populating slide: {e}")

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
