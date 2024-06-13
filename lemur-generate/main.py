import os
import logging
import json
import threading
import requests
from flask import Flask, request, jsonify
from google.cloud import pubsub_v1, logging as cloud_logging, drive_v3, slides_v1
from google.auth import default

# Initialize Google Cloud Logging
cloud_client = cloud_logging.Client()
cloud_client.setup_logging()

# Flask app
app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Google Drive and Slides API setup
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations']
creds, project = default(scopes=SCOPES)
drive_service = drive_v3.DriveService(creds)
slides_service = slides_v1.PresentationService(creds)

# Pub/Sub settings
PROJECT_ID = os.getenv('GCP_PROJECT_ID')
SUBSCRIPTION_NAME = os.getenv('PUBSUB_SUBSCRIPTION')
RESPONSE_TOPIC_NAME = os.getenv('PUBSUB_RESPONSE_TOPIC')
subscriber = pubsub_v1.SubscriberClient()

@app.route('/')
def index():
    return 'Lemur Generate Service'

@app.route('/health', methods=['GET'])
def health_check():
    return jsonify({"status": "healthy"}), 200

def callback(message):
    data = json.loads(message.data.decode('utf-8'))
    logger.info(f"Received message: {data}")
    try:
        presentation_link = process_message(data)
        response_data = {
            "original_parameters": data,
            "presentation_link": presentation_link
        }
        publish_response(response_data)
        logger.info(f"Processed message and published response: {response_data}")
    except Exception as e:
        logger.error(f"Error processing message: {e}")
    message.ack()

def subscribe():
    subscription_path = subscriber.subscription_path(PROJECT_ID, SUBSCRIPTION_NAME)
    streaming_pull_future = subscriber.subscribe(subscription_path, callback=callback)
    logger.info(f"Listening for messages on {subscription_path}")
    try:
        streaming_pull_future.result()
    except Exception as e:
        logger.error(f"Listening for messages on {subscription_path} threw an exception: {e}")
        streaming_pull_future.cancel()

def process_message(data):
    """
    Process the message received from Pub/Sub and generate a presentation.
    """
    try:
        file_id = data['file_id']
        # Call the API to get the slide data
        api_url = os.getenv('API_ENDPOINT_URL')
        response = requests.post(api_url, json={"quarter_no": data['quarter_no'], "year_no": data['year_no'], "file_id": file_id})
        slide_data = response.json()

        presentation_link = create_presentation(slide_data, file_id)
        logger.info(f"Presentation created for file ID: {file_id}")
        return presentation_link
    except Exception as e:
        logger.error(f"Error processing message: {e}")
        return None

def create_presentation(data, file_id):
    """
    Create a presentation and populate it with data.
    """
    try:
        # Copy the template presentation
        template_id = os.getenv('SLIDES_TEMPLATE_ID')
        copied_file = {
            'name': 'Generated Presentation',
            'parents': [os.getenv('DRIVE_FOLDER_ID')]
        }
        presentation = drive_service.files().copy(
            fileId=template_id, body=copied_file).execute()

        presentation_id = presentation['id']
        logger.info(f"Created presentation with ID: {presentation_id}")

        # Populate the presentation with data
        for slide_number, slide_content in data.items():
            slide_id = f'slide_{slide_number}'
            populate_slide(presentation_id, slide_id, slide_content)

        logger.info(f"Presentation populated with data: {data}")
        return f"https://docs.google.com/presentation/d/{presentation_id}"
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")

def populate_slide(presentation_id, slide_id, content):
    """
    Populate a slide with the given content.
    """
    try:
        requests = []
        if 'data' in content:
            # Populate table
            requests.append({
                'updateTableCells': {
                    'tableRange': {
                        'tableCellLocation': {
                            'rowIndex': 0,
                            'columnIndex': 0,
                            'tableStartLocation': {
                                'slideObjectId': slide_id
                            }
                        },
                        'rowSpan': len(content['data']),
                        'columnSpan': len(content['data'][0])
                    },
                    'text': {
                        'textElements': [
                            {'textRun': {'content': str(cell)}}
                            for row in content['data']
                            for cell in row
                        ]
                    },
                    'fields': 'content'
                }
            })

        if 'insights' in content:
            # Populate insights
            requests.extend([{
                'insertText': {
                    'objectId': slide_id,
                    'text': f"\u2022 {insight}\n"
                }
            } for insight in content['insights']])

        # Execute the batch update
        body = {
            'requests': requests
        }
        slides_service.presentations().batchUpdate(
            presentationId=presentation_id, body=body).execute()

        logger.info(f"Slide {slide_id} populated with content: {content}")
    except Exception as e:
        logger.error(f"Error populating slide {slide_id}: {e}")

def publish_response(response_data):
    """
    Publish the response data to the response Pub/Sub topic.
    """
    publisher = pubsub_v1.PublisherClient()
    response_topic_path = publisher.topic_path(PROJECT_ID, RESPONSE_TOPIC_NAME)
    message_data = json.dumps(response_data).encode('utf-8')
    future = publisher.publish(response_topic_path, data=message_data)
    future.result()
    logger.info(f"Published response message to {RESPONSE_TOPIC_NAME}: {response_data}")

if __name__ == '__main__':
    subscription_thread = threading.Thread(target=subscribe)
    subscription_thread.start()
    app.run(host='0.0.0.0', port=8080)
