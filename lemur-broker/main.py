import os
import logging
import json
from flask import Flask, request, jsonify
from google.cloud import pubsub_v1, logging as cloud_logging

# Initialize Google Cloud Logging
client = cloud_logging.Client()
client.setup_logging()

# Flask app
app = Flask(__name__)

# Configure logging
logging.basicConfig(level=logging.INFO)

# Pub/Sub settings
PROJECT_ID = os.getenv('GCP_PROJECT_ID')
TOPIC_NAME = os.getenv('PUBSUB_TOPIC')
publisher = pubsub_v1.PublisherClient()

@app.route('/generate', methods=['POST'])
def generate():
    """
    Endpoint to receive a request to generate a presentation and publish the request to a Pub/Sub topic.
    """
    try:
        data = request.get_json()
        quarter_no = data.get('quarter_no')
        year_no = data.get('year_no')
        file_id = data.get('file_id')
        message_data = json.dumps(data).encode('utf-8')
        
        # Publish message to Pub/Sub
        topic_path = publisher.topic_path(PROJECT_ID, TOPIC_NAME)
        future = publisher.publish(topic_path, data=message_data)
        future.result()
        
        logging.info(f"Published message to {TOPIC_NAME}: {data}")
        return jsonify({"status": "Message published"}), 200
    except Exception as e:
        logging.error(f"Error: {e}")
        return jsonify({"error": str(e)}), 500

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
