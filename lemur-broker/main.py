import os
import logging
import json
from flask import Flask, request, jsonify
from google.cloud import pubsub_v1

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask app
app = Flask(__name__)

# Pub/Sub settings
PROJECT_ID = os.getenv('GCP_PROJECT_ID')
TOPIC_NAME = os.getenv('PUBSUB_TOPIC')
RESPONSE_TOPIC_NAME = os.getenv('PUBSUB_RESPONSE_TOPIC')
SUBSCRIPTION_NAME = os.getenv('PUBSUB_RESPONSE_SUBSCRIPTION')
publisher = pubsub_v1.PublisherClient()
subscriber = pubsub_v1.SubscriberClient()

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
        
        # Wait for the response from the response topic
        response_data = wait_for_response(file_id)
        if response_data:
            return jsonify(response_data), 200
        else:
            return jsonify({"error": "No response received"}), 500
    except Exception as e:
        logging.error(f"Error: {e}")
        return jsonify({"error": str(e)}), 500

def wait_for_response(file_id):
    """
    Wait for a response message for the given file_id.
    """
    response_data = {}

    def callback(message):
        nonlocal response_data
        data = json.loads(message.data.decode('utf-8'))
        if data['original_parameters']['file_id'] == file_id:
            response_data = data
            message.ack()
            streaming_pull_future.cancel()  # Cancel the streaming pull

    subscription_path = subscriber.subscription_path(PROJECT_ID, SUBSCRIPTION_NAME)
    streaming_pull_future = subscriber.subscribe(subscription_path, callback=callback)
    
    try:
        streaming_pull_future.result(timeout=1800)  # Timeout in seconds
    except Exception as e:
        logging.error(f"Error waiting for response: {e}")
        streaming_pull_future.cancel()
    
    return response_data

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
