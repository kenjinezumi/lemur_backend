import os
import logging
import json
import time
from flask import Flask, request, jsonify
from google.cloud import pubsub_v1
from google.auth import default

logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask app
app = Flask(__name__)

# Pub/Sub settings
PROJECT_ID = os.getenv('GCP_PROJECT_ID')
TOPIC_NAME = os.getenv('PUBSUB_TOPIC')
RESPONSE_SUBSCRIPTION_NAME = os.getenv('PUBSUB_RESPONSE_SUBSCRIPTION')
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
        
        # Wait for the response
        response = wait_for_response(data['file_id'])
        return jsonify(response), 200
    except Exception as e:
        logging.error(f"Error: {e}")
        return jsonify({"error": str(e)}), 500

def wait_for_response(file_id):
    """
    Wait for the response message on the response subscription.
    """
    subscription_path = subscriber.subscription_path(PROJECT_ID, RESPONSE_SUBSCRIPTION_NAME)
    response_message = None

    def callback(message):
        nonlocal response_message
        message_data = json.loads(message.data.decode('utf-8'))
        if message_data['original_parameters']['file_id'] == file_id:
            response_message = message_data
            message.ack()

    streaming_pull_future = subscriber.subscribe(subscription_path, callback=callback)
    logging.info("Listening for response messages on %s", subscription_path)
    
    with subscriber:
        timeout = time.time() + 1800  # 30 minutes timeout
        while response_message is None:
            if time.time() > timeout:
                streaming_pull_future.cancel()
                raise TimeoutError("Timeout waiting for the response.")
            time.sleep(5)
    
    streaming_pull_future.cancel()
    return response_message

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
