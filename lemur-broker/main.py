import os
import logging
import json
from flask import Flask, request, jsonify
from google.cloud import pubsub_v1
from google.auth import default
from concurrent.futures import TimeoutError

# Initialize logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Flask app
app = Flask(__name__)

# Pub/Sub settings
PROJECT_ID = os.getenv('GCP_PROJECT_ID')
TOPIC_NAME = os.getenv('PUBSUB_TOPIC')
RESPONSE_TOPIC_NAME = os.getenv('PUBSUB_RESPONSE_TOPIC')
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

        # Subscribe to response topic
        subscription_name = f"projects/{PROJECT_ID}/subscriptions/{RESPONSE_TOPIC_NAME}-sub"
        response = listen_for_response(subscription_name)
        
        return jsonify(response), 200
    except Exception as e:
        logging.error(f"Error: {e}")
        return jsonify({"error": str(e)}), 500

def listen_for_response(subscription_name):
    """
    Listen for a response on the specified Pub/Sub subscription.
    """
    response = {}

    def callback(message):
        nonlocal response
        response = json.loads(message.data.decode('utf-8'))
        message.ack()

    streaming_pull_future = subscriber.subscribe(subscription_name, callback=callback)
    logging.info(f"Listening for messages on {subscription_name}")

    # Wrap subscriber in a 'with' block to automatically call close() when done.
    with subscriber:
        try:
            # Streaming pull future will exit once the callback has been invoked.
            streaming_pull_future.result(timeout=60)
        except TimeoutError:
            streaming_pull_future.cancel()
    
    return response

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
