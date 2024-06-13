# Lemur Backend

## Overview

This project consists of two Google Cloud Run services: `lemur-broker` and `lemur-generate`. These services work together to process requests to generate Google Slides presentations based on provided data.

- `lemur-broker`: Receives requests from the frontend, publishes them to a Pub/Sub topic.
- `lemur-generate`: Subscribes to the Pub/Sub topic, processes the messages, and generates Google Slides presentations.

## Architecture

1. **Frontend** sends a request to `lemur-broker`.
2. **Lemur Broker** publishes the request to a **Pub/Sub** topic.
3. **Lemur Generate** subscribes to the Pub/Sub topic, processes the message, and generates a Google Slides presentation.

## Environment Variables

Ensure the following environment variables are set:

- `DRIVE_FOLDER_ID`: Google Drive folder ID where presentations are saved.
- `API_ENDPOINT_URL`: URL of the API endpoint to fetch data.
- `PUBSUB_TOPIC`: Pub/Sub topic name.
- `PUBSUB_SUBSCRIPTION`: Pub/Sub subscription name.
- `SLIDES_TEMPLATE_ID`: Google Slides template ID.

## Deployment

1. **Build and Push Docker Images**:

    ```sh
    gcloud builds submit --config cloudbuild.yaml --substitutions _PUBSUB_TOPIC=your-pubsub-topic,_PUBSUB_SUBSCRIPTION=your-pubsub-subscription,_SLIDES_TEMPLATE_ID=your-template-id
    ```

2. **Deploy Services**:

    The deployment steps are included in the `cloudbuild.yaml` file, which automatically deploys the services after building and pushing the Docker images.

## Local Development

### Lemur Broker

1. **Navigate to the directory**:

    ```sh
    cd lemur-broker
    ```

2. **Run the service**:

    ```sh
    export DRIVE_FOLDER_ID=your-google-drive-folder-id
    export API_ENDPOINT_URL=http://34.90.192.243/deman_gen_insights
    export PUBSUB_TOPIC=your-pubsub-topic
    gunicorn -b :8080 main:app
    ```

### Lemur Generate

1. **Navigate to the directory**:

    ```sh
    cd lemur-generate
    ```

2. **Run the service**:

    ```sh
    export DRIVE_FOLDER_ID=your-google-drive-folder-id
    export API_ENDPOINT_URL=http://34.90.192.243/deman_gen_insights
    export PUBSUB_SUBSCRIPTION=your-pubsub-subscription
    export SLIDES_TEMPLATE_ID=your-template-id
    gunicorn -b :8080 main:app
    ```

## License

This project is licensed under the MIT License.
