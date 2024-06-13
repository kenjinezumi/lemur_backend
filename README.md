# Lemur Backend

## Overview

This project consists of two Google Cloud Run services: `lemur-broker` and `lemur-generate`. These services work together to process requests to generate Google Slides presentations based on provided data.

- `lemur-broker`: Receives requests from the frontend, publishes them to a Pub/Sub topic.
- `lemur-generate`: Subscribes to the Pub/Sub topic, processes the messages, and generates Google Slides presentations. The result is published back to a response Pub/Sub topic.
- `lemur-broker`: Listens to the response Pub/Sub topic and sends the response back to the frontend.

## Architecture

1. **Frontend** sends a request to `lemur-broker`.
2. **Lemur Broker** publishes the request to a **Pub/Sub** topic.
3. **Lemur Generate** subscribes to the Pub/Sub topic, processes the message, and generates a Google Slides presentation. The result is published back to a response Pub/Sub topic.
4. **Lemur Broker** listens to the response Pub/Sub topic and sends the response back to the frontend.

## Environment Variables

Ensure the following environment variables are set:

- `DRIVE_FOLDER_ID`: Google Drive folder ID where presentations are saved.
- `API_ENDPOINT_URL`: URL of the API endpoint to fetch data.
- `PUBSUB_TOPIC`: Pub/Sub topic name.
- `PUBSUB_SUBSCRIPTION`: Pub/Sub subscription name.
- `PUBSUB_RESPONSE_TOPIC`: Pub/Sub response topic name.
- `PUBSUB_RESPONSE_SUBSCRIPTION`: Pub/Sub response subscription name.
- `SLIDES_TEMPLATE_ID`: Google Slides template ID.
- `GCP_PROJECT_ID`: Your Google Cloud Project ID.

## Deployment

### Create Pub/Sub Topics and Subscriptions

1. **Create Pub/Sub Topic and Subscription**:

    ```sh
    export PROJECT_ID=$(gcloud config get-value project)

    gcloud pubsub topics create lemur
    gcloud pubsub topics create lemur-response

    gcloud pubsub subscriptions create lemur-subscription --topic=lemur
    gcloud pubsub subscriptions create lemur-response-subscription --topic=lemur-response
    ```

### Build and Deploy Services

1. **Build and Push Docker Images**:

    ```sh

    gcloud builds submit --tag gcr.io/$PROJECT_ID/lemur-broker ./lemur-broker
    gcloud builds submit --tag gcr.io/$PROJECT_ID/lemur-generate ./lemur-generate
    ```

2. **Deploy to Cloud Run**:

    ```sh
    gcloud run deploy lemur-broker --image=gcr.io/$PROJECT_ID/lemur-broker --platform=managed --region=us-central1 --allow-unauthenticated \
    --set-env-vars GCP_PROJECT_ID=$PROJECT_ID,DRIVE_FOLDER_ID=1Zi9ejkrvwAOTlJm4VtEJBydWKHJgN8YF,API_ENDPOINT_URL=http://34.90.192.243/deman_gen_insights,PUBSUB_TOPIC=lemur,PUBSUB_RESPONSE_TOPIC=lemur-response,PUBSUB_RESPONSE_SUBSCRIPTION=lemur-response-subscription \
    --timeout=1800

    gcloud run deploy lemur-generate --image=gcr.io/$PROJECT_ID/lemur-generate --platform=managed --region=us-central1 --allow-unauthenticated \
    --set-env-vars GCP_PROJECT_ID=$PROJECT_ID,DRIVE_FOLDER_ID=1Zi9ejkrvwAOTlJm4VtEJBydWKHJgN8YF,PUBSUB_SUBSCRIPTION=lemur-subscription,PUBSUB_RESPONSE_TOPIC=lemur-response,SLIDES_TEMPLATE_ID=1Va_X2HGXRJSEoUJEPmO-CNqxUEoyxNj49sw_GdQeZa4 \
    --timeout=1800
    ```

## Usage

To test the service, you can send a request to `lemur-broker` using `curl`:

```sh
curl --max-time 600 -X POST "https://lemur-broker-dfrarc6doq-uc.a.run.app/generate" -H "Content-Type: application/json" -d '{"quarter_no": "Q1", "year_no" : 2024, "file_id": "22222"}'
