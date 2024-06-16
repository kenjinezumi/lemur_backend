# Lemur Service

## Overview
Lemur Service is a Flask application designed to generate presentations based on data fetched from a specified API. It integrates with Google Drive for storing and sharing the generated presentations. The service exposes endpoints for health checks and generating presentations, leveraging Google APIs and various Python libraries for its functionality.

## Features
- **Health Check Endpoint**: Ensures the service is running correctly.
- **Presentation Generation**: Creates a presentation by fetching data from an API, processing the data, and populating a PowerPoint template.
- **Google Drive Integration**: Uploads the generated presentation to Google Drive and sets appropriate permissions.

## Requirements
- Python 3.7+
- Flask
- flasgger
- google-auth
- google-api-python-client
- python-pptx
- requests

## Local development

- Install the required Python packages:

```shell
pip install -r requirements.txt
```

- Configuration

Configure your environment variables in the deployment command to include:

- GCP_PROJECT_ID: Your Google Cloud project ID.
- DRIVE_FOLDER_ID: The ID of the Google Drive folder where presentations will be uploaded.
- API_ENDPOINT_URL: The URL of the API to fetch slide data from.

## Deployment
- Set your Google Cloud project ID:

```shell 
export PROJECT_ID=$(gcloud config get-value project)
```

- Build the docker image

```shell
gcloud builds submit --tag gcr.io/$PROJECT_ID/lemur-combined ./lemur-combined
```

- Deploy to Google Cloud Run

```shell
gcloud run deploy lemur-combined \
    --image gcr.io/$PROJECT_ID/lemur-combined \
    --platform managed \
    --region us-central1 \
    --allow-unauthenticated \
    --set-env-vars GCP_PROJECT_ID=your_project_id,DRIVE_FOLDER_ID=your_drive_folder_id,API_ENDPOINT_URL=http://your_api_endpoint,SLIDES_TEMPLATE_ID=your_template_id \
    --timeout 3600
```

- Usage

```shell
curl --max-time 3600 -X POST "https://your-service-url/generate" -H "Content-Type: application/json" -d '{"file_id": "your_file_id"}'
```