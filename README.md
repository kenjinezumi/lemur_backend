# Demand Generation Insights

This is a Go application that exposes an API endpoint to generate a PowerPoint presentation based on insights retrieved from an external API. The generated PowerPoint is then uploaded to Google Drive.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Setup](#setup)
- [Running the Application](#running-the-application)
- [Deployment](#deployment)
- [API Endpoint](#api-endpoint)
- [License](#license)

## Prerequisites

- Go 1.20 or later
- Docker
- Google Cloud SDK
- Google Service Account with Drive API enabled

## Setup

1. Clone the repository:

2. Install dependencies:
    ```sh
    go mod download
    ```

3. Place your Google Cloud service account key file in the project root and name it `service-account-key.json`.

4. Set up environment variables:
    - `DRIVE_FOLDER_ID`: The ID of the Google Drive folder where the generated PowerPoint file will be uploaded.

## Running the Application

1. Run the application locally:
    ```sh
    go run main.go
    ```

2. The application will be available at `http://localhost:8080`.

## Deployment

1. Build the Docker image:
    ```sh
    docker build -t gcr.io/YOUR_PROJECT_ID/deman_gen_insights .
    ```

2. Push the Docker image to Google Container Registry:
    ```sh
    docker push gcr.io/YOUR_PROJECT_ID/deman_gen_insights
    ```

3. Create an `app.yaml` file in your project directory:
    ```yaml
    runtime: custom
    env: flex

    handlers:
    - url: /.*
      script: auto
    ```

4. Deploy to Google App Engine:
    ```sh
    gcloud app deploy
    ```

## API Endpoint

### Generate PowerPoint

- **URL:** `/generate`
- **Method:** `POST`
- **Content-Type:** `application/json`

#### Request Payload

```json
{
  "quarter_no": "Q1",
  "year_no": "2024",
  "file_id": "your-file-id"
}
