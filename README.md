# Lemur Backend

This is a backend application that generates PowerPoint presentations based on given parameters and uploads them to Google Drive. The application is built with Flask and deployed on Google Cloud Run.

## Prerequisites

- Python 3.12
- Google Cloud SDK
- Docker

## Setup

1. Clone the repository:
    ```sh
    git clone https://github.com/yourusername/lemur_backend.git
    cd lemur_backend
    ```

2. Create and activate a virtual environment:
    ```sh
    python3.12 -m venv venv
    source venv/bin/activate
    ```

3. Install the dependencies:
    ```sh
    pip install -r requirements.txt
    ```

## Running Locally

1. Set the environment variables:
    ```sh
    export DRIVE_FOLDER_ID=your-google-drive-folder-id
    export API_ENDPOINT_URL=http://34.90.192.243/deman_gen_insights
    ```

2. Run the application:
    ```sh
    python main.py
    ```

3. The application will be available at `http://localhost:8080`.

## Building and Running with Docker

1. Build the Docker image:
    ```sh
    docker build -t lemur_backend .
    ```

2. Run the Docker container:
    ```sh
    docker run -e DRIVE_FOLDER_ID=your-google-drive-folder-id \
               -e API_ENDPOINT_URL=http://34.90.192.243/deman_gen_insights \
               -p 8080:8080 lemur_backend
    ```

3. The application will be available at `http://localhost:8080`.

## Deploying to Google Cloud Run

1. Submit a build to Google Cloud Build:
    ```sh
    gcloud builds submit --config cloudbuild.yaml .
    ```

2. Deploy the application to Cloud Run:
    ```sh
    gcloud run deploy lemur \
    --image gcr.io/$PROJECT_ID/lemur \
    --platform managed \
    --region us-central1 \
    --allow-unauthenticated \
    --set-env-vars DRIVE_FOLDER_ID=1Zi9ejkrvwAOTlJm4VtEJBydWKHJgN8YF,API_ENDPOINT_URL=http://34.90.192.243/deman_gen_insights \
    --min-instances 5
    ```

3. The application will be available at the URL provided by Cloud Run.

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
