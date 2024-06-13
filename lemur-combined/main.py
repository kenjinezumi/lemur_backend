import os
import re
import logging
import json
import requests
from flask import Flask, request, jsonify
from google.auth import default
from googleapiclient.discovery import build
import time
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
from googleapiclient.http import MediaFileUpload
from flasgger import Swagger

# Flask app
app = Flask(__name__)
swagger = Swagger(app)

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

# Google Drive API setup
SCOPES = ['https://www.googleapis.com/auth/drive', 'https://www.googleapis.com/auth/presentations']
creds, project = default(scopes=SCOPES)
drive_service = build('drive', 'v3', credentials=creds)

@app.route('/')
def index():
    return 'Lemur Service'

@app.route('/health', methods=['GET'])
def health_check():
    """
    Health check endpoint.
    ---
    responses:
      200:
        description: Service is healthy
    """
    return jsonify({"status": "healthy"}), 200

@app.route('/generate', methods=['POST'])
def generate():
    """
    Endpoint to generate a presentation.
    ---
    parameters:
      - name: data
        in: body
        required: true
        schema:
          type: object
          properties:
            file_id:
              type: string
              example: "1"
    responses:
      200:
        description: Presentation generated successfully
        schema:
          type: object
          properties:
            original_parameters:
              type: object
            presentation_link:
              type: string
      500:
        description: Error generating presentation
    """
    try:
        data = request.get_json()
        logger.info(f"Received request data: {data}")

        slide_numbers = [11, 14, 15, 16, 17]  
        slide_data = {}

        # Fetch slide data for each slide number
        api_url = 'http://34.90.192.243/insight_slide'
        for slide_no in slide_numbers:
            response = requests.post(api_url, json={"slide_no": str(slide_no)}, timeout=3600)
            logger.info(f"API response status code for slide {slide_no}: {response.status_code}")
            logger.info(f"API response content for slide {slide_no}: {response.text}")

            response.raise_for_status()  # Raise an exception for HTTP errors

            slide_data[slide_no] = response.json()
            logger.info(f"Received slide data from API for slide {slide_no}: {slide_data[slide_no]}")

        # Generate the presentation
        presentation_link = create_presentation(slide_data, data['file_id'])
        logger.info(f"Generated presentation link: {presentation_link}")

        response_data = {
            "original_parameters": data,
            "presentation_link": presentation_link
        }
        return jsonify(response_data), 200
    except requests.exceptions.RequestException as e:
        logger.error(f"API request error: {e}")
        return jsonify({"error": "API request failed", "details": str(e)}), 500
    except Exception as e:
        logger.error(f"Error generating presentation: {e}")
        return jsonify({"error": str(e)}), 500

def create_presentation(data, file_id):
    """
    Create a presentation and populate it with data.
    """
    try:
        # Load the template presentation
        template_path = 'template.pptx'
        prs = Presentation(template_path)
        
        # Populate the presentation with data
        for slide_no, content in data.items():
            if slide_no - 1 < len(prs.slides):
                slide = prs.slides[slide_no - 1]  # Adjust index since slides are 0-indexed
                populate_slide(slide, content, slide_no)
            else:
                logger.error(f"Slide number {slide_no} is out of range for the presentation")

        # Save the modified presentation
        output_path = f'/tmp/{file_id}.pptx'
        prs.save(output_path)
        
        # Upload the presentation to Google Drive
        file_metadata = {
            'name': f'Generated Presentation {file_id}'
        }
        media = MediaFileUpload(output_path, mimetype='application/vnd.openxmlformats-officedocument.presentationml.presentation')
        uploaded_file = upload_to_drive_with_retry(file_metadata, media)
        logger.info(f"Uploaded presentation with ID: {uploaded_file.get('id')}")

        # Set file permissions to make it accessible by anyone with the link
        permission = {
            'type': 'anyone',
            'role': 'reader',
        }
        drive_service.permissions().create(
            fileId=uploaded_file['id'],
            body=permission,
        ).execute()

        return f"https://drive.google.com/file/d/{uploaded_file.get('id')}/view"
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")
        raise e  # Re-raise the exception after logging it

def upload_to_drive_with_retry(file_metadata, media, retries=3):
    attempt = 0
    while attempt < retries:
        try:
            uploaded_file = drive_service.files().create(body=file_metadata, media_body=media, fields='id').execute()
            return uploaded_file
        except Exception as e:
            logger.error(f"Attempt {attempt+1} failed with error: {e}")
            attempt += 1
            time.sleep(2 ** attempt)  # Exponential backoff
    raise Exception("Failed to upload file after several retries")

def set_font(cell, font_name="Arial", font_size=8):
    """
    Set the font for a table cell.
    """
    for paragraph in cell.text_frame.paragraphs:
        for run in paragraph.runs:
            run.font.name = font_name
            run.font.size = Pt(font_size)

def set_yoy_color(cell, yoy_value):
    """
    Set the color for the YoY cell based on its value.
    """
    try:
        yoy = float(yoy_value.strip('%'))
        if yoy > 100:
            color = RGBColor(0, 255, 0)  # Green
        elif 90 <= yoy <= 100:
            color = RGBColor(218, 165, 32)  # Yellow goldenrod
        else:
            color = RGBColor(255, 0, 0)  # Red
        cell.text_frame.paragraphs[0].runs[0].font.color.rgb = color
    except ValueError:
        logger.error(f"Invalid YoY value: {yoy_value}")

def populate_slide(slide, content, slide_number):
    """
    Populate a slide with the given content.
    """
    try:
        main_table = None
        insights_table = None
        table_count = 0

        for shape in slide.shapes:
            if shape.has_table:
                if table_count == 0:
                    main_table = shape.table
                elif table_count == 1:
                    insights_table = shape.table
                table_count += 1

        if main_table:
            logger.info(f"Main table dimensions: {len(main_table.rows)} rows x {len(main_table.columns)} columns")

            if slide_number == 11:
                regions = ["NORTHAM", "EMEA", "JAPAC", "LATAM", "TOTAL"]
                metrics = [
                    ("Ent+Corp Pipeline", "QTD", "Attain"),
                    ("SMB Pipeline", "QTD", "Attain"),
                    ("Total Partner Marketing Sourced", "QTD", "Attain"),
                ]

                gcp_start_row = 3
                gws_start_row = 7

                for i, (metric, qtd_key, attain_key) in enumerate(metrics):
                    for j, region in enumerate(regions):
                        gcp_data = content.get("data", {}).get("GCP", {}).get(region, {}).get(metric, {})
                        gcp_value = f"${gcp_data.get(qtd_key, '')} ({gcp_data.get(attain_key, '')})"
                        cell = main_table.cell(gcp_start_row + i, 1 + j)
                        cell.text = gcp_value if gcp_value.strip() != "()" else ""
                        set_font(cell)
                        logger.info(f"Populated GCP {metric} for {region} with {gcp_value}")

                        gws_data = content.get("data", {}).get("GWS", {}).get(region, {}).get(metric, {})
                        gws_value = f"${gws_data.get(qtd_key, '')} ({gws_data.get(attain_key, '')})"
                        cell = main_table.cell(gws_start_row + i, 1 + j)
                        cell.text = gws_value if gws_value.strip() != "()" else ""
                        set_font(cell)
                        logger.info(f"Populated GWS {metric} for {region} with {gws_value}")

                # Ensure YoY values are applied correctly
                for region in regions:
                    for metric, _, _, yoy_key in [
                        ("Ent+Corp Pipeline", "QTD", "Attain", "YoY"),
                        ("SMB Pipeline", "QTD", "Attain", "YoY"),
                        ("Total Partner Marketing Sourced", "QTD", "Attain", "YoY")
                    ]:
                        yoy_data = content.get("data", {}).get("GCP", {}).get(region, {}).get(metric, {}).get(yoy_key, "")
                        cell = main_table.cell(gcp_start_row + metrics.index((metric, "QTD", "Attain")), 1 + regions.index(region))
                        set_yoy_color(cell, yoy_data)
                        yoy_data = content.get("data", {}).get("GWS", {}).get(region, {}).get(metric, {}).get(yoy_key, "")
                        cell = main_table.cell(gws_start_row + metrics.index((metric, "QTD", "Attain")), 1 + regions.index(region))
                        set_yoy_color(cell, yoy_data)

            elif slide_number in [14, 15, 16, 17]:
                regions = ["NORTHAM", "LATAM", "EMEA", "JAPAC", "PUBLIC SECTOR", "GLOBAL"]

                if slide_number == 14:
                    metrics = [
                        ("Direct Named QSOs", "QTD", "Attain", "YoY"),
                        ("Direct Named Pipeline", "QTD", "Attain", "YoY"),
                        ("Startup QSOs", "QTD", "Attain", "YoY"),
                        ("Startup Pipeline", "QTD", "Attain", "YoY"),
                        ("SMB QSOs", "QTD", "Attain", "YoY"),
                        ("SMB Pipeline", "QTD", "Attain", "YoY"),
                        ("Partner Pipeline", "QTD", "Attain", "YoY"),
                        ("GCP Direct QSOs", "QTD", "Attain", "YoY"),
                        ("GCP Direct + Partner Pipe", "QTD", "Attain", "YoY"),
                    ]
                if slide_number == 15:
                    metrics = [
                        ("Direct Named QSOs", "QTD", "Attain", "YoY"),
                        ("Direct Named Pipeline", "QTD", "Attain", "YoY"),
                        ("Startup QSOs", "QTD", "Attain", "YoY"),
                        ("Startup Pipeline", "QTD", "Attain", "YoY"),
                        ("SMB QSOs", "QTD", "Attain", "YoY"),
                        ("SMB Pipeline", "QTD", "Attain", "YoY"),
                        ("Partner Pipeline", "QTD", "Attain", "YoY"),
                        ("GCP QSOs", "QTD", "Attain", "YoY"),
                        ("GCP Direct + Partner Pipe", "QTD", "Attain", "YoY"),
                    ]
                if slide_number in [16, 17]:
                    metrics = [
                        ("Direct Named QSOs", "QTD", "Attain", "YoY"),
                        ("Direct Named Pipeline", "QTD", "Attain", "YoY"),
                        ("Partner Pipeline", "QTD", "Attain", "YoY"),
                        ("GWS QSOs", "QTD", "Attain", "YoY"),
                        ("GWS Direct + Partner Pipe", "QTD", "Attain", "YoY"),
                    ]

                start_row = 3

                for i, (metric, qtd_key, attain_key, yoy_key) in enumerate(metrics):
                    for j, region in enumerate(regions):
                        region_data = (
                            content.get("data", {})
                            .get(region, {})
                            .get(metric, {"QTD": "", "Attain": "", "YoY": ""})
                        )
                        region_value_qtd_attain = (
                            f"${region_data[qtd_key]} ({region_data[attain_key]})"
                        ) if region_data[qtd_key] or region_data[attain_key] else ""
                        region_value_yoy = region_data[yoy_key]

                        # QTD and Attain
                        cell = main_table.cell(start_row + i, 1 + (j * 2))
                        cell.text = region_value_qtd_attain
                        set_font(cell)

                        # YoY
                        cell = main_table.cell(start_row + i, 2 + (j * 2))
                        cell.text = region_value_yoy
                        set_font(cell)
                        set_yoy_color(cell, region_value_yoy)

                        logger.info(
                            f"Populated {metric} for {region} with QTD+Attain: {region_value_qtd_attain} and YoY: {region_value_yoy}"
                        )

        if insights_table:
            logger.info(
                f"Insights table dimensions: {len(insights_table.rows)} rows x {len(insights_table.columns)} columns"
            )

            # Populate insights table
            drivers = content.get("drivers", [])
            insights = content.get("insights", [])
            for i in range(max(len(drivers), len(insights))):  # Ensure we loop through the longest list
                if i < len(insights_table.rows) - 1:
                    driver_text = drivers[i] if i < len(drivers) else ""
                    insight_text = insights[i] if i < len(insights) else ""
                    
                    cell = insights_table.cell(i, 0)
                    set_font(cell)

                    # Bold the driver_text part
                    paragraph = cell.text_frame.paragraphs[0]
                    run = paragraph.add_run()
                    run.text = driver_text
                    run.font.bold = True
                    run.font.size = Pt(8)
                    run.font.name = 'Arial'
                    run = paragraph.add_run()
                    run.text = f" {insight_text}"
                    run.font.size = Pt(8)
                    run.font.name = 'Arial'
                    logger.info(f'Populated insight {i + 1}: {driver_text} {insight_text}')

            recommendations = content.get("recommendations", [])
            recommendations = [re.sub(r'\*\*', '', rec) for rec in recommendations]
            footnote_text = " ".join(recommendations)
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text += "\n" + footnote_text

    except Exception as e:
        logger.error(f"Error populating slide: {e}")
        raise e  # Re-raise the exception after logging it

if __name__ == '__main__':
    app.run(host='0.0.0.0', port=8080)
