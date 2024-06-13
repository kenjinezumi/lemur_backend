import logging
from pptx import Presentation
from pptx.util import Pt
from pptx.dml.color import RGBColor
import re

# Configure logging
logging.basicConfig(level=logging.INFO)
logger = logging.getLogger(__name__)

def create_presentation(data, file_id, slide_number):
    """
    Create a presentation and populate it with data.
    """
    try:
        # Load the template presentation
        template_path = "template.pptx"
        prs = Presentation(template_path)

        # Populate the specific slide with data
        if slide_number - 1 < len(prs.slides):
            slide = prs.slides[
                slide_number - 1
            ]  # Adjust index since slides are 0-indexed
            populate_slide(slide, data, slide_number)
        else:
            logger.error(
                f"Slide number {slide_number} is out of range for the presentation"
            )

        # Save the modified presentation
        output_path = f"{file_id}.pptx"
        prs.save(output_path)
        logger.info(f"Presentation saved as {output_path}")

        return output_path
    except Exception as e:
        logger.error(f"Error creating presentation: {e}")

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

def populate_slide(slide, data, slide_number):
    """
    Populate a slide with the given content.
    """
    try:
        # Assuming the table structure is known and fixed
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
            logger.info(
                f"Main table dimensions: {len(main_table.rows)} rows x {len(main_table.columns)} columns"
            )

            # Add the logic for each slide
            if slide_number == 11:
                # Logic for Slide 11
                regions = ["NORTHAM", "EMEA", "JAPAC", "LATAM", "Total"]
                metrics = [
                    ("Ent+Corp Pipeline", "QTD", "Attain"),
                    ("SMB Pipeline", "QTD", "Attain"),
                    ("Total Partner Marketing Sourced", "QTD", "Attain"),
                ]
                product = ["GCP", "GWS"]

                gcp_start_row = 3
                gws_start_row = 7

                for i, (metric, qtd_key, attain_key) in enumerate(metrics):
                        for j, region in enumerate(regions):
                            gcp_data = data.get("data", {}).get("GCP", {}).get(region, {}).get(metric, {"QTD": "", "Attain": ""})
                            print(data.get("GCP", {}))

                            gcp_value = f"{gcp_data[qtd_key]} ({gcp_data[attain_key]})"
                            cell = main_table.cell(gcp_start_row + i, 1 + j)
                            cell.text = gcp_value
                            set_font(cell)
                            logger.info(
                                f"Populated GCP {metric} for {region} with {gcp_value}"
                            )
                            gws_data = data.get("data", {}).get("GWS", {}).get(region, {}).get(metric, {"QTD": "", "Attain": ""})
                            gws_value = f"{gws_data[qtd_key]} ({gws_data[attain_key]})"
                            cell = main_table.cell(gws_start_row + i, 1 + j)
                            cell.text = gws_value
                            set_font(cell)
                            logger.info(
                                f"Populated GWS {metric} for {region} with {gws_value}"
                            )

            elif slide_number == 14 or slide_number == 15 or slide_number == 16 or slide_number == 17:
                # Logic for Slides 14, 15, 16, 17
                regions = [
                    "NORTHAM",
                    "LATAM",
                    "EMEA",
                    "JAPAC",
                    "PUBLIC SECTOR",
                    "GLOBAL",
                ]

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
                

                if slide_number == 16:
                    metrics = [
                        ("Direct Named QSOs", "QTD", "Attain", "YoY"),
                        ("Direct Named Pipeline", "QTD", "Attain", "YoY"),
                        ("Partner Pipeline", "QTD", "Attain", "YoY"),
                        ("GWS QSOs", "QTD", "Attain", "YoY"),
                        ("GWS Direct + Partner Pipe", "QTD", "Attain", "YoY"),
                    ]


                if slide_number == 17:
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
                                data.get("data", {})
                                .get(region, {})
                                .get(metric, {"QTD": "", "Attain": "", "YoY": ""})
                            )
                            region_value_qtd_attain = (
                                f"{region_data[qtd_key]} ({region_data[attain_key]})"
                            )
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
            drivers = data.get("drivers", [])
            insights = data.get("insights", [])
            for i in range(len(drivers)):
                if i < len(insights_table.rows) - 1:
                    if i < len(drivers):
                        drivers[i] = re.sub(r'\*', '', drivers[i])

                    if i < len(insights):
                        insights[i] = re.sub(r'\*', '', insights[i])
                    driver_text = drivers[i] if i < len(drivers) else ""
                    insight_text = insights[i] if i < len(insights) else ""
                            
                    cell = insights_table.cell(i + 1, 0)
                    # cell.text = f"{insight_text}"
                    set_font(cell)

                            # Bold the driver_text part
                    paragraph = cell.text_frame.paragraphs[0]
                    run = paragraph.add_run()
                    run.text = driver_text
                    run.font.bold = True
                    run.font.size = Pt(8)
                    run.font.name = 'Arial'
                    run = paragraph.add_run()
                    run.text = insight_text
                    run.font.size = Pt(8)
                    run.font.name = 'Arial'
                    logger.info('Populated insights')


                    
            recommendations = data.get("recommendations", [])
            recommendations = [re.sub(r'\*\*', '', rec) for rec in recommendations]
            footnote_text = " ".join(recommendations)
            notes_slide = slide.notes_slide
            text_frame = notes_slide.notes_text_frame
            text_frame.text += "\n" + footnote_text


    except Exception as e:
        logger.error(f"Error populating slide: {e}")

# Example usage
if __name__ == "__main__":
    # Hardcoded data for local testing
    slide_data_11 = {
        "data": {
            "GCP": {
                "NORTHAM": {
                    "Ent+Corp Pipeline": {"QTD": "579.0M", "Attain": "32.0%"},
                    "SMB Pipeline": {"QTD": "55.9M", "Attain": "49.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "634.9M",
                        "Attain": "33.0%",
                    },
                },
                "LATAM": {
                    "Ent+Corp Pipeline": {"QTD": "152.6M", "Attain": "50.0%"},
                    "SMB Pipeline": {"QTD": "32.3M", "Attain": "66.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "184.9M",
                        "Attain": "53.0%",
                    },
                },
                "EMEA": {
                    "Ent+Corp Pipeline": {"QTD": "572.8M", "Attain": "71.0%"},
                    "SMB Pipeline": {"QTD": "69.6M", "Attain": "48.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "642.4M",
                        "Attain": "67.0%",
                    },
                },
                "JAPAC": {
                    "Ent+Corp Pipeline": {"QTD": "438.4M", "Attain": "56.0%"},
                    "SMB Pipeline": {"QTD": "75.4M", "Attain": "58.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "513.8M",
                        "Attain": "56.0%",
                    },
                },
                "TOTAL": {
                    "Ent+Corp Pipeline": {"QTD": "1775.8M", "Attain": "47.0%"},
                    "SMB Pipeline": {"QTD": "233.2M", "Attain": "53.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "2009.1M",
                        "Attain": "48.0%",
                    },
                },
            },
            "GWS": {
                "NORTHAM": {
                    "Ent+Corp Pipeline": {"QTD": "90.2M", "Attain": "30.0%"},
                    "SMB Pipeline": {"QTD": "36.7M", "Attain": "69.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "126.9M",
                        "Attain": "36.0%",
                    },
                },
                "LATAM": {
                    "Ent+Corp Pipeline": {"QTD": "79.5M", "Attain": "32.0%"},
                    "SMB Pipeline": {"QTD": "60.7M", "Attain": "inf%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "140.2M",
                        "Attain": "56.0%",
                    },
                },
                "EMEA": {
                    "Ent+Corp Pipeline": {"QTD": "82.7M", "Attain": "69.0%"},
                    "SMB Pipeline": {"QTD": "40.7M", "Attain": "77.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "123.4M",
                        "Attain": "71.0%",
                    },
                },
                "JAPAC": {
                    "Ent+Corp Pipeline": {"QTD": "138.5M", "Attain": "46.0%"},
                    "SMB Pipeline": {"QTD": "37.3M", "Attain": "33.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "175.8M",
                        "Attain": "42.0%",
                    },
                },
                "TOTAL": {
                    "Ent+Corp Pipeline": {"QTD": "394.3M", "Attain": "40.0%"},
                    "SMB Pipeline": {"QTD": "175.3M", "Attain": "80.0%"},
                    "Total Partner Marketing Sourced ": {
                        "QTD": "569.7M",
                        "Attain": "47.0%",
                    },
                },
            },
        },
        "insights": [
            "**GCP demonstrates a strong affinity with Enterprise and Corporate clients, achieving a 47% global attainment rate, while GWS excels in the SMB segment with an 80% attainment rate. This suggests a need to tailor demand generation strategies to specific customer segments.**",
            "**EMEA emerges as a leader in both GCP and GWS attainment across all segments, indicating the effectiveness of regional demand generation programs and sales engagement. This success story can offer valuable insights for other regions.**",
            "**North America presents a significant growth opportunity with a healthy pipeline but lagging attainment rates for both GCP (33%) and GWS (36%). Addressing potential bottlenecks in lead quality, campaign effectiveness, and sales follow-up is crucial.**",
        ],
        "recommendations": [
            "**Conduct a comparative analysis of successful demand generation tactics employed in EMEA and adapt them for other regions, particularly North America.**",
            "**Investigate the factors contributing to lower attainment rates in North America, including lead quality, campaign effectiveness, and sales follow-up processes. Implement corrective measures to optimize the sales funnel.**",
            "**Validate the exceptionally high GWS attainment rate in LATAM's SMB segment and address any data reporting discrepancies. Leverage insights from accurate data to replicate successful strategies in other regions.**",
        ],
        "drivers": [
            "Google Cloud Platform shows attainment rates of 60.3% in Corporate, 65.6% in Enterprise, 60.2% in SMB, 63.3% in Select, and 70.7% in Startup, while Google Workspace attainment rates are 43.5% in Corporate, 51.8% in Enterprise, 23.7% in SMB, 60.5% in Select, and 51.1% in Startup. \n",
            "Google Cloud Platform drives 79.7% of opportunities, while Google Workspace contributes 20.3%, demonstrating their respective influence in the EMEA region. \n",
            "Core Products represent the largest contributing product category for both GCP (86.7%) and GWS (57.6%) in North America. \n",
        ],
        "codes": [
            '```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv("data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv")\n\n# Calculate the percentage of opportunities for each product family and account segment\nproduct_family_segment_counts = df.groupby(["Opportunity_Product_Family", "Account_Segment"]).size().unstack(fill_value=0)\nproduct_family_segment_percentages = (product_family_segment_counts / product_family_segment_counts.sum(axis=1)) * 100\n\n# Print the results\nprint("Percentage of opportunities for each product family and account segment:")\nprint(product_family_segment_percentages.round(3))\n\n# Calculate the attainment rate for each product family and account segment\nproduct_family_segment_attainment = df.groupby(["Opportunity_Product_Family", "Account_Segment"])["Opportunity_Target_Attainment_Source"].apply(lambda x: (x == "Direct").sum() / len(x))\n\n# Print the results\nprint("\\nAttainment rate for each product family and account segment:")\nprint(product_family_segment_attainment.round(3))\n```',
            '```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv("data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv")\n\n# Filter for EMEA region\ndf_emea = df[df["Grouped_Marketing_Target_Region"] == "EMEA"]\n\n# Calculate the percentage of opportunities for each product family\ngcp_percentage = (df_emea["Opportunity_Product_Family"] == "Google Cloud Platform").mean()\ngws_percentage = (df_emea["Opportunity_Product_Family"] == "Google Workspace").mean()\n\n# Print the results\nprint(f"In EMEA, Google Cloud Platform accounts for {gcp_percentage:.3f} of opportunities, while Google Workspace accounts for {gws_percentage:.3f} of opportunities.")\n```',
            '```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv("data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv")\n\n# Filter data for North America\ndf_north_america = df[df["Grouped_Marketing_Target_Region"] == "NORTHAM"]\n\n# Calculate product category percentages for GCP and GWS\ngcp_product_categories = df_north_america[df_north_america["Opportunity_Product_Family"] == "Google Cloud Platform"]["Opportunity_Product_Category"].value_counts(normalize=True)\ngws_product_categories = df_north_america[df_north_america["Opportunity_Product_Family"] == "Google Workspace"]["Opportunity_Product_Category"].value_counts(normalize=True)\n\n# Print the results\nprint("Product Category Percentages for GCP in North America:")\nprint(gcp_product_categories.round(3))\nprint("\\nProduct Category Percentages for GWS in North America:")\nprint(gws_product_categories.round(3))\n```\n## Product Category Percentages for GCP in North America:\n\nCore Products: 0.677\nCloud AI Products: 0.223\nGemini for Workspace Products: 0.100\n\n## Product Category Percentages for GWS in North America:\n\nCore Products: 0.857\nCloud AI Products: 0.143\n',
        ],
    }

    slide_data_14 = {
        "data": {
            "NORTHAM": {
                "Direct Named QSOs": {
                    "QTD": "5.8K",
                    "Attain": "45.0%",
                    "YoY": "156.0%",
                },
                "Direct Named Pipeline": {
                    "QTD": "5528.8M",
                    "Attain": "93.0%",
                    "YoY": "316.0%",
                },
                "Startup QSOs": {"QTD": "4.3K", "Attain": "31.0%", "YoY": "55.0%"},
                "Startup Pipeline": {
                    "QTD": "359.9M",
                    "Attain": "40.0%",
                    "YoY": "88.0%",
                },
                "SMB QSOs": {"QTD": "3.5K", "Attain": "48.0%", "YoY": "162.0%"},
                "SMB Pipeline": {"QTD": "186.7M", "Attain": "57.0%", "YoY": "179.0%"},
                "Partner Pipeline": {
                    "QTD": "636.8M",
                    "Attain": "33.0%",
                    "YoY": "92.0%",
                },
                "GCP Direct QSOs": {"QTD": "13.6K", "Attain": "40.0%", "YoY": "113.0%"},
                "GCP Direct + Partner Pipe": {
                    "QTD": "6712.3M",
                    "Attain": "74.0%",
                    "YoY": "250.0%",
                },
            },
            "LATAM": {
                "Direct Named QSOs": {"QTD": "1.6K", "Attain": "48.0%", "YoY": "75.0%"},
                "Direct Named Pipeline": {
                    "QTD": "438.2M",
                    "Attain": "94.0%",
                    "YoY": "295.0%",
                },
                "Startup QSOs": {"QTD": "0.4K", "Attain": "23.0%", "YoY": "6.0%"},
                "Startup Pipeline": {"QTD": "14.5M", "Attain": "25.0%", "YoY": "39.0%"},
                "SMB QSOs": {"QTD": "1.9K", "Attain": "39.0%", "YoY": "41.0%"},
                "SMB Pipeline": {"QTD": "44.0M", "Attain": "30.0%", "YoY": "76.0%"},
                "Partner Pipeline": {
                    "QTD": "184.2M",
                    "Attain": "52.0%",
                    "YoY": "106.0%",
                },
                "GCP Direct QSOs": {"QTD": "3.8K", "Attain": "40.0%", "YoY": "48.0%"},
                "GCP Direct + Partner Pipe": {
                    "QTD": "680.8M",
                    "Attain": "67.0%",
                    "YoY": "189.0%",
                },
            },
            "EMEA": {
                "Direct Named QSOs": {"QTD": "6.3K", "Attain": "41.0%", "YoY": "66.0%"},
                "Direct Named Pipeline": {
                    "QTD": "3310.0M",
                    "Attain": "72.0%",
                    "YoY": "193.0%",
                },
                "Startup QSOs": {"QTD": "0.9K", "Attain": "28.0%", "YoY": "-5.0%"},
                "Startup Pipeline": {"QTD": "90.1M", "Attain": "45.0%", "YoY": "45.0%"},
                "SMB QSOs": {"QTD": "5.0K", "Attain": "41.0%", "YoY": "57.0%"},
                "SMB Pipeline": {"QTD": "184.1M", "Attain": "41.0%", "YoY": "55.0%"},
                "Partner Pipeline": {
                    "QTD": "633.8M",
                    "Attain": "67.0%",
                    "YoY": "179.0%",
                },
                "GCP Direct QSOs": {"QTD": "12.1K", "Attain": "39.0%", "YoY": "54.0%"},
                "GCP Direct + Partner Pipe": {
                    "QTD": "4218.0M",
                    "Attain": "68.0%",
                    "YoY": "175.0%",
                },
            },
            "JAPAC": {
                "Direct Named QSOs": {"QTD": "3.4K", "Attain": "34.0%", "YoY": "42.0%"},
                "Direct Named Pipeline": {
                    "QTD": "1548.8M",
                    "Attain": "75.0%",
                    "YoY": "253.0%",
                },
                "Startup QSOs": {"QTD": "0.5K", "Attain": "30.0%", "YoY": "35.0%"},
                "Startup Pipeline": {
                    "QTD": "26.7M",
                    "Attain": "37.0%",
                    "YoY": "136.0%",
                },
                "SMB QSOs": {"QTD": "3.2K", "Attain": "38.0%", "YoY": "39.0%"},
                "SMB Pipeline": {"QTD": "106.3M", "Attain": "46.0%", "YoY": "85.0%"},
                "Partner Pipeline": {
                    "QTD": "506.2M",
                    "Attain": "55.0%",
                    "YoY": "130.0%",
                },
                "GCP Direct QSOs": {"QTD": "7.1K", "Attain": "35.0%", "YoY": "40.0%"},
                "GCP Direct + Partner Pipe": {
                    "QTD": "2187.9M",
                    "Attain": "66.0%",
                    "YoY": "201.0%",
                },
            },
            "PUBLIC SECTOR": {
                "Direct Named QSOs": {"QTD": "0.4K", "Attain": "45.0%", "YoY": "62.0%"},
                "Direct Named Pipeline": {
                    "QTD": "157.5M",
                    "Attain": "53.0%",
                    "YoY": "57.0%",
                },
                "Partner Pipeline": {
                    "QTD": "33.1M",
                    "Attain": "60.0%",
                    "YoY": "128.0%",
                },
                "GCP Direct QSOs": {"QTD": "0.4K", "Attain": "45.0%", "YoY": "62.0%"},
                "GCP Direct + Partner Pipe": {
                    "QTD": "190.6M",
                    "Attain": "54.0%",
                    "YoY": "66.0%",
                },
            },
            "GLOBAL": {
                "Direct Named QSOs": {
                    "QTD": "17.4K",
                    "Attain": "41.0%",
                    "YoY": "82.0%",
                },
                "Direct Named Pipeline": {
                    "QTD": "10983.2M",
                    "Attain": "82.0%",
                    "YoY": "253.0%",
                },
                "Startup QSOs": {"QTD": "6.1K", "Attain": "30.0%", "YoY": "36.0%"},
                "Startup Pipeline": {
                    "QTD": "491.1M",
                    "Attain": "40.0%",
                    "YoY": "79.0%",
                },
                "SMB QSOs": {"QTD": "13.6K", "Attain": "42.0%", "YoY": "66.0%"},
                "SMB Pipeline": {"QTD": "521.1M", "Attain": "45.0%", "YoY": "94.0%"},
                "Partner Pipeline": {
                    "QTD": "1994.1M",
                    "Attain": "47.0%",
                    "YoY": "126.0%",
                },
                "GCP Direct QSOs": {"QTD": "37.0K", "Attain": "39.0%", "YoY": "67.0%"},
                "GCP Direct + Partner Pipe": {
                    "QTD": "13989.5M",
                    "Attain": "70.0%",
                    "YoY": "209.0%",
                },
            },
        }
    }


   

    slide_data_15 = {"data":{"NORTHAM":{"Direct Named QSOs":{"QTD":"5.8K","Attain":"45.0%","YoY":"156.0%"},"Direct Named Pipeline":{"QTD":"5528.8M","Attain":"93.0%","YoY":"316.0%"},"Startup QSOs":{"QTD":"4.3K","Attain":"31.0%","YoY":"55.0%"},"Startup Pipeline":{"QTD":"359.9M","Attain":"40.0%","YoY":"88.0%"},"SMB QSOs":{"QTD":"3.5K","Attain":"48.0%","YoY":"162.0%"},"SMB Pipeline":{"QTD":"186.7M","Attain":"57.0%","YoY":"179.0%"},"Partner Pipeline":{"QTD":"636.8M","Attain":"33.0%","YoY":"92.0%"},"GCP Direct QSOs":{"QTD":"13.6K","Attain":"40.0%","YoY":"113.0%"},"GCP Direct + Partner Pipe":{"QTD":"6712.3M","Attain":"74.0%","YoY":"250.0%"}},"LATAM":{"Direct Named QSOs":{"QTD":"1.6K","Attain":"48.0%","YoY":"75.0%"},"Direct Named Pipeline":{"QTD":"438.2M","Attain":"94.0%","YoY":"295.0%"},"Startup QSOs":{"QTD":"0.4K","Attain":"23.0%","YoY":"6.0%"},"Startup Pipeline":{"QTD":"14.5M","Attain":"25.0%","YoY":"39.0%"},"SMB QSOs":{"QTD":"1.9K","Attain":"39.0%","YoY":"41.0%"},"SMB Pipeline":{"QTD":"44.0M","Attain":"30.0%","YoY":"76.0%"},"Partner Pipeline":{"QTD":"184.2M","Attain":"52.0%","YoY":"106.0%"},"GCP Direct QSOs":{"QTD":"3.8K","Attain":"40.0%","YoY":"48.0%"},"GCP Direct + Partner Pipe":{"QTD":"680.8M","Attain":"67.0%","YoY":"189.0%"}},"EMEA":{"Direct Named QSOs":{"QTD":"6.3K","Attain":"41.0%","YoY":"66.0%"},"Direct Named Pipeline":{"QTD":"3310.0M","Attain":"72.0%","YoY":"193.0%"},"Startup QSOs":{"QTD":"0.9K","Attain":"28.0%","YoY":"-5.0%"},"Startup Pipeline":{"QTD":"90.1M","Attain":"45.0%","YoY":"45.0%"},"SMB QSOs":{"QTD":"5.0K","Attain":"41.0%","YoY":"57.0%"},"SMB Pipeline":{"QTD":"184.1M","Attain":"41.0%","YoY":"55.0%"},"Partner Pipeline":{"QTD":"633.8M","Attain":"67.0%","YoY":"179.0%"},"GCP Direct QSOs":{"QTD":"12.1K","Attain":"39.0%","YoY":"54.0%"},"GCP Direct + Partner Pipe":{"QTD":"4218.0M","Attain":"68.0%","YoY":"175.0%"}},"JAPAC":{"Direct Named QSOs":{"QTD":"3.4K","Attain":"34.0%","YoY":"42.0%"},"Direct Named Pipeline":{"QTD":"1548.8M","Attain":"75.0%","YoY":"253.0%"},"Startup QSOs":{"QTD":"0.5K","Attain":"30.0%","YoY":"35.0%"},"Startup Pipeline":{"QTD":"26.7M","Attain":"37.0%","YoY":"136.0%"},"SMB QSOs":{"QTD":"3.2K","Attain":"38.0%","YoY":"39.0%"},"SMB Pipeline":{"QTD":"106.3M","Attain":"46.0%","YoY":"85.0%"},"Partner Pipeline":{"QTD":"506.2M","Attain":"55.0%","YoY":"130.0%"},"GCP Direct QSOs":{"QTD":"7.1K","Attain":"35.0%","YoY":"40.0%"},"GCP Direct + Partner Pipe":{"QTD":"2187.9M","Attain":"66.0%","YoY":"201.0%"}},"PUBLIC SECTOR":{"Direct Named QSOs":{"QTD":"0.4K","Attain":"45.0%","YoY":"62.0%"},"Direct Named Pipeline":{"QTD":"157.5M","Attain":"53.0%","YoY":"57.0%"},"Partner Pipeline":{"QTD":"33.1M","Attain":"60.0%","YoY":"128.0%"},"GCP Direct QSOs":{"QTD":"0.4K","Attain":"45.0%","YoY":"62.0%"},"GCP Direct + Partner Pipe":{"QTD":"190.6M","Attain":"54.0%","YoY":"66.0%"}},"GLOBAL":{"Direct Named QSOs":{"QTD":"17.4K","Attain":"41.0%","YoY":"82.0%"},"Direct Named Pipeline":{"QTD":"10983.2M","Attain":"82.0%","YoY":"253.0%"},"Startup QSOs":{"QTD":"6.1K","Attain":"30.0%","YoY":"36.0%"},"Startup Pipeline":{"QTD":"491.1M","Attain":"40.0%","YoY":"79.0%"},"SMB QSOs":{"QTD":"13.6K","Attain":"42.0%","YoY":"66.0%"},"SMB Pipeline":{"QTD":"521.1M","Attain":"45.0%","YoY":"94.0%"},"Partner Pipeline":{"QTD":"1994.1M","Attain":"47.0%","YoY":"126.0%"},"GCP Direct QSOs":{"QTD":"37.0K","Attain":"39.0%","YoY":"67.0%"},"GCP Direct + Partner Pipe":{"QTD":"13989.5M","Attain":"70.0%","YoY":"209.0%"}}},"insights":["**NORTHAM Leads in Direct Named Pipeline but Faces Conversion Challenges:** NORTHAM demonstrates exceptional YoY growth in Direct Named Pipeline at 316%, exceeding all other regions. However, QSO attainment in NORTHAM lags significantly at 48%, indicating a potential bottleneck in converting pipeline to qualified opportunities. This suggests a need to examine lead qualification, sales enablement, and market saturation in the region.","**EMEA Excels in Partner Engagement, While NORTHAM Presents Untapped Potential:** EMEA showcases outstanding Partner Pipeline performance with 179% YoY growth and 67% attainment, highlighting effective partner engagement strategies. In contrast, NORTHAM, despite having the largest Partner Pipeline, achieves only 33% attainment. This presents a significant opportunity to replicate EMEA's best practices and enhance NORTHAM's partner programs to maximize partner-driven growth.","**LATAM's Startup Segment Raises Concerns:** LATAM's Startup segment shows concerning signs with a meager 6% YoY QSO growth and a low 23% attainment rate. This suggests challenges in attracting and engaging startups in the region, potentially due to product-market fit issues, competition, or ineffective outreach. Addressing these challenges requires tailoring the value proposition, exploring partnerships, and leveraging digital channels to improve performance."],"recommendations":["**Improve Lead Conversion in NORTHAM:** Investigate and address factors hindering QSO attainment in NORTHAM, such as lead qualification processes, sales enablement resources, and market saturation levels. Implement targeted initiatives to improve lead conversion rates and maximize returns from the strong pipeline growth.","**Replicate EMEA's Partner Success in NORTHAM:** Analyze and replicate EMEA's best practices for partner engagement and enablement within the NORTHAM region. Enhance partner programs, provide targeted support, and foster stronger relationships to drive higher attainment rates and capitalize on NORTHAM's large partner ecosystem.","**Revitalize LATAM's Startup Engagement:** Conduct a thorough review of the Startup segment strategy in LATAM, including value proposition, target audience, and channel effectiveness. Tailor offerings to better address the specific needs and challenges of startups in the region, explore partnerships to expand reach, and leverage digital channels for effective outreach and engagement."],"drivers":["NORTHAM's Direct Named Pipeline comprises 82.37% Core products and 17.63% AI products, all of which fall under the Google Cloud Platform product family. \n","While EMEA accounts for 34.772% of the Partner Pipeline, NORTHAM represents 28.341%, signifying the untapped potential within the region. \n","Core products constitute the majority of LATAM's Startup segment, representing 92.1% of the total, while AI products make up the remaining 7.9%. \n"],"codes":["```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv(\"data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv\")\n\n# Filter for NORTHAM region and Direct source\nnortham_direct = df[(df[\"Grouped_Marketing_Target_Region\"] == \"NORTHAM\") & (df[\"Opportunity_Target_Attainment_Source\"] == \"Direct\")]\n\n# Calculate product category percentages\nproduct_category_percentages = northam_direct[\"Product_Category\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"Product Category Percentages for NORTHAM Direct:\")\nprint(product_category_percentages.to_string())\n\n# Calculate product family percentages\nproduct_family_percentages = northam_direct[\"Opportunity_Product_Family\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"\\nProduct Family Percentages for NORTHAM Direct:\")\nprint(product_family_percentages.to_string())\n```\n\n**Output:**\n\n```\nProduct Category Percentages for NORTHAM Direct:\nCore    72.73%\nAI      27.27%\nName: Product_Category, dtype: float64\n\nProduct Family Percentages for NORTHAM Direct:\nGoogle Cloud Platform    85.19%\nGoogle Workspace         14.81%\nName: Opportunity_Product_Family, dtype: float64\n```\n\n**Summary:**\n\n* **Product Categories:**\n    * Core products contribute to 72.73% of the Direct Named Pipeline in NORTHAM.\n    * AI products contribute to 27.27%.\n* **Product Families:**\n    * Google Cloud Platform accounts for 85.19% of the Direct Named Pipeline in NORTHAM.\n    * Google Workspace accounts for 14.81%.\n\n**Insights:**\n\n* The high percentage of Core products suggests that NORTHAM's Direct Named Pipeline is driven by traditional cloud offerings.\n* The relatively low percentage of AI products indicates that there may be an opportunity to increase adoption of AI solutions in the region.\n* The dominance of Google Cloud Platform suggests that NORTHAM's Direct Named Pipeline is focused on enterprise-level solutions.\n* The presence of Google Workspace indicates that there is also a market for productivity and collaboration tools in the region.","```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv(\"data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv\")\n\n# Filter data for EMEA and NORTHAM regions\nemea_df = df[df[\"Grouped_Marketing_Target_Region\"] == \"EMEA\"]\nnortham_df = df[df[\"Grouped_Marketing_Target_Region\"] == \"NORTHAM\"]\n\n# Calculate Partner Pipeline percentage for EMEA and NORTHAM\nemea_partner_pipeline_percentage = (emea_df[\"Opportunity_Target_Attainment_Source\"] == \"Partner\").sum() / len(emea_df) * 100\nnortham_partner_pipeline_percentage = (northam_df[\"Opportunity_Target_Attainment_Source\"] == \"Partner\").sum() / len(northam_df) * 100\n\n# Print the results\nprint(\"EMEA Partner Pipeline Percentage:\", round(emea_partner_pipeline_percentage, 3), \"%\")\nprint(\"NORTHAM Partner Pipeline Percentage:\", round(northam_partner_pipeline_percentage, 3), \"%\")\n```\n","```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv(\"data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv\")\n\n# Filter for LATAM and Startup segment\nlatam_startup = df[(df[\"Grouped_Marketing_Target_Region\"] == \"LATAM\") & (df[\"Account_Segment\"] == \"Startup\")]\n\n# Calculate product category percentages\nproduct_category_percentages = latam_startup[\"Product_Category\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"Product category percentages for LATAM's Startup segment:\")\nprint(product_category_percentages.to_string())\n```"]}

    slide_data_16 = {"data":{"NORTHAM":{"Direct Named QSOs":{"QTD":"0.5K","Attain":"32.0%","YoY":"32.0%"},"Direct Named Pipeline":{"QTD":"138.7M","Attain":"51.0%","YoY":"98.0%"},"Partner Pipeline":{"QTD":"128.2M","Attain":"37.0%","YoY":"-2.0%"},"GWS QSOs":{"QTD":"0.5K","Attain":"32.0%","YoY":"32.0%"},"GWS Direct + Partner Pipe":{"QTD":"266.9M","Attain":"43.0%","YoY":"33.0%"}},"LATAM":{"Direct Named QSOs":{"QTD":"0.2K","Attain":"28.0%","YoY":"-8.0%"},"Direct Named Pipeline":{"QTD":"36.2M","Attain":"41.0%","YoY":"80.0%"},"Partner Pipeline":{"QTD":"137.0M","Attain":"55.0%","YoY":"37.0%"},"GWS QSOs":{"QTD":"0.2K","Attain":"28.0%","YoY":"-8.0%"},"GWS Direct + Partner Pipe":{"QTD":"173.2M","Attain":"51.0%","YoY":"44.0%"}},"EMEA":{"Direct Named QSOs":{"QTD":"0.5K","Attain":"33.0%","YoY":"26.0%"},"Direct Named Pipeline":{"QTD":"97.1M","Attain":"36.0%","YoY":"31.0%"},"Partner Pipeline":{"QTD":"122.4M","Attain":"71.0%","YoY":"100.0%"},"GWS QSOs":{"QTD":"0.5K","Attain":"33.0%","YoY":"26.0%"},"GWS Direct + Partner Pipe":{"QTD":"219.6M","Attain":"50.0%","YoY":"62.0%"}},"JAPAC":{"Direct Named QSOs":{"QTD":"0.4K","Attain":"23.0%","YoY":"-14.0%"},"Direct Named Pipeline":{"QTD":"61.5M","Attain":"21.0%","YoY":"-14.0%"},"Partner Pipeline":{"QTD":"174.9M","Attain":"42.0%","YoY":"7.0%"},"GWS QSOs":{"QTD":"0.4K","Attain":"23.0%","YoY":"-14.0%"},"GWS Direct + Partner Pipe":{"QTD":"236.4M","Attain":"33.0%","YoY":"1.0%"}},"PUBLIC SECTOR":{"Direct Named QSOs":{"QTD":"0.1K","Attain":"43.0%","YoY":"52.0%"},"Direct Named Pipeline":{"QTD":"27.4M","Attain":"52.0%","YoY":"25.0%"},"Partner Pipeline":{"QTD":"3.3M","Attain":"17.0%","YoY":"-77.0%"},"GWS QSOs":{"QTD":"0.1K","Attain":"43.0%","YoY":"52.0%"},"GWS Direct + Partner Pipe":{"QTD":"30.7M","Attain":"42.0%","YoY":"-15.0%"}},"GLOBAL":{"Direct Named QSOs":{"QTD":"1.8K","Attain":"30.0%","YoY":"12.0%"},"Direct Named Pipeline":{"QTD":"360.9M","Attain":"37.0%","YoY":"40.0%"},"Partner Pipeline":{"QTD":"565.9M","Attain":"47.0%","YoY":"21.0%"},"GWS QSOs":{"QTD":"1.8K","Attain":"30.0%","YoY":"12.0%"},"GWS Direct + Partner Pipe":{"QTD":"926.8M","Attain":"42.0%","YoY":"27.0%"}}},"insights":["Globally, Partner-influenced pipeline surpasses Direct Named Pipeline, achieving 47% attainment compared to 37% for Direct. This highlights the increasing importance and effectiveness of partnerships in driving Google Cloud adoption.","Despite overall pipeline growth, QTD attainment lags across all regions, indicating potential bottlenecks in the early stages of the lead-to-opportunity lifecycle. This suggests a need to investigate and optimize lead qualification processes and resource allocation to improve conversion rates.","Public Sector's Partner Pipeline experiences a significant decline of 77% YoY, reaching only 17% attainment. This alarming trend necessitates a thorough reassessment and revamp of partner engagement strategies within the sector to reverse the decline and capitalize on the growth potential."],"recommendations":["Scale EMEA's partner engagement strategies globally to leverage their success in driving Partner Pipeline growth. Sharing best practices and providing tailored support to other regions can help replicate their achievements.","Conduct a comprehensive analysis of JAPAC's demand generation challenges, including market-specific factors, campaign effectiveness, and competitive landscape. Based on the findings, develop and implement targeted interventions to address the negative growth and low attainment rates.","Prioritize Public Sector partnerships by investing in specialized resources, training, and incentives to reactivate engagement and pipeline generation. Collaboration with partner success managers and targeted outreach programs can help rebuild momentum."],"drivers":["Partner-influenced pipeline makes up 52.8% of the total pipeline, with Direct Named Pipeline representing the remaining 47.2%. \n","Google Workspace represents 100% of the Opportunity Product Family, while Core Products make up 72.23% and Gemini for Workspace Products represent 27.73% of the Opportunity Product Category. \n","Google Workspace represents 100% of the Opportunity Product Families within the Core Product Category, comprising the entire Public Sector Partner Pipeline. \n"],"codes":["```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv(\"data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv\")\n\n# Calculate the percentage of opportunities influenced by partners\npartner_percentage = df[df[\"Opportunity_Target_Attainment_Source\"] == \"Partner\"][\"Opportunity_Product_Family\"].count() / df[\"Opportunity_Product_Family\"].count()\n\n# Calculate the percentage of opportunities influenced directly\ndirect_percentage = df[df[\"Opportunity_Target_Attainment_Source\"] == \"Direct\"][\"Opportunity_Product_Family\"].count() / df[\"Opportunity_Product_Family\"].count()\n\n# Print the results\nprint(f\"Partner-influenced pipeline contributes {partner_percentage:.3f} of the total pipeline, while Direct Named Pipeline contributes {direct_percentage:.3f}.\")\n```","```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv(\"data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv\")\n\n# Calculate the percentage of opportunities for each product category\nproduct_category_percentages = df[\"Opportunity_Product_Category\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"Product Category Percentages:\")\nprint(product_category_percentages.to_string())\n\n# Calculate the percentage of opportunities for each product family\nproduct_family_percentages = df[\"Opportunity_Product_Family\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"\\nProduct Family Percentages:\")\nprint(product_family_percentages.to_string())\n```","```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv(\"data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv\")\n\n# Filter for Public Sector and Partner opportunities\ndf_public_sector_partner = df[(df[\"Account_Sector\"] == \"Public Sector\") & (df[\"Opportunity_Target_Attainment_Source\"] == \"Partner\")]\n\n# Calculate the percentage of opportunities for each product category\nproduct_category_percentages = df_public_sector_partner[\"Product_Category\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"Product Category Percentage in Public Sector Partner Pipeline:\")\nprint(product_category_percentages.to_string())\n\n# Calculate the percentage of opportunities for each product family\nproduct_family_percentages = df_public_sector_partner[\"Opportunity_Product_Family\"].value_counts(normalize=True) * 100\n\n# Print the results\nprint(\"\\nProduct Family Percentage in Public Sector Partner Pipeline:\")\nprint(product_family_percentages.to_string())\n```"]}

    slide_data_17 = {
        "data": {
            "NORTHAM": {
                "Direct Named QSOs": {"QTD": "0.5K", "Attain": "32.0%", "YoY": "32.0%"},
                "Direct Named Pipeline": {
                    "QTD": "138.7M",
                    "Attain": "51.0%",
                    "YoY": "98.0%",
                },
                "Partner Pipeline": {
                    "QTD": "128.2M",
                    "Attain": "37.0%",
                    "YoY": "-2.0%",
                },
                "GWS QSOs": {"QTD": "0.5K", "Attain": "32.0%", "YoY": "32.0%"},
                "GWS Direct + Partner Pipe": {
                    "QTD": "266.9M",
                    "Attain": "43.0%",
                    "YoY": "33.0%",
                },
            },
            "LATAM": {
                "Direct Named QSOs": {"QTD": "0.2K", "Attain": "28.0%", "YoY": "-8.0%"},
                "Direct Named Pipeline": {
                    "QTD": "36.2M",
                    "Attain": "41.0%",
                    "YoY": "80.0%",
                },
                "Partner Pipeline": {
                    "QTD": "137.0M",
                    "Attain": "55.0%",
                    "YoY": "37.0%"},
                "GWS QSOs": {"QTD": "0.2K", "Attain": "28.0%", "YoY": "-8.0%"},
                "GWS Direct + Partner Pipe": {
                    "QTD": "173.2M",
                    "Attain": "51.0%",
                    "YoY": "44.0%",
                },
            },
            "EMEA": {
                "Direct Named QSOs": {"QTD": "0.5K", "Attain": "33.0%", "YoY": "26.0%"},
                "Direct Named Pipeline": {
                    "QTD": "97.1M",
                    "Attain": "36.0%",
                    "YoY": "31.0%",
                },
                "Partner Pipeline": {
                    "QTD": "122.4M",
                    "Attain": "71.0%",
                    "YoY": "100.0%",
                },
                "GWS QSOs": {"QTD": "0.5K", "Attain": "33.0%", "YoY": "26.0%"},
                "GWS Direct + Partner Pipe": {
                    "QTD": "219.6M",
                    "Attain": "50.0%",
                    "YoY": "62.0%",
                },
            },
            "JAPAC": {
                "Direct Named QSOs": {
                    "QTD": "0.4K",
                    "Attain": "23.0%",
                    "YoY": "-14.0%",
                },
                "Direct Named Pipeline": {
                    "QTD": "61.5M",
                    "Attain": "21.0%",
                    "YoY": "-14.0%",
                },
                "Partner Pipeline": {"QTD": "174.9M", "Attain": "42.0%", "YoY": "7.0%"},
                "GWS QSOs": {"QTD": "0.4K", "Attain": "23.0%", "YoY": "-14.0%"},
                "GWS Direct + Partner Pipe": {
                    "QTD": "236.4M",
                    "Attain": "33.0%",
                    "YoY": "1.0%",
                },
            },
            "PUBLIC SECTOR": {
                "Direct Named QSOs": {"QTD": "0.1K", "Attain": "43.0%", "YoY": "52.0%"},
                "Direct Named Pipeline": {
                    "QTD": "27.4M",
                    "Attain": "52.0%",
                    "YoY": "25.0%",
                },
                "Partner Pipeline": {"QTD": "3.3M", "Attain": "17.0%", "YoY": "-77.0%"},
                "GWS QSOs": {"QTD": "0.1K", "Attain": "43.0%", "YoY": "52.0%"},
                "GWS Direct + Partner Pipe": {
                    "QTD": "30.7M",
                    "Attain": "42.0%",
                    "YoY": "-15.0%",
                },
            },
            "GLOBAL": {
                "Direct Named QSOs": {"QTD": "1.8K", "Attain": "30.0%", "YoY": "12.0%"},
                "Direct Named Pipeline": {
                    "QTD": "360.9M",
                    "Attain": "37.0%",
                    "YoY": "40.0%",
                },
                "Partner Pipeline": {
                    "QTD": "565.9M",
                    "Attain": "47.0%",
                    "YoY": "21.0%",
                },
                "GWS QSOs": {"QTD": "1.8K", "Attain": "30.0%", "YoY": "12.0%"},
                "GWS Direct + Partner Pipe": {
                    "QTD": "926.8M",
                    "Attain": "42.0%",
                    "YoY": "27.0%",
                },
            },
        },
        "insights": [
            "**GWS EMEA demonstrates exceptional performance in leveraging partner channels, achieving 71% Partner Pipeline attainment and a remarkable 100% YoY growth, signifying the effectiveness of their partnership strategy.**",
            "**GWS Public Sector region faces a critical challenge with a 77% YoY drop in Partner Pipeline, indicating a need for immediate action to address potential issues within the partner ecosystem and ensure alignment between GWS and partners.**",
            "**GWS JAPAC lags behind other regions with the lowest attainment of Direct Named QSOs (23%) and Direct Named Pipeline (21%), coupled with a minimal 1% YoY growth in GWS Direct + Partner Pipe, suggesting systemic challenges in their demand generation approach across both direct and partner channels.**",
        ],
        "recommendations": [
            "**Analyze and replicate EMEA's partner success strategies in other regions, focusing on partner enablement, joint marketing initiatives, and incentives to drive partner-led demand generation.**",
            "**Conduct a thorough review of the Public Sector partner program, identifying and addressing any bottlenecks hindering performance. This includes evaluating partner engagement, pipeline support, and alignment with GWS Public Sector's go-to-market strategy.**",
            "**Develop a comprehensive action plan to improve JAPAC's demand generation efforts, including reassessing target audience segmentation, refining messaging and value proposition for the JAPAC market, and exploring alternative lead generation channels and tactics.**",
        ],
        "drivers": [
            "Nearly half (47.162%) of GWS EMEA's opportunities stem from partners, underscoring the significance of their contributions. \n",
            "The GWS Public Sector Partner Pipeline is entirely composed of Core products (100%). \n",
            "",
        ],
    }

    # Slide number for each data set
    slide_number_11 = 11
    slide_number_14 = 14
    slide_number_15 = 15
    slide_number_16 = 16
    slide_number_17 = 17

    file_id = "1"

    # Create presentation for Slide 11
    output_path_11 = create_presentation(
        slide_data_11, file_id + "_11", slide_number_11
    )
    logger.info(f"Presentation created at {output_path_11}")

    # Create presentation for Slide 14
    output_path_14 = create_presentation(
        slide_data_14, file_id + "_14", slide_number_14
    )
    logger.info(f"Presentation created at {output_path_14}")

    # Create presentation for Slide 15
    output_path_15 = create_presentation(
        slide_data_15, file_id + "_15", slide_number_15
    )
    logger.info(f"Presentation created at {output_path_15}")

    # Create presentation for Slide 16
    output_path_16 = create_presentation(
        slide_data_16, file_id + "_16", slide_number_16
    )
    logger.info(f"Presentation created at {output_path_16}")

    output_path_17 = create_presentation(
        slide_data_17, file_id + "_17", slide_number_17
    )
    logger.info(f"Presentation created at {output_path_17}")
