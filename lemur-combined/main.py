import logging
from pptx import Presentation

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


def populate_slide(slide, data, slide_number):
    """
    Populate a slide with the given content.
    """
    try:
        # Assuming the table structure is known and fixed
        table = None
        for shape in slide.shapes:
            if shape.has_table:
                table = shape.table
                break

        if table:
            logger.info(
                f"Table dimensions: {len(table.rows)} rows x {len(table.columns)} columns"
            )

            if slide_number == 11:
                # Existing logic for Slide 11
                regions = ["NORTHAM", "EMEA", "JAPAC", "LATAM"]
                metrics = [
                    ("Ent+Corp Pipeline", "QTD", "Attain"),
                    ("SMB Pipeline", "QTD", "Attain"),
                    ("Total Partner Marketing Sourced", "QTD", "Attain"),
                ]

                gcp_start_row = 3
                gws_start_row = 7

                for i, (metric, qtd_key, attain_key) in enumerate(metrics):
                    for j, region in enumerate(regions):
                        gcp_data = (
                            data["GCP"]
                            .get(region, {})
                            .get(metric, {"QTD": "", "Attain": ""})
                        )
                        gcp_value = f"{gcp_data[qtd_key]} ({gcp_data[attain_key]})"
                        table.cell(gcp_start_row + i, 1 + j).text = gcp_value
                        logger.info(
                            f"Populated GCP {metric} for {region} with {gcp_value}"
                        )

                        gws_data = (
                            data["GWS"]
                            .get(region, {})
                            .get(metric, {"QTD": "", "Attain": ""})
                        )
                        gws_value = f"{gws_data[qtd_key]} ({gws_data[attain_key]})"
                        table.cell(gws_start_row + i, 1 + j).text = gws_value
                        logger.info(
                            f"Populated GWS {metric} for {region} with {gws_value}"
                        )

            elif slide_number == 14:
                # Logic for Slide 14
                regions = [
                    "NORTHAM",
                    "LATAM",
                    "EMEA",
                    "JAPAC",
                    "PUBLIC SECTOR",
                    "GLOBAL",
                ]
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
                        table.cell(start_row + i, 1 + (j * 2)).text = (
                            region_value_qtd_attain
                        )
                        # YoY
                        table.cell(start_row + i, 2 + (j * 2)).text = region_value_yoy

                        logger.info(
                            f"Populated {metric} for {region} with QTD+Attain: {region_value_qtd_attain} and YoY: {region_value_yoy}"
                        )

            elif slide_number == 15:
                # Logic for Slide 15
                regions = [
                    "NORTHAM",
                    "LATAM",
                    "EMEA",
                    "JAPAC",
                    "PUBLIC SECTOR",
                    "GLOBAL",
                ]
                metrics = [
                    ("Direct Named QSOs", "QTD", "Attain", "YoY"),
                    ("Direct Named Pipeline", "QTD", "Attain", "YoY"),
                    ("Partner Pipeline", "QTD", "Attain", "YoY"),
                    ("GCP Direct QSOs", "QTD", "Attain", "YoY"),
                    ("GCP Direct + Partner Pipe", "QTD", "Attain", "YoY"),
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
                        table.cell(start_row + i, 1 + (j * 2)).text = (
                            region_value_qtd_attain
                        )
                        # YoY
                        table.cell(start_row + i, 2 + (j * 2)).text = region_value_yoy

                        logger.info(
                            f"Populated {metric} for {region} with QTD+Attain: {region_value_qtd_attain} and YoY: {region_value_yoy}"
                        )

            elif slide_number == 16:
                # Logic for Slide 16
                regions = [
                    "NORTHAM",
                    "LATAM",
                    "EMEA",
                    "JAPAC",
                    "PUBLIC SECTOR",
                    "GLOBAL",
                ]
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
                        region_data = data.get(region, {}).get(
                            metric, {"QTD": "", "Attain": "", "YoY": ""}
                        )
                        region_value_qtd_attain = (
                            f"{region_data[qtd_key]} ({region_data[attain_key]})"
                        )
                        region_value_yoy = region_data[yoy_key]

                        # QTD and Attain
                        table.cell(start_row + i, 1 + (j * 2)).text = (
                            region_value_qtd_attain
                        )
                        # YoY
                        table.cell(start_row + i, 2 + (j * 2)).text = region_value_yoy

                        logger.info(
                            f"Populated {metric} for {region} with QTD+Attain: {region_value_qtd_attain} and YoY: {region_value_yoy}"
                        )

            elif slide_number == 17:
                    # Logic for Slide 17
                    regions = ["NORTHAM", "LATAM", "EMEA", "JAPAC", "PUBLIC SECTOR", "GLOBAL"]
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
                            region_data = data.get("data", {}).get(region, {}).get(metric, {"QTD": "", "Attain": "", "YoY": ""})
                            region_value_qtd_attain = f"{region_data[qtd_key]} ({region_data[attain_key]})"
                            region_value_yoy = region_data[yoy_key]

                            # QTD and Attain
                            table.cell(start_row + i, 1 + (j * 2)).text = region_value_qtd_attain
                            # YoY
                            table.cell(start_row + i, 2 + (j * 2)).text = region_value_yoy

                            logger.info(f"Populated {metric} for {region} with QTD+Attain: {region_value_qtd_attain} and YoY: {region_value_yoy}")


        else:
            logger.error("No table found in the slide")

    except Exception as e:
        logger.error(f"Error populating slide: {e}")


# Example usage
if __name__ == "__main__":
    # Hardcoded data for local testing
    slide_data_11 = {
        "GCP": {
            "NORTHAM": {
                "Ent+Corp Pipeline": {"QTD": "579.0M", "Attain": "32.0%"},
                "SMB Pipeline": {"QTD": "55.9M", "Attain": "49.0%"},
                "Total Partner Marketing Sourced": {"QTD": "634.9M", "Attain": "33.0%"},
            },
            "LATAM": {
                "Ent+Corp Pipeline": {"QTD": "152.6M", "Attain": "50.0%"},
                "SMB Pipeline": {"QTD": "32.3M", "Attain": "66.0%"},
                "Total Partner Marketing Sourced": {"QTD": "184.9M", "Attain": "53.0%"},
            },
            "EMEA": {
                "Ent+Corp Pipeline": {"QTD": "572.8M", "Attain": "71.0%"},
                "SMB Pipeline": {"QTD": "69.6M", "Attain": "48.0%"},
                "Total Partner Marketing Sourced": {"QTD": "642.4M", "Attain": "67.0%"},
            },
            "JAPAC": {
                "Ent+Corp Pipeline": {"QTD": "438.4M", "Attain": "56.0%"},
                "SMB Pipeline": {"QTD": "75.4M", "Attain": "58.0%"},
                "Total Partner Marketing Sourced": {"QTD": "513.8M", "Attain": "56.0%"},
            },
        },
        "GWS": {
            "NORTHAM": {
                "Ent+Corp Pipeline": {"QTD": "90.2M", "Attain": "30.0%"},
                "SMB Pipeline": {"QTD": "36.7M", "Attain": "69.0%"},
                "Total Partner Marketing Sourced": {"QTD": "126.9M", "Attain": "36.0%"},
            },
            "LATAM": {
                "Ent+Corp Pipeline": {"QTD": "79.5M", "Attain": "32.0%"},
                "SMB Pipeline": {"QTD": "60.7M", "Attain": "inf%"},
                "Total Partner Marketing Sourced": {"QTD": "140.2M", "Attain": "56.0%"},
            },
            "EMEA": {
                "Ent+Corp Pipeline": {"QTD": "82.7M", "Attain": "69.0%"},
                "SMB Pipeline": {"QTD": "40.7M", "Attain": "77.0%"},
                "Total Partner Marketing Sourced": {"QTD": "123.4M", "Attain": "71.0%"},
            },
            "JAPAC": {
                "Ent+Corp Pipeline": {"QTD": "138.5M", "Attain": "46.0%"},
                "SMB Pipeline": {"QTD": "37.3M", "Attain": "33.0%"},
                "Total Partner Marketing Sourced": {"QTD": "175.8M", "Attain": "42.0%"},
            },
        },
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

    slide_data_15 = {
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

    slide_data_16 = {
        "NORTHAM": {
            "Direct Named QSOs": {"QTD": "0.5K", "Attain": "32.0%", "YoY": "32.0%"},
            "Direct Named Pipeline": {
                "QTD": "138.7M",
                "Attain": "51.0%",
                "YoY": "98.0%",
            },
            "Partner Pipeline": {"QTD": "128.2M", "Attain": "37.0%", "YoY": "-2.0%"},
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
            "Partner Pipeline": {"QTD": "137.0M", "Attain": "55.0%", "YoY": "37.0%"},
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
            "Partner Pipeline": {"QTD": "122.4M", "Attain": "71.0%", "YoY": "100.0%"},
            "GWS QSOs": {"QTD": "0.5K", "Attain": "33.0%", "YoY": "26.0%"},
            "GWS Direct + Partner Pipe": {
                "QTD": "219.6M",
                "Attain": "50.0%",
                "YoY": "62.0%",
            },
        },
        "JAPAC": {
            "Direct Named QSOs": {"QTD": "0.4K", "Attain": "23.0%", "YoY": "-14.0%"},
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
            "Partner Pipeline": {"QTD": "565.9M", "Attain": "47.0%", "YoY": "21.0%"},
            "GWS QSOs": {"QTD": "1.8K", "Attain": "30.0%", "YoY": "12.0%"},
            "GWS Direct + Partner Pipe": {
                "QTD": "926.8M",
                "Attain": "42.0%",
                "YoY": "27.0%",
            },
        },
    }

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
                    "YoY": "37.0%",
                },
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
        "codes": [
            '```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv("data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv")\n\n# Filter data for GWS EMEA and Partner opportunities\ngws_emea_partner_df = df[(df["Opportunity_Product_Family"] == "Google Workspace") &\n                          (df["Grouped_Marketing_Target_Region"] == "EMEA") &\n                          (df["Opportunity_Target_Attainment_Source"] == "Partner")]\n\n# Calculate the percentage of Partner opportunities in GWS EMEA\npartner_percentage = (gws_emea_partner_df.shape[0] / df[(df["Opportunity_Product_Family"] == "Google Workspace") &\n                                                        (df["Grouped_Marketing_Target_Region"] == "EMEA")].shape[0]) * 100\n\n# Print the results\nprint("GWS EMEA Partner Pipeline Attainment: 71%")\nprint("GWS EMEA Partner Pipeline YoY Growth: 100%")\nprint("Percentage of Partner opportunities in GWS EMEA:", round(partner_percentage, 3), "%")\n```',
            '```python\nimport pandas as pd\n\n# Load the dataset\ndf = pd.read_csv("data/091597d1-894d-407c-82e5-bf24b6aaefe1.csv")\n\n# Filter for GWS Public Sector region and Partner Pipeline\ndf_gws_public_sector_partner = df[(df["Grouped_Marketing_Target_Region"] == "PUBLIC SECTOR") & (df["Opportunity_Target_Attainment_Source"] == "Partner")]\n\n# Calculate the percentage of each product category in the Partner Pipeline\nproduct_category_percentages = df_gws_public_sector_partner["Product_Category"].value_counts(normalize=True) * 100\n\n# Print the results\nprint("Product Category Percentages in GWS Public Sector Partner Pipeline:")\nprint(product_category_percentages.to_string())\n```\n## Product Category Percentages in GWS Public Sector Partner Pipeline:\n\nCore Products    75.000\nAI               25.000\nName: Product_Category, dtype: float64 \n\n**Insights:**\n\n* The majority (75%) of the Partner Pipeline in GWS Public Sector region consists of Core Products.\n* AI Products contribute to a smaller portion (25%) of the Partner Pipeline.\n\n**Recommendations:**\n\n* Investigate the reasons behind the decline in Partner Pipeline for AI Products.\n* Consider strategies to increase partner engagement and adoption of AI Products in the Public Sector region.\n* Evaluate the effectiveness of existing partner programs and incentives for GWS Public Sector.\n* Collaborate with partners to develop joint marketing and sales initiatives to drive growth in the Public Sector region.\n',
            None,
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
