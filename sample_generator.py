import csv
import warnings
from PIL import ImageFont, ImageDraw, Image
import pickle
import os


def generate_certificate(text_details):
    try:
        csv_file = rectangles['csv_path']
        template_path = rectangles['template_path']
        output_folder = "Samples"
        font_folder = "Fonts"

        with open(csv_file, 'r') as csvfile:
            reader = csv.DictReader(csvfile)
            row = next(reader)
    
            with Image.open(template_path) as image:
                draw = ImageDraw.Draw(image)
                name_value = row.get("Name", "")
                output_filename = f"Sample_Doc.png"

                for col_name, text in row.items():
                        if col_name in text_details:
                            details = text_details[col_name]
                            start_x, start_y, end_x, end_y = details['start_x'], details['start_y'], details['end_x'], details['end_y']
                            font_style, font_size, thickness = details['font_style'], details['font_size'], details['thickness']

                            rectangle_width = end_x - start_x
                            rectangle_height = end_y - start_y

                            if font_size == "Auto":
                                font_size = min(int(rectangle_height * 0.8), int(rectangle_width * 0.8))
                            else:
                                font_size = int(font_size)

                            text_position = ((start_x + end_x) / 2, (start_y + end_y) / 2)

                            font_file = os.path.join(font_folder, font_style)
                            font = ImageFont.truetype(font_file, size=font_size)

                            with warnings.catch_warnings():
                                warnings.simplefilter("ignore")
                                # Calculate text size using textbbox
                                bbox = draw.textbbox((0, 0), text, font=font)
                                text_width = bbox[2] - bbox[0]
                                text_height = bbox[3] - bbox[1]

                            centered_x = text_position[0] - text_width / 2
                            centered_y = text_position[1] - text_height / 2

                            draw.text((centered_x, centered_y), text, font=font, fill=(0, 0, 0), stroke_width=thickness)

                image.save(f'{output_folder}/{output_filename}')
                print(f"Generated certificate for {output_filename}")

    except Exception as e:
        print(f"Error generating document: {e}")


    except Exception as e:
        print(f"Error generating Document: {e}")


def load_rectangles(file_path):
    try:
        with open(file_path, 'rb') as file:
            rectangles = pickle.load(file)
        return rectangles
    except FileNotFoundError:
        print("Error: File not found.")
    except Exception as e:
        print(f"Error loading Data: {e}")
        return {}


rectangles = load_rectangles("AppData/rectangles.pkl")
if rectangles:
    generate_certificate(rectangles)
else:
    print("No Template Data Found")
