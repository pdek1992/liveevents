import os
import time
import pandas as pd
import ctypes
from PIL import Image, ImageDraw, ImageFont, ImageFilter
import textwrap
from datetime import datetime, timedelta
import re
import tkinter as tk
import logging
import logging.handlers
import sys

def resource_path(relative_path):
    """ Get absolute path to resource, works for dev and PyInstaller """
    if hasattr(sys, '_MEIPASS'):
        return os.path.join(sys._MEIPASS, relative_path)
    return os.path.join(os.path.abspath("."), relative_path)
#log_file = resource_path("app.log")
log_file = ".\live.log"

# Configure logging
logger = logging.getLogger("MyAppLogger")
logger.setLevel(logging.DEBUG)  # Log everything (DEBUG, INFO, WARNING, ERROR, CRITICAL)

# Create a rotating file handler (auto-rotates after 48 hours)
handler = logging.handlers.TimedRotatingFileHandler(
    log_file, when="h", interval=48, backupCount=5, encoding="utf-8"
)
handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))

# Add the handler to the logger
logger.addHandler(handler)

# Enable console logging (optional)
console_handler = logging.StreamHandler()
console_handler.setFormatter(logging.Formatter("%(asctime)s - %(levelname)s - %(message)s"))
logger.addHandler(console_handler)


    
def get_resource_path(filename):
    return os.path.join(os.getcwd(), filename)

def read_config(file_path):
    config = {}
    with open(file_path, 'r') as file:
        for line in file:
            line = line.split('#', 1)[0].strip()  # Remove comments and strip spaces
            if not line:  # Skip empty lines
                continue
            if '=' in line:
                key, value = map(str.strip, line.split('=', 1))  # Split and strip spaces/tabs
                config[key] = value
    return config
    

# Read configuration from external file
config = read_config('values.txt')

# Extract values from config
grace_period = int(config.get('grace_period', 5))  
running_color = config.get('running', 'red')
finished_color =  config.get('finished', 'white')
yet_to_start_color =  config.get('yet_to_start', 'grey')
skip_rows =  int(config.get('skip_rows', 2))
run_frequency =  int(config.get('run_frequency', 60))
display_rows = int(config.get('display_rows', 8))

def monitor_directory(directory):
    #last_processed_file = None

    while True:
        try:
            files = [f for f in os.listdir(directory) if f.endswith((".xlsx", ".ods"))]
            if not files:
                #print("No Excel or ODS files found. Waiting...")
                logger.error("No Excel or ODS files found. Waiting...")
            else:
                latest_file = max(files, key=lambda f: os.path.getmtime(os.path.join(directory, f)))
                file_path = os.path.join(directory, latest_file)
                logger.info(f"using {latest_file} from {file_path}")

                
                logger.info(f"Processing: {latest_file}")
                process_file(file_path)  # Process the file

                image_path = os.path.join(directory, "output_image.png")
                set_as_wallpaper(image_path)  # Set the wallpaper

                #last_processed_file = latest_file  # Mark as processed

        except Exception as e:
            logger.error(f"Error: {e}. Retrying in {grace_period} seconds...")

        time.sleep(run_frequency)  # Retry after 30 seconds

def process_file(file_path):
    required_columns = {
        "IST(+ 5.5)": "IST(+ 5.5)",
        "DUR": "DUR",
        "TELECAST": "TELECAST",
        "DESCRIPTION": "DESCRIPTION",
        "CHANNEL": "CHANNEL",
        "LINE INPUT": "LINE INPUT",
        "SOURCE": "SOURCE",
        "CIRCUIT": "CIRCUIT"
    }

    df = pd.read_excel(file_path, sheet_name="PLANNER", skiprows=2)

    if "CHANNEL" in df.columns:
        channel_index = df.columns.get_loc("CHANNEL")
        potential_split_columns = [col for col in df.columns[channel_index + 1:] if "Unnamed:" in str(col)]

        if potential_split_columns:
            for split_col in potential_split_columns:
                df["CHANNEL"] = df.apply(
                    lambda row: str(row["CHANNEL"]) + " " + str(row[split_col]) if pd.notna(row[split_col]) else str(row["CHANNEL"]),
                    axis=1
                )
                df = df.drop(columns=[split_col])

    selected_columns = [col for col in required_columns.keys() if col in df.columns]
    df = df[selected_columns]

    if "TELECAST" in df.columns:
        df = df[df["TELECAST"].str.contains("live", case=False, na=False)]

    today = datetime.now().date()
    df['IST(+ 5.5)'] = pd.to_datetime(df['IST(+ 5.5)'], errors='coerce')
    df = df[df['IST(+ 5.5)'].dt.date == today]

    df = df.sort_values(by='IST(+ 5.5)')
   
    # Filter out finished events
    now = datetime.now()
    df['End Time'] = df.apply(
    lambda row: row['IST(+ 5.5)'] + timedelta(hours=int(row['DUR']), minutes=(row['DUR'] % 1) * 60) + timedelta(minutes=grace_period)
    if pd.notna(row['IST(+ 5.5)']) and pd.notna(row['DUR'])
    else None, axis=1)
    
    df = df[df['End Time'] > now]  # Keep only upcoming or ongoing events
    df = df.drop(columns=['End Time'])  # Remove temporary column

    # Replace NaN with empty string in all columns
    
    df.fillna('', inplace=True)
    

    #print(df.head())
    #print("done")
    create_image(df)
    return df

def wrap_text(draw, text, font, max_width):
    words = text.split()
    lines = []
    current_line = ""
    for word in words:
        bbox = draw.textbbox((0, 0), current_line + " " + word, font=font)
        if bbox[2] <= max_width:
            current_line += (" " if current_line else "") + word
        else:
            lines.append(current_line)
            current_line = word
    lines.append(current_line)
    return lines

def create_image(df):
    background_path = get_resource_path('background.jfif')
    logo_path = get_resource_path('logo.jfif')

    if not os.path.exists(background_path):
        background_path = 'default_background.jfif'
    if not os.path.exists(logo_path):
        logo_path = 'default_logo.jfif'

    background = Image.open(background_path).convert('RGBA')
    logo = Image.open(logo_path).convert('RGBA')

    root = tk.Tk()
    screen_width = root.winfo_screenwidth()
    screen_height = root.winfo_screenheight()
    logger.info(f"{screen_width}X{screen_height}")
    root.destroy()

    background = background.resize((screen_width, screen_height)).filter(ImageFilter.GaussianBlur(5))
    image = Image.new('RGBA', background.size)
    image.paste(background, (0, 0))

    logo_size = (int(image.width * 0.05), int(image.width * 0.05))
    logo = logo.resize(logo_size)
    image.paste(logo, (20, 20), logo)
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(get_resource_path("arialbd.ttf"), 24)
    y_start, x_start, row_height = 0, 0, 50
    column_widths = [int(screen_width * 0.15625), int(screen_width * 0.05208), int(screen_width * 0.1302), int(screen_width * 0.24218), int(screen_width * 0.10416), int(screen_width * 0.10416), int(screen_width * 0.10416), int(screen_width * 0.10416)]
    headers = ["IST(+ 5.5)", "DUR", "TELECAST", "DESCRIPTION", "CHANNEL", "LINE INPUT", "SOURCE", "CIRCUIT"]
    logger.info(f"using headers {headers}")

    for i, header in enumerate(headers):
        x_position = x_start + sum(column_widths[:i])
        draw.rectangle([(x_position, y_start), (x_position + column_widths[i], y_start + row_height)], fill="black")
        draw.text((x_position + 10, y_start + 10), header, font=font, fill="white")

    y_position = y_start + row_height
    now = datetime.now()

    for row_idx, row in df.head(display_rows).iterrows():
        max_lines = 1
        wrapped_texts = []
        for col_idx, header in enumerate(headers):
            cell_text = str(row[header])
            wrapped_text = wrap_text(draw, cell_text, font, column_widths[col_idx] - 20)
            wrapped_texts.append(wrapped_text)
            max_lines = max(max_lines, len(wrapped_text))

        is_running = False
        yet_to_start = False
        finished = False
        start_time = None
        duration_str = None
        now = datetime.now()
        if "IST(+ 5.5)" in row and "DUR" in row:

            start_time = row["IST(+ 5.5)"]
            duration_str = row["DUR"]

        if isinstance(duration_str, (float, int)) and isinstance(start_time, datetime):
            duration = timedelta(hours=duration_str) 
            end_time = start_time + duration
            if start_time <= now <= end_time:
                is_running = True
            elif now < start_time:
                yet_to_start = True
            elif end_time < now:
                finished = True
           

        for col_idx, header in enumerate(headers):
            x_position = x_start + sum(column_widths[:col_idx])
            fill_color = "#90EE90"
            if is_running:
                fill_color = running_color
            if yet_to_start:
                fill_color = yet_to_start_color
            if finished:
                fill_color = finished_color
            
            

            draw.rectangle([(x_position, y_position), (x_position + column_widths[col_idx], y_position + row_height * max_lines)],
                            outline="black", fill=fill_color)

            for line_num, line in enumerate(wrapped_texts[col_idx]):
                left, top, right, bottom = draw.textbbox((x_position, y_position), line, font=font)
                text_width = right - left
                text_height = bottom - top
                x_text = x_position + (column_widths[col_idx] - text_width) // 2
                y_text = y_position + (row_height * max_lines - text_height) // 2 + line_num * 30
                draw.text((x_text, y_text), line, font=font, fill="black")

        y_position += row_height * max_lines

    image = image.convert('RGB')
    image.save(os.path.join(os.getcwd(), 'output_image.png'))
    logger.info("Image created: output_image.png")

def set_as_wallpaper(image_path):
    ctypes.windll.user32.SystemParametersInfoW(20, 0, image_path, 0)
    logger.info("Wallpaper set")

if __name__ == "__main__":
    
    #current_dir = os.path.dirname(os.path.abspath(__file__))
    #print(current_dir)
    current_dir = os.getcwd()
    monitor_directory(current_dir)