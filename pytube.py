# -*- coding: utf-8 -*-
"""
Created on Tue Jun 20 00:05:11 2023

@author: Win10
"""

import pandas as pd
from pytube import YouTube

def download_video(url, output_path):
    try:
        yt = YouTube(url)
        video = yt.streams.get_highest_resolution()
        video.download(output_path=output_path)
        print(f"Video '{yt.title}' downloaded successfully.")
    except Exception as e:
        print(f"Error occurred while downloading '{url}': {str(e)}")


def download_audio(url, output_path):
    try:
        yt = YouTube(url)
        audio = yt.streams.filter(only_audio=True).first()
        audio.download(output_path=output_path, filename_prefix="")
        print(f"Audio '{yt.title}' downloaded successfully.")
    except Exception as e:
        print(f"Error occurred while downloading audio from '{url}': {str(e)}")

# Function to import URLs from an Excel file
def import_urls_from_excel(file_path, sheet_name, column_name):
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        urls = df[column_name].tolist()
        return urls
    except Exception as e:
        print(f"Error occurred while importing URLs: {str(e)}")
        return []

# Example usage
path_base = "C://Users//Win10//Documents//Python//"
excel_file_path = path_base+ "links_for_python.xlsx"
sheet_name = "Sheet1"  # Replace with your sheet name
column_name = "URL"  # Replace with your column name

output_directory = path_base+"output_pytube"  # Replace with your desired directory path

# Import URLs from Excel
urls = import_urls_from_excel(excel_file_path, sheet_name, column_name)

# Download audio from the URLs
for url in urls:
    download_audio(url, output_directory)

