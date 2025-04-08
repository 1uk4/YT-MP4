import os
import pandas as pd
import tkinter as tk
from .metadata_utils import add_metadata
import time
import yt_dlp

def safe_str(value):
    """Convert any value to string safely"""
    if pd.isna(value):  # Check for NaN/empty values
        return ''
    return str(value).strip()

def sanitize_filename(filename):
    """Remove or replace invalid characters for filenames"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename

def get_safe_filepath(title, artist, playlist_folder):
    """Create filename in format 'Song Name (Artist).mp3'"""
    safe_title = sanitize_filename(title)
    safe_artist = sanitize_filename(artist)
    filename = f"{safe_title} ({safe_artist}).mp3"
    return os.path.join(playlist_folder, filename)

def download_with_ytdlp(url, output_path, filename):
    """Download using yt-dlp"""
    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': os.path.join(output_path, filename),
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'quiet': True,
        'no_warnings': True
    }
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        return True, None
    except Exception as e:
        return False, str(e)

def download_music(excel_path, download_folder, status_text, progress_bar, progress_text, root, add_metadata_func):
    try:
        # Read the Excel file
        status_text.insert(tk.END, "Reading Excel file...\n")
        df = pd.read_excel(excel_path)
        status_text.insert(tk.END, f"Found {len(df)} rows\n")
        status_text.insert(tk.END, f"Columns found: {', '.join(df.columns)}\n")
        
        total_songs = len(df)
        
        for index, row in df.iterrows():
            # Safely convert values to strings
            title = safe_str(row.get('Song Name', ''))
            artist = safe_str(row.get('Artist', ''))
            url = safe_str(row.get('YT Link', ''))
            playlist = safe_str(row.get('Playlist', 'Default'))
            
            if not url or not title or not artist:
                status_text.insert(tk.END, f"❌ Skipping row {index + 1}: Missing required information\n")
                continue

            try:
                # Create playlist folder if it doesn't exist
                playlist_folder = os.path.join(download_folder, sanitize_filename(playlist))
                os.makedirs(playlist_folder, exist_ok=True)

                # Check if file already exists
                expected_file_path = get_safe_filepath(title, artist, playlist_folder)
                if os.path.exists(expected_file_path):
                    status_text.insert(tk.END, f"⏭️ Skipping '{title} ({artist})' - Already exists\n")
                    
                    # Update progress for skipped file
                    progress = (index + 1) / total_songs * 100
                    progress_bar['value'] = progress
                    progress_text.config(text=f"{progress:.1f}%")
                    root.update()
                    continue

                # Update status
                status_text.insert(tk.END, f"\nDownloading: {title} ({artist})\n")
                root.update()
                
                # Download using yt-dlp
                temp_filename = f"temp_{sanitize_filename(title)}"
                success, error = download_with_ytdlp(url, playlist_folder, temp_filename)

                if not success:
                    status_text.insert(tk.END, f"❌ Download failed for '{title} ({artist})': {error}\n")
                    continue

                # Get the downloaded file (yt-dlp adds .mp3 extension)
                temp_file = os.path.join(playlist_folder, temp_filename + '.mp3')
                
                if os.path.exists(temp_file):
                    # Rename to final filename
                    os.rename(temp_file, expected_file_path)
                    
                    # Add metadata
                    add_metadata_func(expected_file_path, title, artist)
                    
                    # Update progress
                    progress = (index + 1) / total_songs * 100
                    progress_bar['value'] = progress
                    progress_text.config(text=f"{progress:.1f}%")
                    root.update()
                    
                    status_text.insert(tk.END, f"✅ Successfully downloaded: {title} ({artist})\n")
                else:
                    status_text.insert(tk.END, f"❌ Failed to download '{title} ({artist})'\n")
                
            except Exception as e:
                status_text.insert(tk.END, f"❌ Error processing '{title} ({artist})': {str(e)}\n")
                # Cleanup any partial downloads
                if os.path.exists(expected_file_path):
                    try:
                        os.remove(expected_file_path)
                    except:
                        pass
                continue
            
        return True
        
    except pd.errors.EmptyDataError:
        status_text.insert(tk.END, "❌ Error: Excel file is empty\n")
        return False
    except Exception as e:
        status_text.insert(tk.END, f"❌ Error reading Excel file: {str(e)}\n")
        return False