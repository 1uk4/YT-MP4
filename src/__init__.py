import os
import pandas as pd
import yt_dlp
import subprocess
from tqdm import tqdm

# Get absolute path
script_dir = os.path.dirname(os.path.abspath(__file__))
file_path = os.path.join(script_dir, "data.xlsx")


# Check if file exists
if not os.path.exists(file_path):
    print(f"Error: File not found at {file_path}")
    exit(1)

sheets_dict = pd.read_excel(file_path, sheet_name=None)

# Parent directory for downloaded playlists
output_directory = "Downloaded_Playlists"
os.makedirs(output_directory, exist_ok=True)

# Download progress hook for tqdm
def progress_hook(d):
    if d['status'] == 'downloading':
        pbar.update(int(d.get('downloaded_bytes', 0)) - pbar.n)


# Function to add metadata using FFmpeg
def add_metadata(file_path, title, artist):
    temp_file = file_path.replace(".mp4", "_temp.mp4")
    command = [
        "ffmpeg",
        "-i", file_path,
        "-metadata", f"title={title}",
        "-metadata", f"artist={artist}",
        "-codec", "copy",
        temp_file
    ]
    
    subprocess.run(command, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
    
    # Replace original file with the one containing metadata
    if os.path.exists(temp_file):
        os.replace(temp_file, file_path)



# Iterate through each sheet (playlist)
for playlist_name, df in sheets_dict.items():
    # Create a folder for the playlist
    playlist_path = os.path.join(output_directory, playlist_name)
    os.makedirs(playlist_path, exist_ok=True)
    # Iterate through each row in the sheet

    with tqdm(total=len(df), desc=f"Downloading {playlist_name}", unit="song") as playlist_bar:
        for _, row in df.iterrows():
            song_name = str(row["Song Name"]).strip()
            artist = str(row["Artist"]).strip()
            youtube_url = str(row["YT Link"]).strip()


            if pd.isna(youtube_url) or youtube_url == "":
                ##print(f"Skipping {song_name} by {artist} (No URL provided)")
                continue

            # Define the output filename format
            output_filename = f"{song_name} ({artist})"
            output_path = os.path.join(playlist_path, output_filename)

            # Check if file already exists
            if os.path.exists(output_path):
                ##print(f"Skipping {output_filename}, already downloaded.")
                playlist_bar.update(1)
                continue

            # yt-dlp download options
            with tqdm(desc=f"Downloading {song_name} - {artist}", unit="B", unit_scale=True, leave=False) as pbar:
                ydl_opts = {
                    "format": "bestaudio/best",  # Best quality audio
                    "outtmpl": output_path,  # Save file in the playlist folder
                    "progress_hooks": [progress_hook],  # Attach progress hook
                    "quiet": True,             # Suppress most logs
                    "no-warnings": True,       # Suppress warnings
                    "cookies-from-browser": "chrome",
                    "postprocessors": [{
                        "key": "FFmpegVideoConvertor",  # Ensure it's saved as mp4
                        "preferedformat": "mp4",
                    }],
                }


                # Download the song using yt-dlp
                try:
                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        ydl.download([youtube_url])

                    
                    # Add metadata to the downloaded file
                    add_metadata(output_path, song_name, artist)
                    ##print(f"\n✅ Downloaded and added metadata: {output_filename}")

                except Exception as e:
                    print(f"\n❌ Error downloading {song_name} by {artist}: {e}")
                
            playlist_bar.update(1)

print("Download process completed.")









