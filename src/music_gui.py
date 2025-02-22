import os
import pandas as pd
import yt_dlp
import tkinter as tk
from tkinter import filedialog, messagebox
import threading
import glob
import subprocess


class MusicDownloaderApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Music Downloader")
        self.root.geometry("500x400")

        # File selection
        self.file_label = tk.Label(root, text="Select Excel File:")
        self.file_label.pack(pady=5)
        self.file_button = tk.Button(root, text="Choose File", command=self.select_file)
        self.file_button.pack(pady=5)
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(root, textvariable=self.file_path, width=50)
        self.file_entry.pack(pady=5)

        # Download folder selection
        self.folder_label = tk.Label(root, text="Select Download Folder:")
        self.folder_label.pack(pady=5)
        self.folder_button = tk.Button(root, text="Choose Folder", command=self.select_folder)
        self.folder_button.pack(pady=5)
        self.folder_path = tk.StringVar()
        self.folder_entry = tk.Entry(root, textvariable=self.folder_path, width=50)
        self.folder_entry.pack(pady=5)

        # Start download button
        self.download_button = tk.Button(root, text="Start Download", command=self.start_download)
        self.download_button.pack(pady=10)

        # Status log area
        self.status_text = tk.Text(root, height=10, width=60)
        self.status_text.pack(pady=10)
        self.status_text.insert(tk.END, "Status: Waiting for user input...\n")



    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.file_path.set(file_path)

        # Check if file exists
        if not os.path.exists(file_path):
            print(f"Error: File not found at {file_path}")
            exit(1)


    def select_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_path.set(folder_path)

    def add_metadata(self, file_path, title, artist):
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

    
    def start_download(self):
        file_path = self.file_path.get()
        folder_path = self.folder_path.get()

        print(file_path)
        print(folder_path)

        if not file_path or not folder_path:
            messagebox.showerror("Error", "Please select an Excel file and download folder.")
            return

        self.status_text.insert(tk.END, "Starting download process...\n")
        
        # Run in a separate thread to prevent UI freezing
        thread = threading.Thread(target=self.download_music, args=(file_path, folder_path))
        thread.start()

    def download_music(self, file_path, download_folder):
        try:
            sheets_dict = pd.read_excel(file_path, sheet_name=None)

            for playlist_name, df in sheets_dict.items():
                playlist_path = os.path.join(download_folder, playlist_name)
                os.makedirs(playlist_path, exist_ok=True)

                for _, row in df.iterrows():
                    song_name = str(row["Song Name"]).strip()
                    artist = str(row["Artist"]).strip()
                    youtube_url = str(row["YT Link"]).strip()


                    if pd.isna(youtube_url) or youtube_url == "":
                        self.status_text.insert(tk.END, f"Skipping {song_name} by {artist} (No URL provided)\n")
                        continue

                    output_filename = f"{song_name} ({artist})"
                    output_path = os.path.join(playlist_path, output_filename)


                    if glob.glob(f"{output_path}*.mp3"):
                        self.status_text.insert(tk.END, f"Skipping {output_filename}, already downloaded.\n")
                        continue

                    ydl_opts = {
                        "format": "bestaudio/best",  # Best quality audio
                        "outtmpl": output_path,  # Save file in the playlist folder
                        "quiet": True,             # Suppress most logs
                        "no-warnings": True,       # Suppress warnings
                        "cookies-from-browser": "chrome",
                        "postprocessors": [{
                            "key": "FFmpegExtractAudio",  # ✅ Convert to MP3
                            "preferredcodec": "mp3",
                            "preferredquality": "192",
                        }],
                    }

                    with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                        ydl.download([youtube_url])
                    
                    # Add metadata to the downloaded file
                    self.add_metadata(output_path, song_name, artist)

                    self.status_text.insert(tk.END, f"✅ Downloaded: {output_filename}\n")

            self.status_text.insert(tk.END, "\n🎉 Download process completed!\n")

        except Exception as e:
            self.status_text.insert(tk.END, f"\n❌ Error: {e}\n")

