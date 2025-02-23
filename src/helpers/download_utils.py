import os
import pandas as pd
import yt_dlp
import glob

def update_progress(progress_bar, progress_text, current, total, root):
    """ Updates the progress bar and percentage text """
    percentage = int((current / total) * 100)
    progress_bar["value"] = percentage
    progress_text.config(text=f"{percentage}% ({current}/{total} Songs)")
    root.update_idletasks()

def download_music(file_path, folder_path, status_text, progress_bar, progress_text, root, add_metadata):
    try:
        sheets_dict = pd.read_excel(file_path, sheet_name=None)
        total_songs = sum(len(df) for df in sheets_dict.values())
        downloaded_count = 0

        for playlist_name, df in sheets_dict.items():
            playlist_path = os.path.join(folder_path, playlist_name)
            os.makedirs(playlist_path, exist_ok=True)

            for _, row in df.iterrows():
                song_name = str(row["Song Name"]).strip()
                artist = str(row["Artist"]).strip()
                youtube_url = str(row["YT Link"]).strip()

                if pd.isna(youtube_url) or youtube_url == "":
                    status_text.insert("end", f"Skipping {song_name} by {artist} (No URL provided)\n")
                    downloaded_count += 1
                    update_progress(progress_bar, progress_text, downloaded_count, total_songs, root)
                    continue

                output_filename = f"{song_name} ({artist})"
                output_path = os.path.join(playlist_path, output_filename)

                if glob.glob(f"{output_path}*.mp3"):
                    status_text.insert("end", f"Skipping {output_filename}, already downloaded.\n")
                    downloaded_count += 1
                    update_progress(progress_bar, progress_text, downloaded_count, total_songs, root)
                    continue

                ydl_opts = {
                    "format": "bestaudio/best",
                    "outtmpl": f"{output_path}.%(ext)s",
                    "quiet": True,
                    "no-warnings": True,
                    "cookies-from-browser": "chrome",
                    "postprocessors": [{
                        "key": "FFmpegExtractAudio",
                        "preferredcodec": "mp3",
                        "preferredquality": "192",
                    }],
                }

                with yt_dlp.YoutubeDL(ydl_opts) as ydl:
                    ydl.download([youtube_url])

                mp3_file = f"{output_path}.mp3"
                add_metadata(mp3_file, song_name, artist)

                status_text.insert("end", f"✅ Downloaded: {output_filename}.mp3\n")

                downloaded_count += 1
                update_progress(progress_bar, progress_text, downloaded_count, total_songs, root)

        status_text.insert("end", "\n🎉 Download process completed!\n")
        update_progress(progress_bar, progress_text, total_songs, total_songs, root)  # Set progress to 100%

    except Exception as e:
        status_text.insert("end", f"\n❌ Error: {e}\n")