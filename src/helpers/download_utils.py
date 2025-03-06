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

def format_title_proper(text):
    """ Capitalizes each word except small words like 'of', 'the', 'and'. """
    small_words = {"of", "the", "and", "in", "on", "at", "a", "an", "to"}
    words = text.split()
    return " ".join([word.capitalize() if word.lower() not in small_words else word.lower() for word in words])


def download_music(file_path, folder_path, status_text, progress_bar, progress_text, root, add_metadata):
    try:
        # Read single sheet excel file
        df = pd.read_excel(file_path)
        total_songs = len(df)
        downloaded_count = 0

        # Create a dictionary to group songs by playlist
        playlist_groups = df.groupby("Playlist")

        for playlist_name, playlist_df in playlist_groups:
            playlist_path = os.path.join(folder_path, playlist_name)
            os.makedirs(playlist_path, exist_ok=True)

            for _, row in playlist_df.iterrows():
                song_name = str(row["Song Name"]).strip()
                artist = str(row["Artist"]).strip()
                youtube_url = str(row["YT Link"]).strip()

                if pd.isna(youtube_url) or youtube_url == "":
                    status_text.insert("end", f"Skipping {song_name} by {artist} (No URL provided)\n")
                    downloaded_count += 1
                    update_progress(progress_bar, progress_text, downloaded_count, total_songs, root)
                    continue


                output_filename = f"{format_title_proper(song_name)} ({format_title_proper(artist)})"
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