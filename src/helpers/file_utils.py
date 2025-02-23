import os
import subprocess

def add_metadata(file_path, title, artist):
        temp_file = file_path.replace(".mp3", "_temp.mp3")
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
