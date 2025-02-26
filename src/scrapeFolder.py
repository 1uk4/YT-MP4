import os
import glob
import re
import pandas as pd

def clean_filename(filename):
    """Cleans filename by removing extra characters and formatting properly."""
    filename = os.path.splitext(filename)[0]  # Remove file extension
    filename = re.sub(r'[^a-zA-Z0-9 ]', ' ', filename)  # Remove special characters
    filename = re.sub(r'\s+', ' ', filename).strip()  # Remove extra spaces
    return filename

def list_mp3_files(folder_path):
    """
    Lists all MP3 file names, cleans them, and stores them in a pandas DataFrame.
    :param folder_path: Path to the folder containing MP3 files.
    :return: Pandas DataFrame with raw and cleaned file names.
    """
    if not os.path.isdir(folder_path):
        print("Error: The specified folder does not exist.")
        return pd.DataFrame()
    
    mp3_files = glob.glob(os.path.join(folder_path, "*.mp3"))
    
    data = []
    for file in mp3_files:
        raw_name = os.path.basename(file)
        cleaned_name = clean_filename(raw_name)
        data.append({"file_name": raw_name, "cleaned_name": cleaned_name})
    
    return pd.DataFrame(data)

def save_to_markdown(df, folder_path):
    """Saves the cleaned MP3 file names to a Markdown file."""
    markdown_path = os.path.join(folder_path, "mp3_files.md")
    with open(markdown_path, "w") as md_file:
        md_file.write("# Cleaned MP3 File Names\n\n")
        for _, row in df.iterrows():
            md_file.write(f"- {row['cleaned_name']}\n")
    print(f"Markdown file saved to {markdown_path}")

if __name__ == "__main__":
    folder = "../../../../Documents/Music/allsongs"

    mp3_df = list_mp3_files(folder)
    
    # Save cleaned names to a Markdown file
    save_to_markdown(mp3_df, folder)
    