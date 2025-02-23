import os
import tkinter as tk
import threading
import subprocess
from tkinter import filedialog, messagebox, ttk
from helpers import add_metadata, download_music


class MusicDownloaderApp:

    def __init__(self, root):
        self.root = root
        self.root.title("Music Downloader")
        self.root.geometry("600x450")

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

        # Open in Finder/File Explorer button
        self.open_folder_button = tk.Button(root, text="Open in Finder", command=self.open_download_folder)
        self.open_folder_button.pack(pady=5)


        # Progress Bar
        self.progress_label = tk.Label(root, text="Download Progress:")
        self.progress_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
        self.progress_bar.pack(pady=5)
        self.progress_text = tk.Label(root, text="0%")
        self.progress_text.pack()

        # Status log area
        self.status_text = tk.Text(root, height=10, width=50)
        self.status_text.pack(pady=10, padx=10)
        self.status_text.insert(tk.END, "Status: Waiting for user input...\n")

    def select_file(self):
        file_path = filedialog.askopenfilename(filetypes=[("Excel Files", "*.xlsx")])
        self.file_path.set(file_path)

    def select_folder(self):
        folder_path = filedialog.askdirectory()
        self.folder_path.set(folder_path)

    def open_download_folder(self):
        """ Opens the selected download folder in Finder (Mac) or File Explorer (Windows/Linux) """
        folder_path = self.folder_path.get()
        
        if not folder_path or not os.path.exists(folder_path):
            messagebox.showerror("Error", "Please select a valid download folder.")
            return
        
        if os.name == "nt":  # Windows
            subprocess.run(["explorer", folder_path], shell=True)
        elif os.name == "posix":
            if "darwin" in os.uname().sysname.lower():  # macOS
                subprocess.run(["open", folder_path])
            else:  # Linux
                subprocess.run(["xdg-open", folder_path])


    def start_download(self):
        file_path = self.file_path.get()
        folder_path = self.folder_path.get()

        if not file_path or not folder_path:
            messagebox.showerror("Error", "Please select an Excel file and download folder.")
            return

        self.status_text.insert(tk.END, "Starting download process...\n")
        thread = threading.Thread(target=download_music, args=(file_path, folder_path, self.status_text, self.progress_bar, self.progress_text, self.root, add_metadata))
        thread.start()
