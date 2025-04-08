import tkinter as tk
import threading
from tkinter import filedialog, messagebox, ttk
import os
import subprocess
from helpers.download_utils import download_music
from helpers.metadata_utils import add_metadata

class MusicDownloaderApp:
    def __init__(self, root):
        self.root = root
        self.root.title("Music Downloader")
        self.root.geometry("600x450")

        self.skip_current = False
        self.download_in_progress = False

        self.file_label = tk.Label(root, text="Select Excel File:")
        self.file_label.pack(pady=5)
        self.file_button = tk.Button(root, text="Choose File", command=self.select_file)
        self.file_button.pack(pady=5)
        self.file_path = tk.StringVar()
        self.file_entry = tk.Entry(root, textvariable=self.file_path, width=50)
        self.file_entry.pack(pady=5)

        self.folder_label = tk.Label(root, text="Select Download Folder:")
        self.folder_label.pack(pady=5)
        self.folder_button = tk.Button(root, text="Choose Folder", command=self.select_folder)
        self.folder_button.pack(pady=5)
        self.folder_path = tk.StringVar()
        self.folder_entry = tk.Entry(root, textvariable=self.folder_path, width=50)
        self.folder_entry.pack(pady=5)

        self.button_frame = tk.Frame(root)
        self.button_frame.pack(pady=10)

        self.download_button = tk.Button(self.button_frame, text="Start Download", command=self.start_download)
        self.download_button.pack(side=tk.LEFT, padx=5)

        self.skip_button = tk.Button(self.button_frame, text="Skip Current", command=self.skip_song, state=tk.DISABLED)
        self.skip_button.pack(side=tk.LEFT, padx=5)

        self.open_folder_button = tk.Button(self.button_frame, text="Open in Finder", command=self.open_download_folder)
        self.open_folder_button.pack(side=tk.LEFT, padx=5)

        self.progress_label = tk.Label(root, text="Download Progress:")
        self.progress_label.pack(pady=5)
        self.progress_bar = ttk.Progressbar(root, length=400, mode="determinate")
        self.progress_bar.pack(pady=5)
        self.progress_text = tk.Label(root, text="0%")
        self.progress_text.pack()

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
        folder_path = self.folder_path.get()

        if not folder_path or not os.path.exists(folder_path):
            messagebox.showerror("Error", "Please select a valid download folder.")
            return

        if os.name == "nt":
            subprocess.run(["explorer", folder_path], shell=True)
        elif os.name == "posix":
            if "darwin" in os.uname().sysname.lower():
                subprocess.run(["open", folder_path])
            else:
                subprocess.run(["xdg-open", folder_path])

    def skip_song(self):
        if self.download_in_progress:
            self.skip_current = True
            self.status_text.insert(tk.END, "⏭️ Skipping current song...\n")
            self.root.update()

    def start_download(self):
        file_path = self.file_path.get()
        folder_path = self.folder_path.get()

        if not file_path or not folder_path:
            messagebox.showerror("Error", "Please select an Excel file and download folder.")
            return

        self.download_in_progress = True
        self.skip_current = False
        self.skip_button.config(state=tk.NORMAL)
        self.download_button.config(state=tk.DISABLED)

        self.status_text.insert(tk.END, "Starting download process...\n")
        thread = threading.Thread(target=self.download_thread)
        thread.start()

    def download_thread(self):
        try:
            self.root.after(0, lambda: self.skip_button.config(state=tk.NORMAL))
            
            success = download_music(
                self.file_path.get(),
                self.folder_path.get(),
                self.status_text,
                self.progress_bar,
                self.progress_text,
                self.root,
                add_metadata
            )
            
            if success:
                self.status_text.insert("end", "\n🎉 Download process completed!\n")
        
        except Exception as e:
            self.status_text.insert(tk.END, f"Error: {str(e)}\n")
            return

        finally:
            self.download_in_progress = False
            self.root.after(0, lambda: self.skip_button.config(state=tk.DISABLED))
            self.root.after(0, lambda: self.download_button.config(state=tk.NORMAL))
            self.skip_current = False