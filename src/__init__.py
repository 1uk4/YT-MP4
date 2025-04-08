# src/__init__.py
import tkinter as tk
from music_gui import MusicDownloaderApp

def main():
    root = tk.Tk()
    app = MusicDownloaderApp(root)
    root.mainloop()

if __name__ == "__main__":
    main()