# OpenCode Configuration

## Commands
- **Run GUI**: `python src/music_gui.py` or `python src/__init__.py`
- **Install dependencies**: `pip install -r requirements.txt`
- **Activate venv**: `source .venv/bin/activate` (if using virtual environment)
- **Test single file**: `python -m pytest path/to/test_file.py::test_function` (no tests currently exist)

## Code Style Guidelines
- **Imports**: Standard library first, then third-party (pandas, tkinter, yt_dlp), then local imports
- **Naming**: snake_case for functions/variables, PascalCase for classes (e.g., `MusicDownloaderApp`)
- **Functions**: Use descriptive docstrings with param/return descriptions
- **Error handling**: Use try/except blocks with specific error messages, return tuples (success, error_msg)
- **File paths**: Use `os.path.join()` for cross-platform compatibility
- **String handling**: Use f-strings for formatting, `.strip()` for cleaning input
- **GUI**: Use tkinter with proper widget organization and threading for long operations
- **Dialog handling**: Use `root.after()` and `wait_variable()` for thread-safe GUI operations

## Project Structure
- `src/`: Main source code
- `src/helpers/`: Utility modules (download_utils.py, metadata_utils.py)
- Excel files for music data input/output
- Dependencies: yt-dlp, pandas, openpyxl, tkinter

## Input File Format
- Excel file with columns: **Title | Artist | YouTube Link | Genre**
- Used by GUI to download and organize music files
- **New Feature**: If YouTube Link is empty, app will search YouTube with "Title Artist" and show 4 options to choose from
- **Skip Option**: Users can skip songs if no correct search results are found
- **Keyboard Shortcuts**: Press 1-4 to select videos, S to skip, Esc to cancel

## Key Libraries
- **yt-dlp**: YouTube downloading with audio extraction to MP3
- **pandas**: Excel file processing and data manipulation
- **tkinter**: GUI framework with progress bars and threading
- **ffmpeg**: Audio metadata addition (external dependency)