import os
import pandas as pd
import tkinter as tk
from tkinter import ttk, messagebox
from .metadata_utils import add_metadata
import time
import yt_dlp

def safe_str(value):
    """Convert any value to string safely"""
    if pd.isna(value):  # Check for NaN/empty values
        return ''
    return str(value).strip()

def sanitize_filename(filename):
    """Remove or replace invalid characters for filenames"""
    invalid_chars = '<>:"/\\|?*'
    for char in invalid_chars:
        filename = filename.replace(char, '_')
    return filename

def get_safe_filepath(title, artist, playlist_folder):
    """Create filename in format 'Song Name (Artist).mp3'"""
    safe_title = sanitize_filename(title)
    safe_artist = sanitize_filename(artist)
    filename = f"{safe_title} ({safe_artist}).mp3"
    return os.path.join(playlist_folder, filename)

def search_youtube(title, artist):
    """Search YouTube for a song and return top 4 results"""
    search_query = f"{title} {artist} extended audio explicit"
    
    ydl_opts = {
        'quiet': True,
        'no_warnings': True,
        'extract_flat': True,
    }
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            search_results = ydl.extract_info(
                f"ytsearch4:{search_query}",
                download=False
            )
            
            results = []
            for entry in search_results.get('entries', []):
                result = {
                    'title': entry.get('title', 'Unknown Title'),
                    'uploader': entry.get('uploader', 'Unknown Channel'),
                    'duration': entry.get('duration', 0),
                    'view_count': entry.get('view_count', 0),
                    'url': f"https://www.youtube.com/watch?v={entry['id']}",
                    'thumbnail': entry.get('thumbnail', ''),
                    'id': entry['id'],
                    'search_query': search_query  # Add search query to results
                }
                results.append(result)
            
            return results
    except Exception as e:
        print(f"Search error: {e}")
        return []

def format_duration(seconds):
    """Convert seconds to MM:SS format"""
    if not seconds:
        return "Unknown"
    try:
        seconds = int(seconds)  # Convert to int to handle floats
        minutes = seconds // 60
        seconds = seconds % 60
        return f"{minutes}:{seconds:02d}"
    except (ValueError, TypeError):
        return "Unknown"

def format_view_count(views):
    """Convert view count to readable format (e.g., 1.2M, 543K)"""
    if not views or views == 0:
        return "Unknown views"
    try:
        views = int(views)
        if views >= 1_000_000:
            return f"{views / 1_000_000:.1f}M views"
        elif views >= 1_000:
            return f"{views / 1_000:.1f}K views"
        else:
            return f"{views} views"
    except (ValueError, TypeError):
        return "Unknown views"

class YouTubeSearchDialog:
    def __init__(self, parent, title, artist, search_results):
        self.selected_url = None
        self.search_results = search_results
        self.dialog = tk.Toplevel(parent)
        self.dialog.title(f"Select YouTube Video for: {title} - {artist}")
        self.dialog.geometry("900x700")  # Larger dialog
        self.dialog.transient(parent)
        self.dialog.grab_set()
        self.dialog.configure(bg="#2b2b2b")  # Dark background
        
        # Center the dialog
        self.dialog.update_idletasks()
        x = (self.dialog.winfo_screenwidth() // 2) - (900 // 2)
        y = (self.dialog.winfo_screenheight() // 2) - (700 // 2)
        self.dialog.geometry(f"900x700+{x}+{y}")
        
        # Bind keyboard events
        self.dialog.bind('<Key>', self.on_key_press)
        self.dialog.focus_set()  # Make sure dialog can receive key events
        
        self.create_widgets(search_results)
        
    def create_widgets(self, search_results):
        # Main container with padding
        main_frame = tk.Frame(self.dialog, bg="#2b2b2b")
        main_frame.pack(fill="both", expand=True, padx=20, pady=20)
        
        # Top section with search info and buttons
        top_frame = tk.Frame(main_frame, bg="#2b2b2b")
        top_frame.pack(fill="x", pady=(0, 15))
        
        # Search query display
        if search_results:
            search_query = search_results[0].get('search_query', 'Unknown')
            query_label = tk.Label(top_frame, text=f"Search: \"{search_query}\"", 
                                 font=("Arial", 12, "italic"), fg="#cccccc", bg="#2b2b2b")
            query_label.pack(pady=(0, 10))
        
        # Title and keyboard hint
        title_label = tk.Label(top_frame, text="Select the correct video (or press 1-4):", 
                              font=("Arial", 14, "bold"), fg="#ffffff", bg="#2b2b2b")
        title_label.pack(pady=(0, 10))
        
        # Buttons at the top
        button_frame = tk.Frame(top_frame, bg="#2b2b2b")
        button_frame.pack(pady=(0, 15))
        
        skip_btn = tk.Button(button_frame, text="Skip Song (S)", command=self.skip_song,
                            font=("Arial", 11, "bold"), padx=25, pady=8, bg="#ff9800", fg="white")
        skip_btn.pack(side=tk.LEFT, padx=(0, 15))
        
        cancel_btn = tk.Button(button_frame, text="Cancel All (Esc)", command=self.cancel,
                              font=("Arial", 11, "bold"), padx=25, pady=8, bg="#666666", fg="white")
        cancel_btn.pack(side=tk.LEFT)
        
        # Scrollable frame for results (takes up most of the space)
        canvas = tk.Canvas(main_frame, bg="#2b2b2b", highlightthickness=0)
        scrollbar = ttk.Scrollbar(main_frame, orient="vertical", command=canvas.yview)
        scrollable_frame = tk.Frame(canvas, bg="#2b2b2b")
        
        scrollable_frame.bind(
            "<Configure>",
            lambda e: canvas.configure(scrollregion=canvas.bbox("all"))
        )
        
        canvas.create_window((0, 0), window=scrollable_frame, anchor="nw")
        canvas.configure(yscrollcommand=scrollbar.set)
        
        # Create result cards
        for i, result in enumerate(search_results):
            self.create_result_card(scrollable_frame, result, i)
        
        # Add helpful text if no good results
        help_text = tk.Label(scrollable_frame, 
                           text="💡 Don't see the right video? Use 'Skip Song' to skip this track.",
                           font=("Arial", 10, "italic"), fg="#888888", bg="#2b2b2b")
        help_text.pack(pady=(15, 0))
        
        canvas.pack(side="left", fill="both", expand=True)
        scrollbar.pack(side="right", fill="y")
        
    def create_result_card(self, parent, result, index):
        # Card frame with hover effects and dark theme
        card_bg = "#3c3c3c"
        hover_bg = "#4a4a4a"
        
        card = tk.Frame(parent, relief="solid", borderwidth=1, padx=15, pady=15, 
                       bg=card_bg, cursor="hand2")
        card.pack(fill="x", padx=5, pady=5)  # More padding between cards
        
        # Make entire card clickable
        def on_click(event=None):
            self.select_video(result['url'])
        
        def on_enter(event=None):
            card.config(bg=hover_bg, relief="raised", borderwidth=2)
            for widget in card.winfo_children():
                if isinstance(widget, tk.Frame):
                    widget.config(bg=hover_bg)
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg=hover_bg)
                elif isinstance(widget, tk.Label):
                    widget.config(bg=hover_bg)
        
        def on_leave(event=None):
            card.config(bg=card_bg, relief="solid", borderwidth=1)
            for widget in card.winfo_children():
                if isinstance(widget, tk.Frame):
                    widget.config(bg=card_bg)
                    for child in widget.winfo_children():
                        if isinstance(child, tk.Label):
                            child.config(bg=card_bg)
                elif isinstance(widget, tk.Label):
                    widget.config(bg=card_bg)
        
        # Bind click events to card and all its children
        def bind_events(widget):
            widget.bind("<Button-1>", on_click)
            widget.bind("<Enter>", on_enter)
            widget.bind("<Leave>", on_leave)
            for child in widget.winfo_children():
                bind_events(child)
        
        bind_events(card)
        
        # Left side: Number and thumbnail
        left_frame = tk.Frame(card, bg=card_bg)
        left_frame.pack(side="left", padx=(0, 15))
        
        # Keyboard number indicator
        number_label = tk.Label(left_frame, text=f"{index + 1}", font=("Arial", 18, "bold"), 
                               fg="#ff9800", bg=card_bg, width=2)
        number_label.pack(pady=(0, 5))
        
        # Thumbnail placeholder with better styling
        thumb_frame = tk.Frame(left_frame, width=120, height=90, bg="#555555", relief="sunken", borderwidth=1)
        thumb_frame.pack()
        thumb_frame.pack_propagate(False)
        
        thumb_label = tk.Label(thumb_frame, text="🎵\nVideo", bg="#555555", 
                              font=("Arial", 10), fg="#cccccc")
        thumb_label.pack(expand=True)
        
        # Right side: Info frame with more space
        info_frame = tk.Frame(card, bg=card_bg)
        info_frame.pack(side="left", fill="both", expand=True)
        
        # Title with better styling and more space
        title_label = tk.Label(info_frame, text=result['title'], font=("Arial", 13, "bold"), 
                              wraplength=650, justify="left", bg=card_bg, fg="#ffffff", anchor="w")
        title_label.pack(anchor="w", fill="x", pady=(0, 8))
        
        # Channel with icon
        channel_label = tk.Label(info_frame, text=f"📺 {result['uploader']}", 
                                font=("Arial", 11), fg="#cccccc", bg=card_bg, anchor="w")
        channel_label.pack(anchor="w", fill="x", pady=(0, 5))
        
        # Stats row (duration and views)
        stats_frame = tk.Frame(info_frame, bg=card_bg)
        stats_frame.pack(anchor="w", fill="x", pady=(0, 8))
        
        # Duration with icon
        duration_label = tk.Label(stats_frame, text=f"⏱️ {format_duration(result['duration'])}", 
                                 font=("Arial", 10), fg="#cccccc", bg=card_bg)
        duration_label.pack(side="left", padx=(0, 20))
        
        # View count with icon
        view_count_label = tk.Label(stats_frame, text=f"👁️ {format_view_count(result['view_count'])}", 
                                   font=("Arial", 10), fg="#cccccc", bg=card_bg)
        view_count_label.pack(side="left")
        
        # Visual indicator that it's clickable
        click_hint = tk.Label(info_frame, text=f"Press {index + 1} or click to select →", 
                             font=("Arial", 10, "italic"), fg="#ff9800", bg=card_bg, anchor="w")
        click_hint.pack(anchor="w", fill="x")
        
    def on_key_press(self, event):
        """Handle keyboard shortcuts"""
        key = event.keysym
        
        # Number keys 1-4 for selecting videos
        if key in ['1', '2', '3', '4']:
            index = int(key) - 1
            if index < len(self.search_results):
                self.select_video(self.search_results[index]['url'])
        
        # S for skip
        elif key.lower() == 's':
            self.skip_song()
        
        # Escape for cancel
        elif key == 'Escape':
            self.cancel()
    
    def select_video(self, url):
        self.selected_url = url
        self.dialog.destroy()
        
    def skip_song(self):
        self.selected_url = "SKIP"  # Special value to indicate skip
        self.dialog.destroy()
        
    def cancel(self):
        self.selected_url = None
        self.dialog.destroy()

def download_with_ytdlp(url, output_path, filename):
    """Download using yt-dlp"""
    ydl_opts = {
        'format': 'bestaudio/best',
        'outtmpl': os.path.join(output_path, filename),
        'postprocessors': [{
            'key': 'FFmpegExtractAudio',
            'preferredcodec': 'mp3',
            'preferredquality': '192',
        }],
        'quiet': True,
        'no_warnings': True
    }
    
    try:
        with yt_dlp.YoutubeDL(ydl_opts) as ydl:
            ydl.download([url])
        return True, None
    except Exception as e:
        return False, str(e)

def find_row_index_by_title_artist(excel_path, target_title, target_artist):
    """Find the correct row index for a given title and artist"""
    try:
        df = pd.read_excel(excel_path)
        
        # Handle column structure
        if df.shape[1] == 4:
            df.columns = ['Title', 'Artist', 'YouTube Link', 'Genre']
        
        # Search for the row with matching title and artist
        for index, row in df.iterrows():
            title = safe_str(row.get('Title', ''))
            artist = safe_str(row.get('Artist', ''))
            
            if title == target_title and artist == target_artist:
                print(f"Found '{target_title}' by '{target_artist}' at row {index}")
                return index
        
        print(f"ERROR: Could not find '{target_title}' by '{target_artist}' in Excel file")
        return None
    except Exception as e:
        print(f"Error finding row index: {e}")
        return None

def update_excel_with_url(excel_path, row_index, url, expected_title=None, expected_artist=None):
    """Update Excel file with selected YouTube URL or SKIPPED status"""
    try:
        # Read the current Excel file
        df = pd.read_excel(excel_path)
        
        # Handle different possible column names for YouTube links
        youtube_column = None
        for col in df.columns:
            if col in ['YouTube Link', 'YT Link', 'YT LINK']:
                youtube_column = col
                break
        
        if youtube_column is None:
            # If no YouTube column found, assume it's the 3rd column (index 2)
            if len(df.columns) > 2:
                youtube_column = df.columns[2]
            else:
                print(f"Error: Not enough columns in Excel file. Found {len(df.columns)} columns")
                return False
        
        print(f"Updating row {row_index}, column '{youtube_column}' with: {url}")
        print(f"Current value in that cell: {df.at[row_index, youtube_column] if row_index < len(df) else 'ROW NOT FOUND'}")
        
        # Show the title and artist for this row to verify we're updating the right row
        if row_index < len(df):
            title_col = df.columns[0] if len(df.columns) > 0 else 'Unknown'
            artist_col = df.columns[1] if len(df.columns) > 1 else 'Unknown'
            current_title = df.at[row_index, title_col] if title_col in df.columns else 'Unknown'
            current_artist = df.at[row_index, artist_col] if artist_col in df.columns else 'Unknown'
            print(f"Row {row_index} contains: Title='{current_title}', Artist='{current_artist}'")
            
            # Verify we're updating the correct row
            if expected_title and expected_artist:
                if safe_str(current_title) != safe_str(expected_title) or safe_str(current_artist) != safe_str(expected_artist):
                    print(f"ERROR: Row mismatch! Expected '{expected_title}' by '{expected_artist}' but found '{current_title}' by '{current_artist}'")
                    return False
                else:
                    print(f"✅ Row verification passed: Correct title and artist match")
        else:
            print(f"ERROR: Row {row_index} does not exist in dataframe with {len(df)} rows")
            return False
        
        # Update the specific cell
        df.at[row_index, youtube_column] = url
        
        # Save immediately to Excel file
        df.to_excel(excel_path, index=False)
        print(f"Successfully saved Excel file with {len(df)} rows")
        
        # Verify the save worked
        verify_df = pd.read_excel(excel_path)
        saved_value = verify_df.at[row_index, youtube_column] if row_index < len(verify_df) else "NOT FOUND"
        print(f"Verification: Value now in Excel: {saved_value}")
        
        return True
    except Exception as e:
        print(f"Error updating Excel: {e}")
        import traceback
        traceback.print_exc()
        return False

def download_music(excel_path, download_folder, status_text, progress_bar, progress_text, root, add_metadata_func):
    try:
        # Read the Excel file
        status_text.insert(tk.END, "Reading Excel file...\n")
        
        # First, try to read with headers
        df = pd.read_excel(excel_path)
        
        # Check if we have the expected columns
        expected_columns = ['Title', 'Artist', 'YouTube Link', 'Genre']
        if not all(col in df.columns for col in expected_columns):
            # If not, assume no headers and add them
            status_text.insert(tk.END, "No proper headers found, detecting column structure...\n")
            df = pd.read_excel(excel_path, header=None)
            
            # Handle different column structures
            if df.shape[1] == 4:
                df.columns = ['Title', 'Artist', 'YouTube Link', 'Genre']
            elif df.shape[1] == 5:
                # Check if there are YouTube links in multiple columns
                status_text.insert(tk.END, "Detected 5 columns, merging YouTube link columns...\n")
                df.columns = ['Title', 'Artist', 'YouTube Link 1', 'Genre', 'YouTube Link 2']
                
                # Merge the YouTube link columns (prefer non-empty values)
                def merge_links(row):
                    link1 = safe_str(row['YouTube Link 1'])
                    link2 = safe_str(row['YouTube Link 2'])
                    # Skip header-like values
                    if link1 in ['Unnamed: 2', 'YT Link'] or not link1:
                        link1 = ''
                    if link2 in ['Unnamed: 2', 'YT Link'] or not link2:
                        link2 = ''
                    # Return the first non-empty link
                    return link1 if link1 else link2
                
                df['YouTube Link'] = df.apply(merge_links, axis=1)
                # Keep only the expected columns
                df = df[['Title', 'Artist', 'YouTube Link', 'Genre']]
            else:
                status_text.insert(tk.END, f"❌ Unexpected number of columns: {df.shape[1]}. Expected 4 or 5.\n")
                return False
        
        status_text.insert(tk.END, f"Found {len(df)} rows\n")
        status_text.insert(tk.END, f"Columns found: {', '.join(df.columns)}\n")
        
        # Debug: Show first few rows to verify structure
        status_text.insert(tk.END, "DEBUG: First 3 rows of data:\n")
        for i in range(min(3, len(df))):
            title = safe_str(df.iloc[i]['Title'] if 'Title' in df.columns else df.iloc[i, 0])
            artist = safe_str(df.iloc[i]['Artist'] if 'Artist' in df.columns else df.iloc[i, 1])
            url = safe_str(df.iloc[i]['YouTube Link'] if 'YouTube Link' in df.columns else df.iloc[i, 2])
            status_text.insert(tk.END, f"  Row {i}: '{title}' | '{artist}' | '{url}'\n")
        root.update()
        
        total_songs = len(df)
        
        # PHASE 1: Handle YouTube link searches
        status_text.insert(tk.END, "\n🔍 PHASE 1: Searching for missing YouTube links...\n")
        root.update()
        
        songs_needing_search = []
        for index, row in df.iterrows():
            title = safe_str(row.get('Title', ''))
            artist = safe_str(row.get('Artist', ''))
            url = safe_str(row.get('YouTube Link', ''))
            
            if not title or not artist:
                status_text.insert(tk.END, f"❌ Skipping row {index + 1}: Missing title or artist\n")
                continue
            
            # Skip if already has a YouTube URL (contains youtube.com or youtu.be)
            if url and ('youtube.com' in url.lower() or 'youtu.be' in url.lower()):
                status_text.insert(tk.END, f"✅ Skipping '{title} ({artist})' - Already has YouTube URL\n")
                continue
                
            # Skip if already marked as SKIPPED
            if url.upper() == "SKIPPED":
                status_text.insert(tk.END, f"⏭️ Skipping '{title} ({artist})' - Previously marked as skipped\n")
                continue
                
            # Only add to search list if no URL at all
            if not url:
                songs_needing_search.append((index, title, artist))
        
        if songs_needing_search:
            status_text.insert(tk.END, f"Found {len(songs_needing_search)} songs needing YouTube links\n")
            
            for i, (index, title, artist) in enumerate(songs_needing_search):
                status_text.insert(tk.END, f"🔍 Searching for '{title} ({artist})' ({i+1}/{len(songs_needing_search)})...\n")
                status_text.insert(tk.END, f"DEBUG: Processing row index {index} - Title: '{title}', Artist: '{artist}'\n")
                root.update()
                
                # Search and show dialog on main thread
                search_results = search_youtube(title, artist)
                if not search_results:
                    status_text.insert(tk.END, f"❌ No search results for '{title} ({artist})' - Skipping\n")
                    continue
                
                # Create a variable to store the result
                result_var = tk.StringVar()
                
                def show_dialog():
                    dialog = YouTubeSearchDialog(root, title, artist, search_results)
                    root.wait_window(dialog.dialog)
                    result_var.set(dialog.selected_url or "")
                
                # Schedule dialog on main thread
                root.after(0, show_dialog)
                root.wait_variable(result_var)
                
                selected_url = result_var.get()
                if not selected_url:
                    status_text.insert(tk.END, f"❌ Search cancelled - Stopping process\n")
                    return False
                elif selected_url == "SKIP":
                    status_text.insert(tk.END, f"⏭️ Skipping '{title} ({artist})' - User requested skip\n")
                    # Mark as SKIPPED in Excel file
                    status_text.insert(tk.END, f"DEBUG: About to mark '{title}' ({artist}) as SKIPPED in row {index}\n")
                    # Find the correct row index by searching for title and artist
                    correct_row_index = find_row_index_by_title_artist(excel_path, title, artist)
                    if correct_row_index is not None:
                        if update_excel_with_url(excel_path, correct_row_index, "SKIPPED", title, artist):
                            status_text.insert(tk.END, f"✅ Marked '{title} ({artist})' as SKIPPED in Excel\n")
                            # Reload the dataframe to ensure consistency
                            df = pd.read_excel(excel_path)
                            # Handle column structure again after reload
                            if df.shape[1] == 4:
                                df.columns = ['Title', 'Artist', 'YouTube Link', 'Genre']
                            root.update()
                        else:
                            status_text.insert(tk.END, f"⚠️ Failed to mark '{title} ({artist})' as SKIPPED\n")
                    else:
                        status_text.insert(tk.END, f"⚠️ Could not find '{title} ({artist})' in Excel file to mark as SKIPPED\n")
                    continue
                
                # Update Excel file with selected URL
                status_text.insert(tk.END, f"DEBUG: About to save URL for '{title}' ({artist}) to row {index}\n")
                # Find the correct row index by searching for title and artist
                correct_row_index = find_row_index_by_title_artist(excel_path, title, artist)
                if correct_row_index is not None:
                    if update_excel_with_url(excel_path, correct_row_index, selected_url, title, artist):
                        status_text.insert(tk.END, f"✅ Updated Excel with URL for '{title} ({artist})'\n")
                        # Reload the dataframe to ensure consistency
                        df = pd.read_excel(excel_path)
                        # Handle column structure again after reload
                        if df.shape[1] == 4:
                            df.columns = ['Title', 'Artist', 'YouTube Link', 'Genre']
                        root.update()
                    else:
                        status_text.insert(tk.END, f"⚠️ Failed to update Excel for '{title} ({artist})'\n")
                else:
                    status_text.insert(tk.END, f"⚠️ Could not find '{title} ({artist})' in Excel file\n")
        else:
            status_text.insert(tk.END, "✅ All songs already have YouTube links\n")
        
        # PHASE 2: Download all songs
        status_text.insert(tk.END, "\n⬇️ PHASE 2: Downloading songs...\n")
        root.update()
        
        downloaded_count = 0
        for index, row in df.iterrows():
            # Safely convert values to strings
            title = safe_str(row.get('Title', ''))
            artist = safe_str(row.get('Artist', ''))
            url = safe_str(row.get('YouTube Link', ''))
            genre = safe_str(row.get('Genre', 'Default'))
            
            if not title or not artist:
                continue
                
            # Skip if marked as SKIPPED
            if url.upper() == "SKIPPED":
                status_text.insert(tk.END, f"⏭️ Skipping '{title} ({artist})' - Marked as skipped\n")
                continue
                
            if not url:
                status_text.insert(tk.END, f"⏭️ Skipping '{title} ({artist})' - No YouTube link\n")
                continue

            try:
                # Create genre folder if it doesn't exist
                genre_folder = os.path.join(download_folder, sanitize_filename(genre))
                os.makedirs(genre_folder, exist_ok=True)

                # Check if file already exists
                expected_file_path = get_safe_filepath(title, artist, genre_folder)
                if os.path.exists(expected_file_path):
                    status_text.insert(tk.END, f"⏭️ Skipping '{title} ({artist})' - Already exists\n")
                    
                    # Update progress for skipped file
                    progress = (index + 1) / total_songs * 100
                    progress_bar['value'] = progress
                    progress_text.config(text=f"{progress:.1f}%")
                    root.update()
                    continue

                # Update status
                status_text.insert(tk.END, f"\nDownloading: {title} ({artist})\n")
                root.update()
                
                # Download using yt-dlp
                temp_filename = f"temp_{sanitize_filename(title)}"
                success, error = download_with_ytdlp(url, genre_folder, temp_filename)

                if not success:
                    status_text.insert(tk.END, f"❌ Download failed for '{title} ({artist})': {error}\n")
                    continue

                # Get the downloaded file (yt-dlp adds .mp3 extension)
                temp_file = os.path.join(genre_folder, temp_filename + '.mp3')
                
                if os.path.exists(temp_file):
                    # Rename to final filename
                    os.rename(temp_file, expected_file_path)
                    
                    # Add metadata
                    add_metadata_func(expected_file_path, title, artist)
                    
                    downloaded_count += 1
                    status_text.insert(tk.END, f"✅ Successfully downloaded: {title} ({artist})\n")
                else:
                    status_text.insert(tk.END, f"❌ Failed to download '{title} ({artist})'\n")
                
            except Exception as e:
                status_text.insert(tk.END, f"❌ Error processing '{title} ({artist})': {str(e)}\n")
                # Cleanup any partial downloads
                if os.path.exists(expected_file_path):
                    try:
                        os.remove(expected_file_path)
                    except:
                        pass
                continue
            
            # Update progress
            progress = (index + 1) / total_songs * 100
            progress_bar['value'] = progress
            progress_text.config(text=f"{progress:.1f}%")
            root.update()
            
        status_text.insert(tk.END, f"\n🎉 Download process completed! Downloaded {downloaded_count} new songs.\n")
        return True
        
    except pd.errors.EmptyDataError:
        status_text.insert(tk.END, "❌ Error: Excel file is empty\n")
        return False
    except Exception as e:
        status_text.insert(tk.END, f"❌ Error reading Excel file: {str(e)}\n")
        return False