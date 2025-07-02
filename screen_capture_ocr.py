#!/usr/bin/env python3
"""
Screen Capture OCR to CSV Script
Captures screenshots intermittently and uses Mistral OCR API to extract table data as CSV.
"""

import os
import csv
import json
import base64
import time
from datetime import datetime
import argparse
import sys
import re
import shutil

# Try to import dotenv, but don't fail if it's not installed
try:
    from dotenv import load_dotenv
    load_dotenv()
    DOTENV_AVAILABLE = True
except ImportError:
    DOTENV_AVAILABLE = False
    print("Note: python-dotenv package not installed. .env file support unavailable.")
    print("To install: pip install python-dotenv")

# Mistral SDK import
try:
    from mistralai import Mistral
    MISTRAL_SDK_AVAILABLE = True
except ImportError:
    MISTRAL_SDK_AVAILABLE = False
    print("Note: mistralai package not installed. OCR functionality unavailable.")
    print("To install: pip install mistralai")

# Screen capture imports
try:
    import mss
    MSS_AVAILABLE = True
except ImportError:
    MSS_AVAILABLE = False
    print("Note: mss package not installed. Screen capture unavailable.")
    print("To install: pip install mss")
    
try:
    import pyautogui
    PYAUTOGUI_AVAILABLE = True
except ImportError:
    PYAUTOGUI_AVAILABLE = False
    print("Note: pyautogui package not installed. Alternative screen capture unavailable.")
    print("To install: pip install pyautogui")

# Image comparison imports
try:
    from PIL import Image
    import imagehash
    PIL_AVAILABLE = True
except ImportError:
    PIL_AVAILABLE = False
    print("Note: Pillow package not installed. Image comparison unavailable.")
    print("To install: pip install Pillow")

# Window management imports
try:
    import pygetwindow as gw
    PYGETWINDOW_AVAILABLE = True
except ImportError:
    PYGETWINDOW_AVAILABLE = False

# Modern window management library (recommended)
try:
    import pywinctl as pwc
    PYWINCTL_AVAILABLE = True
except ImportError:
    PYWINCTL_AVAILABLE = False
    
# Print import status
if not PYGETWINDOW_AVAILABLE and not PYWINCTL_AVAILABLE:
    print("Note: No window management library found.")
    print("Install PyWinCtl (recommended): pip install pywinctl")
    print("Or PyGetWindow: pip install pygetwindow")

# GUI imports for preview window
try:
    import tkinter as tk
    from tkinter import ttk
    import threading
    TKINTER_AVAILABLE = True
except ImportError:
    TKINTER_AVAILABLE = False
    print("Note: tkinter not available. Preview window unavailable.")

# Configuration
MISTRAL_API_KEY = os.environ.get("MISTRAL_API_KEY")
OCR_MODEL_NAME = "mistral-ocr-latest"

def list_windows_macos():
    """Lists windows on macOS using AppleScript."""
    import subprocess
    
    try:
        # Get list of applications with windows
        script = '''
        tell application "System Events"
            set appList to {}
            repeat with p in (every process whose visible is true)
                try
                    set appName to name of p
                    set windowList to every window of p
                    if (count of windowList) > 0 then
                        repeat with w in windowList
                            set windowTitle to name of w
                            if windowTitle is not "" then
                                set end of appList to (appName & " - " & windowTitle)
                            end if
                        end repeat
                    end if
                end try
            end repeat
            return appList
        end tell
        '''
        
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        if result.returncode == 0:
            window_titles = result.stdout.strip().split(', ')
            return [title for title in window_titles if title]
        else:
            return []
    except Exception as e:
        print(f"Error listing macOS windows: {e}")
        return []

def list_windows():
    """Lists all available windows."""
    if sys.platform == 'darwin':  # macOS
        return list_windows_macos()
    
    if not PYGETWINDOW_AVAILABLE:
        return []
    
    try:
        # Try different API methods (pygetwindow versions vary)
        if hasattr(gw, 'getAllWindows'):
            windows = gw.getAllWindows()
        elif hasattr(gw, 'getWindowsAt'):
            # Some versions use different methods
            windows = []
            # This is a fallback that may not work well
        else:
            print("pygetwindow API not compatible with this version")
            return []
            
        visible_windows = [w for w in windows if w.title and w.width > 0 and w.height > 0]
        return visible_windows
    except Exception as e:
        print(f"Error listing windows: {e}")
        print("Window selection may not be available on this system.")
        return []

def select_window():
    """Prompts user to select a window to capture."""
    windows = list_windows()
    
    if not windows:
        print("No windows found.")
        return None
    
    print("\nAvailable windows:")
    for i, window in enumerate(windows, 1):
        # Handle different window types
        if isinstance(window, str):
            # macOS AppleScript returns strings
            title = window[:50] + ('...' if len(window) > 50 else '')
        elif hasattr(window, 'title') and callable(window.title):
            # PyGetWindow object with title method
            title = window.title()[:50] + ('...' if len(window.title()) > 50 else '')
        elif hasattr(window, 'title'):
            # Object with title attribute
            title = str(window.title)[:50] + ('...' if len(str(window.title)) > 50 else '')
        else:
            # Fallback to string conversion
            title = str(window)[:50] + ('...' if len(str(window)) > 50 else '')
        print(f"{i:2d}. {title}")
    
    while True:
        try:
            choice = input(f"\nSelect window (1-{len(windows)}) or press Enter for full screen: ").strip()
            if not choice:
                return None  # Full screen
            
            choice = int(choice)
            if 1 <= choice <= len(windows):
                selected_window = windows[choice - 1]
                # Handle different window types for display
                if isinstance(selected_window, str):
                    print(f"Selected: {selected_window}")
                elif hasattr(selected_window, 'title') and callable(selected_window.title):
                    print(f"Selected: {selected_window.title()}")
                elif hasattr(selected_window, 'title'):
                    print(f"Selected: {selected_window.title}")
                else:
                    print(f"Selected: {selected_window}")
                return selected_window
            else:
                print(f"Please enter a number between 1 and {len(windows)}")
        except ValueError:
            print("Please enter a valid number")

def capture_window_macos(window_name, output_path):
    """Capture a specific window on macOS using AppleScript."""
    import subprocess
    
    try:
        # Parse app name and window title from window_name
        if " - " in window_name:
            app_name, window_title = window_name.split(" - ", 1)
        else:
            app_name = window_name
            window_title = ""
        
        # Try different approaches for capturing the window
        
        # Method 1: Try using process name directly
        script1 = f'''
        tell application "System Events"
            tell process "{app_name}"
                set frontmost to true
                if "{window_title}" is not "" then
                    set targetWindow to window "{window_title}"
                else
                    set targetWindow to window 1
                end if
                
                set {{x, y}} to position of targetWindow
                set {{w, h}} to size of targetWindow
                
                do shell script "screencapture -R" & x & "," & y & "," & w & "," & h & " '{output_path}'"
            end tell
        end tell
        '''
        
        result = subprocess.run(['osascript', '-e', script1], capture_output=True, text=True)
        if result.returncode == 0:
            return output_path
        
        # Method 2: Try finding the real application name
        print(f"Direct process capture failed, trying to find real app name for '{app_name}'...")
        
        # Get the actual bundle identifier or application name
        find_app_script = f'''
        tell application "System Events"
            set appList to {{}}
            repeat with p in (every process whose visible is true)
                set pName to name of p
                if pName contains "{app_name}" then
                    set end of appList to pName
                end if
            end repeat
            return appList
        end tell
        '''
        
        app_result = subprocess.run(['osascript', '-e', find_app_script], capture_output=True, text=True)
        if app_result.returncode == 0 and app_result.stdout.strip():
            real_app_names = app_result.stdout.strip().split(', ')
            for real_app_name in real_app_names:
                script2 = f'''
                tell application "System Events"
                    tell process "{real_app_name}"
                        set frontmost to true
                        if "{window_title}" is not "" then
                            set targetWindow to window "{window_title}"
                        else
                            set targetWindow to window 1
                        end if
                        
                        set {{x, y}} to position of targetWindow
                        set {{w, h}} to size of targetWindow
                        
                        do shell script "screencapture -R" & x & "," & y & "," & w & "," & h & " '{output_path}'"
                    end tell
                end tell
                '''
                
                result2 = subprocess.run(['osascript', '-e', script2], capture_output=True, text=True)
                if result2.returncode == 0:
                    print(f"Successfully captured using app name: {real_app_name}")
                    return output_path
        
        print(f"Could not capture window '{window_name}' - app not found or not accessible")
        return None
            
    except Exception as e:
        print(f"Error capturing macOS window: {e}")
        return None

def take_screenshot(output_path="screenshot.png", window=None):
    """Takes a screenshot using available libraries."""
    # Handle macOS window capture
    if window and sys.platform == 'darwin' and isinstance(window, str):
        result = capture_window_macos(window, output_path)
        if result:
            return result
        else:
            print("Falling back to full screen capture...")
    
    # Handle pygetwindow objects
    elif window and PYGETWINDOW_AVAILABLE and hasattr(window, 'left'):
        try:
            # Capture specific window
            if MSS_AVAILABLE:
                with mss.mss() as sct:
                    # Get window bounds
                    left, top, width, height = window.left, window.top, window.width, window.height
                    monitor = {"top": top, "left": left, "width": width, "height": height}
                    screenshot = sct.grab(monitor)
                    from mss.tools import to_png
                    to_png(screenshot.rgb, screenshot.size, output=output_path)
                    return output_path
            elif PYAUTOGUI_AVAILABLE:
                # Bring window to front and capture
                window.activate()
                time.sleep(0.5)  # Wait for window to come to front
                screenshot = pyautogui.screenshot(region=(window.left, window.top, window.width, window.height))
                screenshot.save(output_path)
                return output_path
        except Exception as e:
            print(f"Error capturing window: {e}")
            print("Falling back to full screen capture...")
    
    # Full screen capture (fallback or default)
    if MSS_AVAILABLE:
        with mss.mss() as sct:
            # Get the primary monitor
            monitor = sct.monitors[1]
            screenshot = sct.grab(monitor)
            from mss.tools import to_png
            to_png(screenshot.rgb, screenshot.size, output=output_path)
            return output_path
    elif PYAUTOGUI_AVAILABLE:
        screenshot = pyautogui.screenshot()
        screenshot.save(output_path)
        return output_path
    else:
        raise RuntimeError("No screen capture library available. Install mss or pyautogui.")

def get_user_column_headers():
    """Prompts user for column headers/data points to extract."""
    print("\n=== Screen Capture OCR CSV Extractor ===")
    print("This script will capture screenshots and extract table data to CSV.")
    print("\nWhat data points do you want to extract from the tables?")
    print("Enter each column header/data point (press Enter twice when done):")
    
    headers = []
    while True:
        header = input(f"Column {len(headers) + 1}: ").strip()
        if not header:
            if headers:
                break
            else:
                print("Please enter at least one column header.")
                continue
        headers.append(header)
    
    print(f"\nYou entered {len(headers)} column headers: {', '.join(headers)}")
    return headers

def encode_image(image_path):
    """Encode the image to base64."""
    try:
        with open(image_path, "rb") as image_file:
            return base64.b64encode(image_file.read()).decode('utf-8')
    except FileNotFoundError:
        print(f"Error: The file {image_path} was not found.")
        return None
    except Exception as e:
        print(f"Error: {e}")
        return None

def perform_ocr_on_image(api_key, model_name, image_path):
    """Performs OCR on an image using Mistral's OCR API."""
    if not MISTRAL_SDK_AVAILABLE:
        raise RuntimeError("Mistral SDK not available. Install with: pip install mistralai")
    
    # Initialize Mistral client
    client = Mistral(api_key=api_key)
    
    # Encode image to base64
    base64_image = encode_image(image_path)
    if not base64_image:
        raise ValueError(f"Failed to encode image: {image_path}")
    
    # Determine image format from file extension
    image_ext = os.path.splitext(image_path)[1].lower()
    if image_ext == '.png':
        mime_type = 'image/png'
    elif image_ext in ['.jpg', '.jpeg']:
        mime_type = 'image/jpeg'
    else:
        mime_type = 'image/png'  # Default fallback
    
    # Perform OCR using Mistral SDK
    ocr_response = client.ocr.process(
        model=model_name,
        document={
            "type": "image_url",
            "image_url": f"data:{mime_type};base64,{base64_image}"
        },
        include_image_base64=False  # We don't need images back
    )
    
    return ocr_response

def parse_markdown_table_from_text(markdown_text):
    """Parses the first markdown table found in a string."""
    if not markdown_text:
        return None
        
    lines = markdown_text.strip().split('\n')
    table_started = False
    table_data = []
    header_parsed = False
    column_count = 0

    for line in lines:
        stripped_line = line.strip()
        
        is_table_line = stripped_line.startswith('|') and stripped_line.endswith('|')
        
        if is_table_line:
            if not table_started:
                 table_started = True
            
            cells = [cell.strip() for cell in stripped_line[1:-1].split('|')]

            if not header_parsed:
                # This is the header row
                table_data.append(cells)
                header_parsed = True
                column_count = len(cells)
            else:
                # Check if this is a separator line
                if len(cells) == column_count:
                    is_separator = all(all(c in '-: ' for c in cell) and cell for cell in cells)
                    if is_separator:
                        continue
                
                # This is a data row
                table_data.append(cells)
        elif table_started: 
            break 
            
    if not table_data or len(table_data) < 1:
        return None
    if not any(cell for row in table_data for cell in row):
        return None
        
    return table_data

def format_ocr_with_structured_output(api_key, ocr_text, target_headers):
    """Uses Mistral chat completion with JSON mode to properly format OCR text into structured table data."""
    if not MISTRAL_SDK_AVAILABLE:
        raise RuntimeError("Mistral SDK not available")
    
    # Initialize Mistral client
    client = Mistral(api_key=api_key)
    
    # Create the prompt for structured output
    headers_list = ', '.join(target_headers)
    prompt = f"""You are a data extraction assistant. Please extract table data from the following OCR text and format it as JSON.

Target columns: {headers_list}

OCR Text:
{ocr_text}

Instructions:
1. Look for tabular data in the text
2. Extract each row of data
3. Map the data to the target columns: {headers_list}
4. Return a JSON object with a "rows" array
5. Each row should be an object with keys matching the target column names
6. If a column value is missing or unclear, use an empty string
7. Clean up any formatting issues or incomplete text
8. Only include actual data rows, not headers

Example format:
{{
  "rows": [
    {{"Person Name": "John Doe", "Company Name": "Acme Corp", "Job Title": "Manager"}},
    {{"Person Name": "Jane Smith", "Company Name": "Tech Inc", "Job Title": "Developer"}}
  ]
}}"""

    try:
        # Make the API call with JSON mode
        chat_response = client.chat.complete(
            model="mistral-medium-2505",
            messages=[{
                "role": "user",
                "content": prompt
            }],
            response_format={
                "type": "json_object"
            }
        )
        
        # Parse the JSON response
        response_content = chat_response.choices[0].message.content
        parsed_data = json.loads(response_content)
        
        # Convert to table format
        if 'rows' in parsed_data and parsed_data['rows']:
            formatted_table = [target_headers]  # Header row
            
            for row_obj in parsed_data['rows']:
                row = []
                for header in target_headers:
                    row.append(row_obj.get(header, ""))
                formatted_table.append(row)
            
            return formatted_table
        
        return None
        
    except Exception as e:
        print(f"Error in structured formatting: {e}")
        return None

def extract_table_with_headers(ocr_response, target_headers, api_key=None):
    """Extracts table data and matches it to target headers."""
    # First get the raw OCR text
    raw_text = ""
    
    # Handle Mistral SDK response format
    if hasattr(ocr_response, 'pages'):
        pages = ocr_response.pages
    elif isinstance(ocr_response, dict) and 'pages' in ocr_response:
        pages = ocr_response['pages']
    else:
        return None
    
    if not pages:
        return None
    
    # Collect all text content
    for page_data in pages:
        # Handle both dict and object attributes
        if hasattr(page_data, 'markdown'):
            raw_text += page_data.markdown + "\n"
        elif hasattr(page_data, 'text'):
            raw_text += page_data.text + "\n"
        elif isinstance(page_data, dict):
            raw_text += page_data.get("markdown", "") + "\n"
            raw_text += page_data.get("text", "") + "\n"
    
    if raw_text.strip() and api_key:
        # Use structured output to format the data properly
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Using Mistral Medium for structured data formatting...")
        formatted_table = format_ocr_with_structured_output(api_key, raw_text.strip(), target_headers)
        if formatted_table:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] ✓ Structured formatting successful")
            return formatted_table
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Structured formatting failed, falling back to markdown parsing")
    
    # Fallback to original parsing method
    for page_data in pages:
        # Handle both dict and object attributes
        if hasattr(page_data, 'markdown'):
            markdown_content = page_data.markdown
        elif isinstance(page_data, dict):
            markdown_content = page_data.get("markdown", "")
        else:
            continue
            
        if markdown_content:
            table_data = parse_markdown_table_from_text(markdown_content)
            if table_data and len(table_data) > 0:
                # Use target headers as the first row
                formatted_table = [target_headers]
                
                # Extract data rows (skip original header if exists)
                data_rows = table_data[1:] if len(table_data) > 1 else []
                
                # Pad or trim rows to match header count
                for row in data_rows:
                    if len(row) >= len(target_headers):
                        formatted_table.append(row[:len(target_headers)])
                    else:
                        padded_row = row + [''] * (len(target_headers) - len(row))
                        formatted_table.append(padded_row)
                
                return formatted_table
    
    return None

def save_table_to_csv(table_data, csv_file_path):
    """Saves table data (list of lists) to a CSV file."""
    with open(csv_file_path, "w", newline="", encoding="utf-8") as f:
        writer = csv.writer(f)
        writer.writerows(table_data)

def deduplicate_csv(csv_file_path):
    """Remove duplicate rows from CSV file while preserving header."""
    try:
        # Read the CSV file
        with open(csv_file_path, "r", newline="", encoding="utf-8") as f:
            reader = csv.reader(f)
            rows = list(reader)
        
        if not rows:
            print("CSV file is empty")
            return 0
        
        # Keep header and deduplicate data rows
        header = rows[0]
        data_rows = rows[1:]
        
        # Convert rows to tuples for deduplication, then back to lists
        original_count = len(data_rows)
        unique_rows = []
        seen = set()
        
        for row in data_rows:
            row_tuple = tuple(row)
            if row_tuple not in seen:
                seen.add(row_tuple)
                unique_rows.append(row)
        
        # Write deduplicated data back to file
        all_rows = [header] + unique_rows
        with open(csv_file_path, "w", newline="", encoding="utf-8") as f:
            writer = csv.writer(f)
            writer.writerows(all_rows)
        
        duplicates_removed = original_count - len(unique_rows)
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Deduplication complete:")
        print(f"  - Original rows: {original_count}")
        print(f"  - Unique rows: {len(unique_rows)}")
        print(f"  - Duplicates removed: {duplicates_removed}")
        
        return duplicates_removed
        
    except Exception as e:
        print(f"Error during deduplication: {e}")
        return 0

def images_are_similar(image_path1, image_path2, threshold=5):
    """Compare two images using perceptual hashing to detect similarity."""
    if not PIL_AVAILABLE:
        # If PIL not available, always return False (process all images)
        return False
    
    try:
        # Open both images
        img1 = Image.open(image_path1)
        img2 = Image.open(image_path2)
        
        # Calculate perceptual hashes
        hash1 = imagehash.average_hash(img1)
        hash2 = imagehash.average_hash(img2)
        
        # Calculate hamming distance (lower = more similar)
        distance = hash1 - hash2
        
        # Return True if images are similar (within threshold)
        return distance <= threshold
        
    except Exception as e:
        print(f"Error comparing images: {e}")
        return False

class PreviewWindow:
    """A preview window to show what's being captured."""
    
    def __init__(self, selected_window=None):
        self.selected_window = selected_window
        self.running = False
        self.root = None
        self.canvas = None
        self.photo = None
        self.update_interval = 1000  # Update every 1 second
        
    def start(self):
        """Start the preview window in a separate thread."""
        if not TKINTER_AVAILABLE:
            print("Preview window unavailable - tkinter not installed")
            return
            
        # Check for macOS threading issues
        if sys.platform == 'darwin':
            print("Preview window disabled on macOS due to threading restrictions")
            print("Use --no-preview to suppress this message")
            return
            
        self.running = True
        preview_thread = threading.Thread(target=self._run_preview, daemon=True)
        preview_thread.start()
        
    def stop(self):
        """Stop the preview window."""
        self.running = False
        if self.root:
            self.root.quit()
            
    def _run_preview(self):
        """Run the preview window GUI."""
        try:
            self.root = tk.Tk()
            self.root.title("Screen Capture Preview")
            self.root.geometry("400x300")
            
            # Create canvas for image display
            self.canvas = tk.Canvas(self.root, bg='black')
            self.canvas.pack(fill=tk.BOTH, expand=True)
            
            # Add status label
            status_text = f"Previewing: {self.selected_window.title if self.selected_window else 'Full Screen'}"
            status_label = ttk.Label(self.root, text=status_text)
            status_label.pack(pady=5)
            
            # Schedule first update
            self.root.after(100, self._update_preview)
            
            # Handle window close
            self.root.protocol("WM_DELETE_WINDOW", self._on_closing)
            
            self.root.mainloop()
        except Exception as e:
            print(f"Preview window error: {e}")
            
    def _update_preview(self):
        """Update the preview image."""
        if not self.running:
            return
            
        try:
            # Take a screenshot for preview
            temp_path = "temp_preview.png"
            take_screenshot(temp_path, self.selected_window)
            
            # Load and resize image for preview
            if PIL_AVAILABLE:
                img = Image.open(temp_path)
                
                # Calculate size to fit canvas while maintaining aspect ratio
                canvas_width = self.canvas.winfo_width()
                canvas_height = self.canvas.winfo_height()
                
                if canvas_width > 1 and canvas_height > 1:  # Canvas is initialized
                    img.thumbnail((canvas_width, canvas_height), Image.Resampling.LANCZOS)
                    
                    # Convert to PhotoImage
                    try:
                        from PIL import ImageTk
                        self.photo = ImageTk.PhotoImage(img)
                        
                        # Clear canvas and display image
                        self.canvas.delete("all")
                        self.canvas.create_image(
                            canvas_width // 2, canvas_height // 2,
                            image=self.photo, anchor=tk.CENTER
                        )
                    except ImportError:
                        # Fallback if ImageTk not available
                        self.canvas.delete("all")
                        self.canvas.create_text(
                            canvas_width // 2, canvas_height // 2,
                            text="Preview unavailable\n(PIL ImageTk required)",
                            fill="white", justify=tk.CENTER
                        )
                        
            # Clean up temp file
            if os.path.exists(temp_path):
                os.remove(temp_path)
                
        except Exception as e:
            print(f"Preview update error: {e}")
            
        # Schedule next update
        if self.running and self.root:
            self.root.after(self.update_interval, self._update_preview)
            
    def _on_closing(self):
        """Handle window closing."""
        self.running = False
        self.root.destroy()

def activate_window_macos(window_name):
    """Activate a specific window on macOS using AppleScript."""
    import subprocess
    
    try:
        # Parse app name from window_name
        if " - " in window_name:
            app_name, window_title = window_name.split(" - ", 1)
        else:
            app_name = window_name
            window_title = ""
        
        # Try to activate the window
        script = f'''
        tell application "System Events"
            tell process "{app_name}"
                set frontmost to true
                if "{window_title}" is not "" then
                    tell window "{window_title}" to perform action "AXRaise"
                else
                    tell window 1 to perform action "AXRaise"
                end if
            end tell
        end tell
        '''
        
        result = subprocess.run(['osascript', '-e', script], capture_output=True, text=True)
        return result.returncode == 0
        
    except Exception as e:
        print(f"Error activating window: {e}")
        return False

def send_arrow_keys(count=11):
    """Send arrow down keystrokes."""
    if PYAUTOGUI_AVAILABLE:
        try:
            for i in range(count):
                pyautogui.press('down')
                time.sleep(0.1)  # Small delay between keystrokes
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Sent {count} arrow down keystrokes")
            return True
        except Exception as e:
            print(f"Error sending keystrokes: {e}")
            return False
    else:
        print("PyAutoGUI not available - cannot send keystrokes")
        return False

def screen_capture_mode(api_key, model_name, target_headers, output_csv, interval=10, selected_window=None, show_preview=True, debug_mode=False, wait_time=5, arrow_strokes=11):
    """Runs screen capture mode with intermittent OCR processing."""
    print(f"\n=== Starting Screen Capture Mode ===")
    if selected_window:
        # Handle different window types
        if isinstance(selected_window, str):
            print(f"Capturing window: {selected_window}")
        elif hasattr(selected_window, 'title') and callable(selected_window.title):
            print(f"Capturing window: {selected_window.title()}")
        elif hasattr(selected_window, 'title'):
            print(f"Capturing window: {selected_window.title}")
        else:
            print(f"Capturing window: {selected_window}")
    else:
        print("Capturing full screen")
    print(f"Target columns: {', '.join(target_headers)}")
    print(f"Screenshot interval: {interval} seconds")
    print(f"Output file: {output_csv}")
    print(f"Press Ctrl+C to stop\n")
    
    # Start preview window if requested
    preview_window = None
    if show_preview and TKINTER_AVAILABLE:
        preview_window = PreviewWindow(selected_window)
        preview_window.start()
        print("Preview window opened - you can see what's being captured")
        time.sleep(2)  # Give preview window time to start
    elif show_preview:
        print("Preview window unavailable - tkinter not installed")
    
    # Create screenshots directory (clear if exists)
    screenshots_dir = "screenshots"
    if os.path.exists(screenshots_dir):
        print(f"Clearing existing screenshots directory...")
        shutil.rmtree(screenshots_dir)
    os.makedirs(screenshots_dir)
    print(f"Screenshots will be saved to: {screenshots_dir}/")
    
    # Initialize CSV with headers
    save_table_to_csv([target_headers], output_csv)
    
    # Activate the selected window at start
    if selected_window and isinstance(selected_window, str):
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Activating target window...")
        if activate_window_macos(selected_window):
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Window activated successfully")
        else:
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Warning: Could not activate window")
        time.sleep(1)  # Give window time to come to front
    
    screenshot_count = 0
    
    try:
        while True:
            screenshot_count += 1
            timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
            screenshot_path = os.path.join(screenshots_dir, f"screenshot_{timestamp}.png")
            
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Taking screenshot {screenshot_count}...")
            
            try:
                # Take screenshot
                take_screenshot(screenshot_path, selected_window)
                
                # Always perform OCR to get table data
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Processing with OCR...")
                ocr_response = perform_ocr_on_image(
                    api_key,
                    model_name,
                    screenshot_path
                )
                
                # Extract table with target headers using structured output
                table_data = extract_table_with_headers(ocr_response, target_headers, api_key)
                
                if table_data and len(table_data) > 1:  # Has data beyond headers
                    # Always append new data to CSV (skip header row)
                    with open(output_csv, "a", newline="", encoding="utf-8") as f:
                        writer = csv.writer(f)
                        for row in table_data[1:]:  # Skip header
                            writer.writerow(row)
                    
                    data_rows = len(table_data) - 1
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] ✓ Extracted {data_rows} rows")
                else:
                    print(f"[{datetime.now().strftime('%H:%M:%S')}] No table data found")
                    
            except Exception as e:
                print(f"[{datetime.now().strftime('%H:%M:%S')}] Error: {e}")
                # Keep screenshot on error for debugging, but don't update previous_screenshot
            
            # Wait for next interval
            print(f"[{datetime.now().strftime('%H:%M:%S')}] Waiting {wait_time} seconds...")
            
            # Send arrow down keystrokes to scroll to next entries
            if selected_window and isinstance(selected_window, str):
                # Ensure window is still active
                activate_window_macos(selected_window)
                time.sleep(0.5)  # Brief pause to ensure window is active
                send_arrow_keys(arrow_strokes)  # Send configurable arrow down keystrokes
            
            print()  # Add newline for better formatting
            time.sleep(wait_time)
            
    except KeyboardInterrupt:
        print(f"\n[{datetime.now().strftime('%H:%M:%S')}] Stopping screen capture mode...")
        
        # Debug mode: pause before deduplication
        if debug_mode:
            print(f"\n[DEBUG MODE] Raw data captured to: {output_csv}")
            print("You can now review the raw CSV file before deduplication.")
            print("The file contains all captured data including potential duplicates.")
            
            while True:
                response = input("\nProceed with deduplication? (y/n): ").strip().lower()
                if response in ['y', 'yes']:
                    break
                elif response in ['n', 'no']:
                    print("Skipping deduplication. Raw data preserved.")
                    print(f"Final output saved to: {output_csv}")
                    # Stop preview window
                    if preview_window:
                        preview_window.stop()
                    return
                else:
                    print("Please enter 'y' for yes or 'n' for no.")
        
        # Deduplicate the CSV file
        print(f"[{datetime.now().strftime('%H:%M:%S')}] Deduplicating CSV file...")
        deduplicate_csv(output_csv)
        
        print(f"Final output saved to: {output_csv}")
        
        # Stop preview window
        if preview_window:
            preview_window.stop()

def main():
    parser = argparse.ArgumentParser(description='Screen capture OCR table extractor')
    parser.add_argument('--output', type=str, default="screen_capture_table.csv",
                        help='Path to save the extracted table CSV')
    parser.add_argument('--model', type=str, default=OCR_MODEL_NAME,
                        help=f'Mistral OCR model name (default: {OCR_MODEL_NAME})')
    parser.add_argument('--interval', type=int, default=10,
                        help='Screenshot interval in seconds (default: 10)')
    parser.add_argument('--api-key', type=str, default=None,
                        help='Mistral API key (overrides environment variable)')
    parser.add_argument('--no-preview', action='store_true',
                        help='Disable preview window')
    parser.add_argument('--debug', action='store_true',
                        help='Enable debug mode - pause before CSV deduplication')
    
    args = parser.parse_args()
    
    # Handle API key
    api_key = args.api_key or MISTRAL_API_KEY
    
    if not api_key:
        print("Error: Mistral API key not found.")
        print("Please provide it using one of these methods:")
        print("1. Command-line argument: --api-key YOUR_API_KEY")
        print("2. Environment variable: export MISTRAL_API_KEY=YOUR_API_KEY")
        if DOTENV_AVAILABLE:
            print("3. .env file: Create a .env file with MISTRAL_API_KEY=YOUR_API_KEY")
        return 1
    
    # Check required dependencies
    if not MISTRAL_SDK_AVAILABLE:
        print("Error: Mistral SDK required for OCR functionality.")
        print("Install with: pip install mistralai")
        return 1
        
    if not MSS_AVAILABLE and not PYAUTOGUI_AVAILABLE:
        print("Error: Screen capture requires mss or pyautogui package.")
        print("Install with: pip install mss")
        print("Or: pip install pyautogui")
        return 1
    
    # Check image comparison availability
    if not PIL_AVAILABLE:
        print("Warning: Pillow not installed - image comparison disabled.")
        print("All screenshots will be processed. Install with: pip install Pillow")
    
    # Get user-defined column headers
    target_headers = get_user_column_headers()
    
    # Get navigation settings
    print(f"\n=== Navigation Settings ===")
    
    # Wait time configuration
    wait_time_input = input(f"Enter wait time between captures in seconds (default: 5): ").strip()
    try:
        wait_time = int(wait_time_input) if wait_time_input else 5
        if wait_time < 1:
            print("Wait time must be at least 1 second. Using default of 5 seconds.")
            wait_time = 5
    except ValueError:
        print("Invalid wait time. Using default of 5 seconds.")
        wait_time = 5
    
    # Arrow strokes configuration
    arrow_strokes_input = input(f"Enter number of down arrow key strokes per cycle (default: 11): ").strip()
    try:
        arrow_strokes = int(arrow_strokes_input) if arrow_strokes_input else 11
        if arrow_strokes < 0:
            print("Arrow strokes cannot be negative. Using default of 11.")
            arrow_strokes = 11
    except ValueError:
        print("Invalid arrow strokes count. Using default of 11.")
        arrow_strokes = 11
    
    print(f"Navigation configured: {wait_time}s wait, {arrow_strokes} arrow strokes")
    
    # Window selection
    selected_window = None
    if sys.platform == 'darwin' or PYGETWINDOW_AVAILABLE:
        selected_window = select_window()
    else:
        print("Window selection unavailable. Using full screen capture.")
        print("To enable window selection, install: pip install pygetwindow")
    
    # Start screen capture mode
    show_preview = not args.no_preview
    debug_mode = args.debug
    screen_capture_mode(api_key, args.model, target_headers, args.output, args.interval, selected_window, show_preview, debug_mode, wait_time, arrow_strokes)
    return 0

if __name__ == "__main__":
    sys.exit(main())