# Screen Capture OCR to CSV

An automated Python script that captures screenshots intermittently, uses Mistral's OCR API to extract table data, and saves it as CSV with customizable column headers. Perfect for data entry automation and extracting structured data from applications.

## Features

- **Automated Screen Capture**: Intermittent screenshot capture with configurable intervals
- **Smart OCR Processing**: Uses Mistral OCR API for accurate table data extraction
- **Structured Output**: Uses Mistral Medium with JSON mode for proper column formatting
- **Window Selection**: Target specific application windows (macOS and Windows)
- **Navigation Automation**: Configurable arrow key automation for scrolling through data
- **Preview Window**: Real-time preview of captured screenshots (optional)
- **Deduplication**: Smart duplicate removal with debug mode for review
- **Cross-Platform**: Works on macOS, Windows, and Linux

## Prerequisites

### Required
- Python 3.7+
- Mistral API key (get one at [console.mistral.ai](https://console.mistral.ai))

### System Requirements
- **macOS**: Requires Screen Recording permission for screenshot capture
- **Windows/Linux**: No special permissions required

## Installation

1. **Clone or download the script**
   ```bash
   git clone <repository-url>
   cd screen-cap-ocr
   ```

2. **Install dependencies**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up your API key** (choose one method):
   
   **Option A: Environment Variable**
   ```bash
   export MISTRAL_API_KEY="your-api-key-here"
   ```
   
   **Option B: .env File**
   ```bash
   echo "MISTRAL_API_KEY=your-api-key-here" > .env
   ```
   
   **Option C: Command Line**
   ```bash
   python screen_capture_ocr.py --api-key your-api-key-here --output data.csv
   ```

## Usage

### Basic Usage

```bash
python screen_capture_ocr.py --output my_data.csv
```

### Command Line Options

```bash
python screen_capture_ocr.py [OPTIONS]

Options:
  --output PATH        Output CSV file path (default: screen_capture_table.csv)
  --model MODEL        Mistral OCR model (default: pixtral-12b-2409)
  --interval SECONDS   Screenshot interval (default: 10 seconds)
  --api-key KEY        Mistral API key (overrides environment variable)
  --no-preview         Disable preview window
  --debug              Enable debug mode (pause before deduplication)
  -h, --help          Show help message
```

### Interactive Setup

When you run the script, it will guide you through an interactive setup:

1. **Column Headers**: Define your target CSV columns
   ```
   Enter column header 1: Person Name
   Enter column header 2: Company Name  
   Enter column header 3: Job Title
   Enter column header 4: (press Enter to finish)
   ```

2. **Navigation Settings**: Configure automation behavior
   ```
   Enter wait time between captures in seconds (default: 5): 3
   Enter number of down arrow key strokes per cycle (default: 11): 15
   ```

3. **Window Selection**: Choose target application window
   ```
   Available windows:
   1. Safari - LinkedIn
   2. Chrome - Company Directory
   3. Full screen
   Select window (1-3): 1
   ```

4. **Start Capture**: The script begins automated data collection

### Example Workflow

1. **Prepare your data source**: Open the application/website with tabular data
2. **Position the view**: Navigate to the starting point of your data
3. **Run the script**: `python screen_capture_ocr.py --output contacts.csv`
4. **Configure columns**: Set up headers like "Name", "Email", "Phone"
5. **Set navigation**: Configure timing and scroll behavior  
6. **Select window**: Choose your data source window
7. **Let it run**: The script automatically captures, processes, and scrolls
8. **Stop when done**: Press Ctrl+C to finish and deduplicate results

## How It Works

1. **Screenshot Capture**: Takes periodic screenshots of selected window/screen
2. **OCR Processing**: Sends images to Mistral OCR API for text extraction
3. **Structured Formatting**: Uses Mistral Medium with JSON mode to format data into your target columns
4. **CSV Output**: Appends properly formatted rows to your output file
5. **Navigation**: Sends arrow key strokes to scroll to next data entries
6. **Deduplication**: Removes duplicate entries at the end (with optional debug review)

## Advanced Features

### Debug Mode
```bash
python screen_capture_ocr.py --output data.csv --debug
```
Pauses before final deduplication, allowing you to review raw captured data.

### Custom Navigation
- **Wait Time**: Control delay between capture cycles (1-60 seconds recommended)
- **Arrow Strokes**: Number of down arrow keys sent per cycle (adjust based on your data density)

### Window Selection
- **macOS**: Uses AppleScript for robust window management
- **Windows/Linux**: Uses PyGetWindow library
- **Fallback**: Full screen capture if window selection unavailable

## Troubleshooting

### macOS Permissions
If you get permission errors:
1. Go to System Preferences → Security & Privacy → Screen Recording
2. Add Terminal (or your terminal app) to the allowed apps
3. Restart your terminal

### OCR Not Working
- Verify your Mistral API key is correct
- Check your account has billing enabled at console.mistral.ai
- Ensure you have API credits available

### Window Selection Issues
- Try selecting "Full screen" option
- Install PyGetWindow: `pip install pygetwindow`
- On macOS, the script uses AppleScript as fallback

### Poor OCR Results
- Ensure your data source has clear, readable text
- Try adjusting the window size for better text visibility
- Consider using structured output mode (automatically enabled)

## File Structure

```
screen-cap-ocr/
├── screen_capture_ocr.py    # Main script
├── requirements.txt         # Python dependencies
├── README.md               # This file
├── screenshots/            # Captured screenshots (auto-created)
└── .env                   # API key file (optional)
```

## Dependencies

- `mistralai`: Mistral AI SDK for OCR and chat completion
- `mss`: Fast cross-platform screen capture
- `pyautogui`: Keyboard automation and fallback screen capture
- `Pillow`: Image processing and preview functionality
- `python-dotenv`: Environment variable management (optional)
- `pygetwindow`: Window management (Windows/Linux)

## Tips for Best Results

1. **Clean Data Source**: Ensure your data is displayed clearly in a table format
2. **Consistent Layout**: Keep the same table structure throughout capture
3. **Proper Timing**: Adjust wait time based on your application's response speed
4. **Window Focus**: Keep the target window in focus during capture
5. **Test First**: Run a short test with 2-3 entries to verify column mapping
6. **Monitor Progress**: Watch the console output to ensure proper data extraction

## License

This project is provided as-is for educational and automation purposes. Please ensure you comply with the terms of service of any applications or websites you're extracting data from.