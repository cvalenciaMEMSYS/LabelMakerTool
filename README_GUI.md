# Zebra Label Printer - GUI Application (USB Focus)

A simple graphical interface for printing MEMSYS sensor labels with custom serial numbers using a USB-connected Zebra printer.

## Table of Contents
- [Overview](#overview)
- [Features](#features)
- [Requirements](#requirements)
- [Installation](#installation)
- [Usage](#usage)
- [How It Works](#how-it-works)
- [Recreating This Program](#recreating-this-program)
- [Troubleshooting](#troubleshooting)
- [File Structure](#file-structure)

---

## Overview

This application provides a user-friendly GUI for printing labels from a Zebra thermal transfer printer. Instead of learning ZebraDesigner, operators simply:
1. Select the printer
2. Type the serial number
3. Click "Print" (or press Enter)

The app handles all the ZPL code manipulation and USB communication automatically.

---

## Features

✅ **Graphical User Interface** - No command line knowledge needed  
✅ **USB Printer Support** - Works with Zebra printers connected via USB  
✅ **Serial Number Entry** - Simple text input with validation  
✅ **Quantity Control** - Print 1 to 999 labels with +/- buttons  
✅ **Printer Selection** - Dropdown list of all Windows printers  
✅ **Template Loading** - Load any PRN file as the label template  
✅ **Real-time Status** - Visual feedback during printing  
✅ **Keyboard Shortcuts** - Press Enter to print, +/- for quantity  

---

## Requirements

### Software
- **Python 3.6 or higher** (3.8+ recommended)
- **Windows OS** (required for USB printing via win32print)
- **pywin32** Python package (for USB printer communication)

### Hardware
- Zebra thermal transfer printer (e.g., ZD421T, ZD620T, ZT411)
- USB cable connecting printer to computer
- Printer must be installed in Windows (Control Panel → Devices and Printers)

### Files
- `zebra_print_gui.py` - Main application
- `Memsys_Zebra_label.prn` - Your ZPL label template (exported from ZebraDesigner)
- `requirements.txt` - Python dependencies

---

## Installation

### Step 1: Install Python

1. Download Python from https://www.python.org/downloads/
2. Run installer
3. ✅ **IMPORTANT**: Check "Add Python to PATH" during installation
4. Click "Install Now"
5. Verify installation:
   ```bash
   python --version
   ```
   Should show: `Python 3.x.x`

### Step 2: Install Dependencies

Open Command Prompt or PowerShell and run:

```bash
pip install pywin32
```

Or use the requirements file:

```bash
pip install -r requirements.txt
```

### Step 3: Setup Printer

1. Connect Zebra printer via USB
2. Install printer driver (from Zebra website or Windows Update)
3. Verify printer appears in **Control Panel → Devices and Printers**
4. Print a test page to confirm it works

### Step 4: Prepare Template File

1. Open your label design in ZebraDesigner
2. Export as PRN file: **File → Export → Print to File (.prn)**
3. Save as `Memsys_Zebra_label.prn`
4. Place in the same folder as `zebra_print_gui.py`

---

## Usage

### Starting the Application

**Method 1: Double-click**
- Double-click `zebra_print_gui.py`
- Windows will open it with Python

**Method 2: Command line**
```bash
python zebra_print_gui.py
```

**Method 3: Create a shortcut (recommended for production)**
1. Right-click `zebra_print_gui.py`
2. Send to → Desktop (create shortcut)
3. Rename shortcut to "Print Labels"
4. Operators can now double-click the desktop icon

### Application Interface

```
┌────────────────────────────────────────────┐
│      MEMSYS Zebra Label Printer            │
├────────────────────────────────────────────┤
│ Printer Selection                          │
│ [Zebra ZD421-203dpi ZPL ▼] [Refresh]      │
│                                            │
│ Serial Number                              │
│ [                                    ]     │
│                                            │
│ Quantity                                   │
│      [ - ]    [ 1 ]    [ + ]              │
│                                            │
│ ┌────────────────────────────────────┐    │
│ │        PRINT LABELS                │    │
│ └────────────────────────────────────┘    │
│                                            │
│ Status: Ready                              │
│                                            │
│ [Load Template]              [Help]        │
└────────────────────────────────────────────┘
```

### Printing Labels

1. **Select Printer**: Choose your Zebra printer from dropdown
2. **Enter Serial**: Type serial number (e.g., "EWS-B-001")
3. **Set Quantity**: Use +/- buttons or leave at 1
4. **Print**: Click "PRINT LABELS" or press Enter
5. **Confirm**: Dialog asks for confirmation
6. **Done**: Labels print, serial field clears for next entry

### Keyboard Shortcuts

- **Enter** - Print labels (when serial number field has focus)
- **Tab** - Move between fields
- **Esc** - Cancel dialogs

---

## How It Works

### Technical Architecture

```
User Input (GUI)
     ↓
Serial Number + Quantity
     ↓
Load PRN Template File
     ↓
Replace XXXXXXX1 with Serial Number in ZPL Code
     ↓
Send ZPL to Printer via USB (win32print)
     ↓
Printer Prints Labels
```

### ZPL Code Manipulation

The PRN file contains ZPL (Zebra Programming Language) code like this:

```zpl
^XA
^FO100,50^A0N,30,30^FDMEMSYS B.V.^FS
^FO100,100^A0N,30,30^FDEWS-B v0.2^FS
^FO100,150^A0N,30,30^FDS/N: XXXXXXX1^FS
^XZ
```

The application:
1. Reads this template
2. Finds the line with `FDXXXXXXX1^FS`
3. Replaces `XXXXXXX1` with the user's input (e.g., `EWS-B-042`)
4. Sends modified ZPL to printer

Result:
```zpl
^FO100,150^A0N,30,30^FDS/N: EWS-B-042^FS
```

### USB Communication

Uses Windows win32print API:
1. `OpenPrinter()` - Connect to printer
2. `StartDocPrinter()` - Begin print job
3. `WritePrinter()` - Send ZPL code as raw data
4. `EndDocPrinter()` - Complete job
5. `ClosePrinter()` - Disconnect

---

## Recreating This Program

### Using VS Code with Claude Opus via Copilot

This section provides detailed prompts and context to recreate this application using Claude Opus in VS Code Copilot.

#### Initial Setup Prompt

```
I need to create a Python GUI application for printing Zebra labels via USB. 
The application should:

1. Use tkinter for the GUI (built into Python, no extra dependencies)
2. Have a modern, clean interface with:
   - A header with the title "MEMSYS Zebra Label Printer"
   - Dropdown to select from Windows printers
   - Text entry field for serial number
   - +/- buttons to set quantity (1-999)
   - Large "PRINT LABELS" button
   - Status label showing current state
   - "Load Template" and "Help" buttons at bottom

3. USB Printing functionality:
   - Use win32print library (from pywin32 package)
   - Read a PRN file containing ZPL code
   - Replace placeholder text "XXXXXXX1" with user's serial number
   - Send ZPL to selected printer using RAW mode
   - Print specified quantity of labels

4. Error handling:
   - Validate serial number is not empty
   - Validate printer is selected
   - Show confirmation dialog before printing
   - Show error messages if something fails
   - Provide helpful error messages

File structure:
- zebra_print_gui.py (main application)
- Memsys_Zebra_label.prn (ZPL template file)

The target users are production workers with no technical background, 
so the interface should be extremely simple and intuitive.
```

#### Code Structure Prompts

**For the main application class:**
```
Create a Python class called ZebraLabelPrinter that:
- Inherits from nothing (tkinter doesn't require inheritance)
- Takes a tk.Tk() root window as __init__ parameter
- Has instance variables for:
  - prn_file (path to template)
  - zpl_template (loaded ZPL code string)
  - printer_name (selected printer)
  - win32print module reference
  - usb_available flag (True if win32print imported successfully)
- Has methods:
  - create_widgets() - builds entire GUI
  - load_template() - reads PRN file
  - refresh_printers() - gets list of Windows printers
  - replace_serial_number() - modifies ZPL code
  - send_to_usb_printer() - sends to printer
  - print_label() - main print workflow
```

**For the GUI layout:**
```
Create a tkinter GUI with this layout structure:

1. Header frame (bg="#2C3E50", height=60px):
   - White bold title text "MEMSYS Zebra Label Printer"

2. Content frame (padx=20, pady=20):
   
   a. Printer Selection LabelFrame:
      - Combobox (dropdown) showing available printers
      - "Refresh" button next to it
   
   b. Serial Number LabelFrame:
      - Entry widget, font size 14, width 30 chars
      - Bind Enter key to trigger print
   
   c. Quantity LabelFrame:
      - Horizontal layout with:
        - "-" button (red, decrease quantity)
        - Label showing current number (large, bold)
        - "+" button (green, increase quantity)
   
   d. Print Button:
      - Large button "PRINT LABELS"
      - Green background, white text
      - Height 2, bold font size 14
   
   e. Status Label:
      - Shows current status/errors
      - Color changes: green=success, red=error, blue=working
   
   f. Bottom buttons:
      - "Load Template File" (left)
      - "Help" (right)

Use these colors:
- Primary green: #27AE60
- Primary blue: #3498DB
- Primary red: #E74C3C
- Dark header: #2C3E50
- Light gray: #95A5A6
```

**For USB printing:**
```
Create a function to send ZPL code to a Windows USB printer:

def send_to_usb_printer(zpl_code, printer_name):
    """
    Sends ZPL code to printer using win32print API
    
    Args:
        zpl_code (str): Complete ZPL code to print
        printer_name (str): Windows printer name from EnumPrinters
    
    Steps:
    1. Import win32print (from pywin32 package)
    2. Open printer handle: win32print.OpenPrinter(printer_name)
    3. Start document: win32print.StartDocPrinter(handle, 1, ("Label", None, "RAW"))
       - Use "RAW" mode to send ZPL directly without Windows processing
    4. Start page: win32print.StartPagePrinter(handle)
    5. Write data: win32print.WritePrinter(handle, zpl_code.encode('utf-8'))
    6. End page: win32print.EndPagePrinter(handle)
    7. End document: win32print.EndDocPrinter(handle)
    8. Close printer: win32print.ClosePrinter(handle)
    
    Error handling:
    - Wrap in try/finally to ensure printer is always closed
    - Catch and re-raise exceptions with helpful messages
    
    Returns: True on success, raises Exception on failure
    """
```

**For getting Windows printers:**
```
Create a function to get list of installed Windows printers:

def get_windows_printers():
    """
    Returns list of printer names from Windows
    
    Uses win32print.EnumPrinters(2) to get local printers
    - Flag 2 = PRINTER_ENUM_LOCAL (local printers only)
    - Returns list of tuples: (flags, description, name, comment)
    - Extract the name field (index 2) from each tuple
    
    Also get default printer:
    - win32print.GetDefaultPrinter()
    
    Returns:
        list of str: Printer names
        str: Default printer name
    """
```

**For ZPL code manipulation:**
```
Create a function to replace serial number in ZPL code:

def replace_serial_number(zpl_template, new_serial):
    """
    Replace placeholder serial number in ZPL code
    
    Args:
        zpl_template (str): Original ZPL code with placeholder
        new_serial (str): Serial number to insert
    
    Implementation:
    - Use str.replace() to find and replace
    - Search for: 'FDXXXXXXX1^FS'
    - Replace with: f'FD{new_serial}^FS'
    
    ZPL field structure:
    - ^FD = Field Data (start of text content)
    - XXXXXXX1 = Placeholder text
    - ^FS = Field Separator (end of text content)
    
    Returns:
        str: Modified ZPL code with serial number
    """
```

#### Advanced Features Prompts

**For template file browsing:**
```
Add a file dialog to select PRN template:
- Use tkinter.filedialog.askopenfilename()
- Filter: [("PRN files", "*.prn"), ("All files", "*.*")]
- Update self.prn_file with selected path
- Call load_template() to reload
- Update status label with new filename
```

**For validation:**
```
Add input validation before printing:
1. Check template is loaded (self.zpl_template not None)
2. Check serial number is not empty (after strip())
3. Check printer is selected and valid
4. Check quantity is between 1 and 999
5. Show error messagebox if any validation fails
6. Return early if validation fails
```

**For user feedback:**
```
Add visual feedback during printing:
1. Disable print button (set state=DISABLED)
2. Update status label to "Printing..." with blue color
3. Call root.update() to refresh GUI
4. Loop through quantity, updating status for each label
5. Show success messagebox when complete
6. Re-enable print button in finally block
7. Clear serial entry and set focus back to it
```

#### Testing Prompts

```
Create test scenarios for the application:

1. Test with missing pywin32:
   - Mock ImportError for win32print
   - Verify error message shown
   - Verify printers dropdown shows helpful message

2. Test with no printers:
   - Mock EnumPrinters returning empty list
   - Verify "No printers found" message

3. Test serial validation:
   - Empty string → error dialog
   - Valid serial → proceeds to print

4. Test file loading:
   - Missing PRN file → error status
   - Valid PRN file → success status, template loaded

5. Test printing workflow:
   - Mock win32print calls
   - Verify ZPL code has serial number replaced
   - Verify correct printer name used
   - Verify quantity loop works
```

---

## Troubleshooting

### "pywin32 not installed" Error

**Problem**: Application shows "pywin32 not installed - see README"

**Solution**:
```bash
pip install pywin32
```

After installation, close and restart the application.

### "No printers found" Error

**Problem**: Printer dropdown shows "No printers found"

**Solutions**:
1. Check printer is connected via USB
2. Verify printer is installed in Windows:
   - Control Panel → Devices and Printers
   - Should see your Zebra printer
3. Install Zebra printer driver from https://www.zebra.com/
4. Click "Refresh" button in the application

### Template File Not Found

**Problem**: Status shows "Template file not found: Memsys_Zebra_label.prn"

**Solutions**:
1. Export your label from ZebraDesigner as .prn file
2. Save it as `Memsys_Zebra_label.prn`
3. Place in same folder as `zebra_print_gui.py`
4. Or click "Load Template File" and browse to it

### Serial Number Not Changing

**Problem**: Printed labels all have "XXXXXXX1" instead of entered serial

**Solutions**:
1. Check your PRN file contains exactly `FDXXXXXXX1^FS`
2. If placeholder is different, edit line 147 in the script:
   ```python
   modified_zpl = zpl_code.replace('FDYOURPLACEHOLDER^FS', f'FD{new_serial}^FS')
   ```

### Labels Don't Print

**Problem**: Click print but nothing happens

**Solutions**:
1. Check printer is turned on
2. Check USB cable is connected
3. Check printer has paper and ribbon loaded
4. Print a test page from Windows to verify printer works
5. Check printer status lights - should be solid green
6. Try calibrating printer (hold FEED button until it beeps)

### "Access Denied" Error

**Problem**: win32print shows access denied when printing

**Solutions**:
1. Check printer is not paused in Windows
2. Check no other jobs are stuck in queue
3. Clear print queue: Control Panel → Printers → right-click printer → "See what's printing" → Cancel all
4. Restart print spooler service:
   - Services → Print Spooler → Restart

---

## File Structure

```
zebra_label_printer/
├── zebra_print_gui.py          # Main GUI application
├── requirements.txt             # Python dependencies
├── README.md                    # This file
├── Memsys_Zebra_label.prn      # ZPL template (your label design)
└── docs/                        # Optional: additional documentation
    ├── screenshots/             # GUI screenshots
    └── zpl_reference.md         # ZPL code reference
```

### zebra_print_gui.py Structure

```python
# Imports
import tkinter
import win32print
import os, sys

# Main Application Class
class ZebraLabelPrinter:
    def __init__(root)           # Initialize GUI and variables
    def create_widgets()         # Build GUI layout
    def load_template()          # Read PRN file
    def refresh_printers()       # Get Windows printer list
    def replace_serial_number()  # Modify ZPL code
    def send_to_usb_printer()    # Send to printer via USB
    def print_label()            # Main print workflow
    def show_help()              # Help dialog
    # ... other helper methods

# Main entry point
if __name__ == "__main__":
    root = tk.Tk()
    app = ZebraLabelPrinter(root)
    root.mainloop()
```

---

## Additional Resources

### ZPL Programming

- **Zebra Programming Guide**: https://www.zebra.com/content/dam/zebra/manuals/printers/common/programming/zpl-zbi2-pm-en.pdf
- **Online ZPL Viewer**: http://labelary.com/viewer.html
- **ZPL Command Reference**: Search for "ZPL commands" on zebra.com

### Python Libraries

- **tkinter Documentation**: https://docs.python.org/3/library/tkinter.html
- **pywin32 Documentation**: https://pypi.org/project/pywin32/
- **win32print Examples**: Search "python win32print examples"

### Zebra Printer Support

- **Driver Downloads**: https://www.zebra.com/support
- **Knowledge Base**: https://supportcommunity.zebra.com/
- **Printer Manuals**: Search your printer model on zebra.com

---

## License

This is internal MEMSYS software for production use.

## Support

For issues or questions:
1. Check this README thoroughly
2. Check printer is working (print test page from Windows)
3. Verify pywin32 is installed: `pip list | findstr pywin32`
4. Check PRN template file exists and is valid ZPL code

---

## Version History

**v1.0** - Initial GUI version with USB support
- Tkinter GUI interface
- USB printing via win32print
- Serial number entry with validation
- Quantity control
- Template file loading
- Status feedback

---

**End of README**
