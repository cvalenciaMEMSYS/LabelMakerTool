# MEMSYS Label Printer Tool

A Windows GUI application for printing MEMSYS sensor labels via USB to a Zebra thermal printer.

## Overview

This tool is designed for the **Zebra ZD421T 300 DPI** thermal transfer printer with **38x19mm labels**. It provides a simple interface for printing serial-numbered sensor labels and plain text labels.

## Features

- **Two printing modes**:
  - **Template mode**: Load a PRN template file and replace the serial number placeholder (`XXXX-XXXX` pattern)
  - **Text-only mode**: Print dynamic plain text labels with automatic font sizing
- Template selection via file browser
- Serial number entry with placeholder replacement
- Quantity control (1-999 labels)
- Printer selection dropdown with refresh capability
- **Text-only auto-sizing**:
  - 6 characters or less: large font, 1 line
  - 7-14 characters: medium font, 2 lines
  - 15+ characters: small font, 3 lines
- Immediate printing (no confirmation dialogs)
- Keyboard shortcut: Press Enter to print
- Status feedback in the interface

## Printer Setup

Before using this tool, the Zebra printer drivers must be installed:

1. **Get the drivers**: Download the Zebra ZD421T printer drivers from the company SharePoint drive
2. **Connect the printer**: Plug in the Zebra ZD421T via USB
3. **Install drivers**: Run the driver installer while the printer is connected
4. **Verify installation**: The printer should appear in Windows **Devices and Printers**
5. **Ready to use**: Launch this tool and select the printer from the dropdown

## Supported Label Types

| Label Type | Mode | Status |
|------------|------|--------|
| EWS labels | Template (PRN with serial number) | Supported |
| Simple text labels | Text-only mode | Supported |
| Gateway labels | Template | Planned |
| Ethernet connectivity | - | Planned |

## Installation

### Requirements

- Windows OS
- Python 3.8+
- Zebra ZD421T printer with drivers installed

### Setup

1. (Recommended) Create a virtual environment:
   ```
   python -m venv .venv
   .venv\Scripts\activate
   ```

2. Install dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python zebra_print_gui.py
   ```

## Usage

### Printing with a Template (EWS Labels)

1. Click **Browse** and select your PRN template file
2. Enter the serial number (replaces `XXXX-XXXX` in the template)
3. Set the quantity
4. Select the printer from the dropdown
5. Click **PRINT LABELS** or press **Enter**
6. Check the status field for confirmation

### Printing Text-Only Labels

1. Check the **Text-only label** checkbox (template section hides)
2. Enter your text in the Label Text field
3. Set the quantity
4. Select the printer
5. Click **PRINT LABELS** or press **Enter**

Text is automatically sized based on length for optimal readability on the 38x19mm label.

## Troubleshooting

### Printer not appearing in dropdown

- Ensure the Zebra drivers are installed (see Printer Setup)
- Check that the printer is connected via USB and powered on
- Click **Refresh** to update the printer list
- Verify the printer appears in Windows Devices and Printers

### Labels not printing

- Check that the correct printer is selected
- Ensure the printer has labels loaded and ribbon installed
- Check the status message for error details
- Try printing a Windows test page to verify the printer works

### Wrong serial number on labels

- Template files use `XXXX-XXXX` as the placeholder pattern
- Ensure your PRN template contains this exact placeholder
- The serial number you enter replaces all occurrences of `XXXX-XXXX`

### Text-only labels look wrong

- Text is auto-sized: long text uses smaller fonts with word wrapping
- For best results, keep text under 20 characters
- Check printer calibration if alignment is off

## Dependencies

- `tkinter` - GUI framework (included with Python)
- `pywin32` - Windows printer communication via USB
