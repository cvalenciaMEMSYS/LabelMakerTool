# MEMSYS Label Printer Tool

A cross-platform GUI application for printing MEMSYS sensor labels via USB to a Zebra thermal printer. Works on Windows and Linux (including Raspberry Pi).

## Overview

This tool is designed for the **Zebra ZD421T 300 DPI** thermal transfer printer. It supports two label sizes:

| Product | Label Size | Template File |
|---------|-----------|---------------|
| EWS | 38×19mm | `Memsys EWS Zebra label.prn` |
| Gateway | 57×32mm | `Memsys GW Zebra label.prn` |

## Features

- **Two printing modes**:
  - **Template mode**: Load a PRN template file — serial placeholder is auto-detected from `S/N:` pattern
  - **Text-only mode**: Print dynamic plain text labels with automatic font sizing
- **Smart placeholder detection**: Automatically finds the serial number placeholder in the ZPL template. Falls back to a visual editor if no `S/N:` pattern is found
- **Clipboard auto-paste**: Monitors the system clipboard for new serial numbers (8-char → EWS with dash, 6-char → Gateway as-is)
- **Auto-print**: When enabled, prints one label per new serial number detected on the clipboard
- Template selection via file browser
- Serial number entry with placeholder replacement
- Quantity control (1-999 labels)
- Printer selection dropdown with refresh capability
- **Text-only auto-sizing** (word-aware):
  - Each word prints on its own centered line
  - Font size based on longest word: ≤10 chars → large, 11-12 → medium, 13+ → small
  - Up to 3 lines
- **Reprint**: quick reprint of the last 5 labels from dropdown
- **Label size detection**: auto-reads template dimensions and prompts for confirmation
- **Cross-platform**: Windows (win32print) and Linux/RPi (CUPS)
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

## Swapping Label Rolls

When switching between label sizes (e.g. EWS 38×19mm ↔ Gateway 57×32mm):

1. **Open the cover** — pull the latch forward to release
2. **Open the yellow media guides** — they are spring-loaded
3. **Remove the current roll** and insert the new one between the guides
4. **Release the guides** — they self-center on the new roll
5. **Feed labels** through the media path and under the printhead
6. **Check the feed sensor** — the yellow square under the printhead must be centered on the labels
7. **Close the cover**
8. **Calibrate** — press and hold **PAUSE + CANCEL** for 2 seconds. The printer feeds several labels and auto-detects the new size
9. **Print a test label** to confirm alignment

> When loading a template, the app detects the label dimensions and asks you to confirm the correct labels are loaded.

## Supported Label Types

| Label Type | Label Size | Mode | Status |
|------------|-----------|------|--------|
| EWS labels | 38×19mm | Template (PRN with serial number) | Supported |
| Gateway labels | 57×32mm | Template (PRN with serial number) | Supported |
| Simple text labels | Any | Text-only mode | Supported |
| Ethernet connectivity | - | - | Planned |

## Installation

### Requirements

- Python 3.8+
- Zebra ZD421T printer with drivers installed
- **Windows**: `pywin32` for printer access
- **Linux / Raspberry Pi**: Either `pycups` (CUPS) or direct USB (zero dependencies)

### Setup (Windows)

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

### Setup (Linux / Raspberry Pi)

#### Option A: Direct USB (simplest — no extra packages)

1. Connect the Zebra printer via USB
2. Add yourself to the `lp` group (one-time):
   ```bash
   sudo usermod -a -G lp $USER
   ```
3. Log out and back in
4. Run the application:
   ```bash
   python3 zebra_print_gui.py
   ```

The printer appears as `/dev/usb/lp0` in the dropdown. No CUPS, no pip packages — just Python + USB cable.

#### Option B: Via CUPS (for managed setups)

1. Install CUPS and development headers:
   ```bash
   sudo apt install cups libcups2-dev
   ```

2. Add the Zebra printer as a **Raw Queue** in CUPS:
   ```bash
   sudo lpadmin -p ZebraZD421 -E -v usb://Zebra/ZD421 -m raw
   ```
   (Adjust the URI — use `lpinfo -v` to find the correct USB path.)

3. Install pycups:
   ```bash
   pip install pycups
   ```

4. Run the application:
   ```bash
   python3 zebra_print_gui.py
   ```

## Usage

### Printing with a Template (EWS / Gateway Labels)

1. Click **Browse** and select your PRN template file (`Memsys EWS Zebra label.prn` or `Memsys GW Zebra label.prn`)
2. The app detects the label size and asks you to confirm the correct labels are loaded
3. The serial placeholder is auto-detected from the `S/N:` pattern
4. Enter the serial number
5. Set the quantity
5. Select the printer from the dropdown
6. Click **PRINT LABELS** or press **Enter**
7. Check the status field for confirmation

### Printing Text-Only Labels

1. Check the **Text-only label** checkbox (template section hides)
2. Enter your text in the Label Text field
3. Set the quantity
4. Select the printer
5. Click **PRINT LABELS** or press **Enter**

Text is automatically sized based on word length for optimal readability.

### Reprinting a Recent Label

1. Select a label from the **Recent Labels** dropdown (stores last 5 prints)
2. Click **Reprint** to send a single identical copy to the printer

### Auto-Paste and Auto-Print

1. Enable **Auto-paste from clipboard** — the app watches the clipboard every 500ms
2. Copy a serial number in another application:
   - **8 characters** (e.g. `ABCD1234`) → auto-formatted as `ABCD-1234` (EWS)
   - **6 characters** (e.g. `ABC123`) → pasted as-is (Gateway)
3. The serial number auto-fills the entry field (only when empty)
4. Optionally enable **Auto-print on paste** — prints one label per new serial automatically
   - Duplicate guard: same serial copied twice only prints once
   - Entry clears after printing, ready for the next serial

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

- The app auto-detects the serial placeholder from the `S/N:` pattern in the template
- If no `S/N:` is found, a viewer opens to let you select the placeholder text manually
- If the template has `S/N:` with no text after it, the serial is inserted directly after the colon

### Text-only labels look wrong

- Text is auto-sized: long text uses smaller fonts with word wrapping
- For best results, keep text under 20 characters
- Check printer calibration if alignment is off

## Dependencies

- `tkinter` - GUI framework (included with Python)
- `pywin32` - Windows printer communication via USB
- `pycups` - Linux printer communication via CUPS (optional — direct USB works without it)
