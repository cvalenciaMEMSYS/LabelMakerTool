#!/usr/bin/env python3
"""
Zebra Label Printer - GUI Version with USB Support
Simple graphical interface for printing labels with custom serial numbers
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog
import os
import re
import sys
import tempfile
from pathlib import Path

# Platform-specific printing backends
_PLATFORM = "unsupported"
win32print = None
cups = None

if sys.platform == "win32":
    try:
        import win32print  # type: ignore[import-untyped,no-redef]
        _PLATFORM = "windows"
    except ImportError:
        pass
elif sys.platform.startswith("linux"):
    try:
        import cups  # type: ignore[import-untyped,no-redef]
        _PLATFORM = "linux"
    except ImportError:
        pass

_INVALID_PRINTERS = {
    "No printers found",
    "pywin32 not installed - see README",
    "pycups not installed - see README",
    "Unsupported platform",
}


class ZebraLabelPrinter:
    """Main application class for the Zebra Label Printer GUI"""
    
    def __init__(self, root):
        self.root = root
        self.root.title("MEMSYS - Zebra Label Printer")
        self.root.geometry("600x700")
        self.root.resizable(True, True)
        self.root.minsize(400, 450)
        
        # Configuration
        self.prn_file = "Memsys_Zebra_label.prn"
        self.zpl_template = None
        self.printer_name = None
        self.usb_available = _PLATFORM != "unsupported"
        self._recent_labels: list[dict] = []
        self._serial_placeholder: str | None = None
        self._serial_insert_pos: int | None = None
        self._last_clipboard = ""
        self._last_printed_serial = ""
        self._clipboard_poll_id: str | None = None
        
        # Create UI
        self.create_widgets()
        
        # Load template on startup
        self.load_template()
        
        # Load available printers
        self.refresh_printers()
    
    def create_widgets(self):
        """Create all GUI widgets"""
        
        # Header
        header_frame = tk.Frame(self.root, bg="#2C3E50", height=60)
        header_frame.pack(fill=tk.X)
        header_frame.pack_propagate(False)
        
        title_label = tk.Label(
            header_frame,
            text="MEMSYS Zebra Label Printer",
            font=("Arial", 16, "bold"),
            bg="#2C3E50",
            fg="white"
        )
        title_label.pack(pady=15)
        
        # Main content frame
        content_frame = tk.Frame(self.root, padx=20, pady=20)
        content_frame.pack(fill=tk.BOTH, expand=True)
        
        # Mode selection (template vs text-only)
        mode_frame = tk.Frame(content_frame)
        mode_frame.pack(fill=tk.X, pady=(0, 10))
        
        self.text_only_var = tk.BooleanVar(value=False)
        self.text_only_check = tk.Checkbutton(
            mode_frame,
            text="Text-only label (no template)",
            variable=self.text_only_var,
            command=self._on_mode_change,
            font=("Arial", 10)
        )
        self.text_only_check.pack(side=tk.LEFT)
        
        # Auto-paste from clipboard
        auto_frame = tk.Frame(content_frame)
        auto_frame.pack(fill=tk.X, pady=(0, 5))

        self.auto_paste_var = tk.BooleanVar(value=False)
        self.auto_paste_check = tk.Checkbutton(
            auto_frame,
            text="Auto-paste from clipboard",
            variable=self.auto_paste_var,
            command=self._on_auto_paste_toggle,
            font=("Arial", 10)
        )
        self.auto_paste_check.pack(side=tk.LEFT)

        # Auto-print on paste (indented under auto-paste)
        auto_print_frame = tk.Frame(content_frame)
        auto_print_frame.pack(fill=tk.X, pady=(0, 10))

        self.auto_print_var = tk.BooleanVar(value=False)
        self.auto_print_check = tk.Checkbutton(
            auto_print_frame,
            text="    Auto-print on paste (1 label per new serial)",
            variable=self.auto_print_var,
            font=("Arial", 9),
            state=tk.DISABLED
        )
        self.auto_print_check.pack(side=tk.LEFT)

        # Template selection
        self.template_frame = tk.LabelFrame(content_frame, text="Label Template", padx=10, pady=10)
        self.template_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.template_label = tk.Label(
            self.template_frame,
            text="No template loaded",
            font=("Arial", 9),
            fg="#7F8C8D",
            anchor=tk.W
        )
        self.template_label.pack(side=tk.LEFT, fill=tk.X, expand=True, padx=(0, 10))
        
        browse_btn = tk.Button(
            self.template_frame,
            text="Browse...",
            command=self.browse_template,
            bg="#8E44AD",
            fg="white",
            padx=10
        )
        browse_btn.pack(side=tk.RIGHT)
        
        # Printer selection
        printer_frame = tk.LabelFrame(content_frame, text="Printer Selection", padx=10, pady=10)
        printer_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.printer_var = tk.StringVar()
        self.printer_dropdown = ttk.Combobox(
            printer_frame,
            textvariable=self.printer_var,
            state="readonly",
            width=40
        )
        self.printer_dropdown.pack(side=tk.LEFT, padx=(0, 10))
        
        refresh_btn = tk.Button(
            printer_frame,
            text="Refresh",
            command=self.refresh_printers,
            bg="#3498DB",
            fg="white",
            padx=10
        )
        refresh_btn.pack(side=tk.LEFT)
        
        # Serial number / text input
        self.serial_frame = tk.LabelFrame(content_frame, text="Serial Number", padx=10, pady=10)
        self.serial_frame.pack(fill=tk.X, pady=(0, 15))
        
        self.serial_entry = tk.Entry(self.serial_frame, font=("Arial", 14), width=30)
        self.serial_entry.pack()
        self.serial_entry.focus()
        
        # Bind Enter key to print
        self.serial_entry.bind('<Return>', lambda e: self.print_label())
        
        # Quantity selection
        quantity_frame = tk.LabelFrame(content_frame, text="Quantity", padx=10, pady=10)
        quantity_frame.pack(fill=tk.X, pady=(0, 15))
        
        quantity_inner = tk.Frame(quantity_frame)
        quantity_inner.pack()
        
        self.quantity_var = tk.IntVar(value=1)
        
        minus_btn = tk.Button(
            quantity_inner,
            text="-",
            command=self.decrease_quantity,
            font=("Arial", 12, "bold"),
            width=3,
            bg="#E74C3C",
            fg="white"
        )
        minus_btn.pack(side=tk.LEFT, padx=5)
        
        quantity_label = tk.Label(
            quantity_inner,
            textvariable=self.quantity_var,
            font=("Arial", 16, "bold"),
            width=5
        )
        quantity_label.pack(side=tk.LEFT, padx=10)
        
        plus_btn = tk.Button(
            quantity_inner,
            text="+",
            command=self.increase_quantity,
            font=("Arial", 12, "bold"),
            width=3,
            bg="#27AE60",
            fg="white"
        )
        plus_btn.pack(side=tk.LEFT, padx=5)
        
        # Print button
        self.print_btn = tk.Button(
            content_frame,
            text="PRINT LABELS",
            command=self.print_label,
            font=("Arial", 14, "bold"),
            bg="#27AE60",
            fg="white",
            height=2,
            cursor="hand2"
        )
        self.print_btn.pack(fill=tk.X, pady=(0, 10))
        
        # Status label
        self.status_label = tk.Label(
            content_frame,
            text="Ready",
            font=("Arial", 10),
            fg="#7F8C8D"
        )
        self.status_label.pack()
        
        # Recent labels (reprint)
        reprint_frame = tk.LabelFrame(content_frame, text="Recent Labels", padx=10, pady=10)
        reprint_frame.pack(fill=tk.X, pady=(0, 10))

        self.reprint_var = tk.StringVar()
        self.reprint_dropdown = ttk.Combobox(
            reprint_frame,
            textvariable=self.reprint_var,
            state="readonly",
            width=35
        )
        self.reprint_dropdown.pack(side=tk.LEFT, padx=(0, 10), fill=tk.X, expand=True)

        self.reprint_btn = tk.Button(
            reprint_frame,
            text="Reprint",
            command=self._reprint_label,
            bg="#F39C12",
            fg="white",
            padx=10,
            state=tk.DISABLED
        )
        self.reprint_btn.pack(side=tk.RIGHT)

        # Bottom button frame
        bottom_frame = tk.Frame(content_frame)
        bottom_frame.pack(fill=tk.X, pady=(10, 0))
        
        help_btn = tk.Button(
            bottom_frame,
            text="Help",
            command=self.show_help,
            bg="#3498DB",
            fg="white",
            padx=10
        )
        help_btn.pack(side=tk.RIGHT)
    
    def _detect_label_size(self, zpl_code: str) -> tuple[int, int] | None:
        """Determine label size in mm. Uses filename override for EWS templates,
        otherwise parses ^PW and ^LL from ZPL (at 300 DPI)."""
        # EWS-only templates have a known size (PRN dimensions are slightly off)
        name = os.path.basename(self.prn_file).upper()
        has_ews = "EWS" in name
        has_gw = "GW" in name or "GATEWAY" in name
        if has_ews and not has_gw:
            return (38, 19)

        pw_match = re.search(r'\^PW(\d+)', zpl_code)
        ll_match = re.search(r'\^LL0?(\d+)', zpl_code)
        if pw_match and ll_match:
            pw_dots = int(pw_match.group(1))
            ll_dots = int(ll_match.group(1))
            width_mm = round(pw_dots / 11.811)
            height_mm = round(ll_dots / 11.811)
            return (width_mm, height_mm)
        return None

    def load_template(self):
        """Load the ZPL template from PRN file"""
        if not os.path.exists(self.prn_file):
            self.template_label.config(text="No template loaded", fg="#E74C3C")
            self.status_label.config(text=f"Template file not found: {self.prn_file}", fg="red")
            return False
        
        try:
            with open(self.prn_file, 'r', encoding='utf-8-sig') as f:
                self.zpl_template = f.read()
            display_name = os.path.basename(self.prn_file)
            size = self._detect_label_size(self.zpl_template)
            if size:
                display_name += f"  ({size[0]}×{size[1]}mm)"
            self.template_label.config(text=display_name, fg="#27AE60")
            self.status_label.config(text=f"Template loaded: {self.prn_file}", fg="green")
            
            # Detect serial placeholder
            self._serial_insert_pos = None
            self._serial_placeholder = self._detect_serial_placeholder(self.zpl_template)
            if self._serial_placeholder is None:
                self._serial_placeholder = self._show_placeholder_editor(
                    self.zpl_template
                )
                if self._serial_placeholder is None:
                    self.status_label.config(
                        text="Warning: No serial placeholder defined", fg="orange"
                    )
            return True
        except Exception as e:
            self.template_label.config(text="Load error", fg="#E74C3C")
            self.status_label.config(text=f"Error loading template: {e}", fg="red")
            return False
    
    def browse_template(self):
        """Browse for a PRN template file"""
        filename = filedialog.askopenfilename(
            title="Select PRN Template File",
            filetypes=[("PRN files", "*.prn"), ("All files", "*.*")]
        )
        
        if filename:
            self.prn_file = filename
            if self.load_template() and self.zpl_template:
                size = self._detect_label_size(self.zpl_template)
                if size:
                    confirm = messagebox.askyesno(
                        "Confirm Label Size",
                        f"This template is designed for {size[0]}×{size[1]}mm labels.\n\n"
                        "Are the correct labels loaded in the printer?"
                    )
                    if not confirm:
                        self.status_label.config(
                            text="Swap labels and reload template when ready",
                            fg="orange"
                        )
    
    def refresh_printers(self):
        """Refresh the list of available printers"""
        if not self.usb_available:
            if sys.platform == "win32":
                msg = "pywin32 not installed - see README"
                hint = "Install pywin32: pip install pywin32"
            elif sys.platform.startswith("linux"):
                msg = "pycups not installed - see README"
                hint = "Install pycups: pip install pycups"
            else:
                msg = "Unsupported platform"
                hint = "Windows or Linux required for printing"
            self.printer_dropdown['values'] = [msg]
            self.status_label.config(text=hint, fg="red")
            return

        try:
            printers = []
            if _PLATFORM == "windows":
                assert win32print is not None
                for printer_info in win32print.EnumPrinters(2):
                    printers.append(printer_info[2])
            elif _PLATFORM == "linux":
                assert cups is not None
                conn = cups.Connection()
                printers = list(conn.getPrinters().keys())

            if not printers:
                self.printer_dropdown['values'] = ["No printers found"]
                self.status_label.config(text="No printers found", fg="red")
            else:
                self.printer_dropdown['values'] = printers
                # Select default printer
                try:
                    if _PLATFORM == "windows":
                        assert win32print is not None
                        default = win32print.GetDefaultPrinter()
                        self.printer_var.set(default)
                    elif _PLATFORM == "linux":
                        assert cups is not None
                        conn = cups.Connection()
                        default = conn.getDefault()
                        if default and default in printers:
                            self.printer_var.set(default)
                        else:
                            self.printer_var.set(printers[0])
                except Exception:
                    self.printer_var.set(printers[0])

                self.status_label.config(text=f"Found {len(printers)} printer(s)", fg="green")

        except Exception as e:
            self.status_label.config(text=f"Error getting printers: {e}", fg="red")
    
    def increase_quantity(self):
        """Increase quantity by 1"""
        current = self.quantity_var.get()
        if current < 999:
            self.quantity_var.set(current + 1)
    
    def decrease_quantity(self):
        """Decrease quantity by 1"""
        current = self.quantity_var.get()
        if current > 1:
            self.quantity_var.set(current - 1)
    
    def replace_serial_number(self, zpl_code, new_serial):
        """Replace the serial number placeholder in the ZPL code"""
        if self._serial_insert_pos is not None:
            # Index-based insertion (user clicked a position in the editor)
            pos = self._serial_insert_pos
            return zpl_code[:pos] + new_serial + zpl_code[pos:]
        if self._serial_placeholder == "":
            # Insertion mode: serial goes right after S/N:
            return zpl_code.replace("S/N:", f"S/N:{new_serial}")
        if self._serial_placeholder:
            return zpl_code.replace(self._serial_placeholder, new_serial)
        return zpl_code

    def _detect_serial_placeholder(self, zpl_code: str) -> str | None:
        """Detect serial number placeholder from S/N: pattern in ZPL.
        Returns the placeholder string, empty string for insertion mode,
        or None if no S/N: found at all."""
        # Try explicit placeholder (e.g. S/N:XXXX-XXXX)
        match = re.search(r'S/N:(.+?)(?:\^FS|\^)', zpl_code)
        if match:
            return match.group(1)
        # Check if S/N: exists with nothing after it (insertion mode)
        if re.search(r'S/N:\^', zpl_code):
            return ""
        return None

    def _show_placeholder_editor(self, zpl_code: str) -> str | None:
        """Show ZPL content and let user select text or click to set insertion point."""
        result: list[str | None] = [None]

        dlg = tk.Toplevel(self.root)
        dlg.title("Select Serial Number Placeholder")
        dlg.geometry("650x600")
        dlg.transient(self.root)
        dlg.grab_set()

        tk.Label(
            dlg,
            text=(
                "No S/N: placeholder detected in this template.\n\n"
                "Option 1: Highlight text to replace, then click 'Replace Selection'\n"
                "Option 2: Click where the serial should be inserted, then click 'Insert Here'"
            ),
            font=("Arial", 10),
            justify=tk.LEFT,
            padx=10, pady=10
        ).pack(fill=tk.X)

        # Scrollable text widget showing ZPL
        text_frame = tk.Frame(dlg)
        text_frame.pack(fill=tk.BOTH, expand=True, padx=10)

        scrollbar = tk.Scrollbar(text_frame)
        scrollbar.pack(side=tk.RIGHT, fill=tk.Y)

        text_widget = tk.Text(
            text_frame,
            wrap=tk.NONE,
            font=("Consolas", 10),
            yscrollcommand=scrollbar.set
        )
        text_widget.pack(fill=tk.BOTH, expand=True)
        scrollbar.config(command=text_widget.yview)

        text_widget.insert(tk.END, zpl_code)

        # Cursor position label
        pos_label = tk.Label(dlg, text="Cursor: —", font=("Arial", 9), fg="#555")
        pos_label.pack(pady=(5, 0))

        def update_pos_label(event=None):
            idx = text_widget.index(tk.INSERT)
            # Convert tkinter index "line.col" to flat character offset
            line, col = map(int, idx.split("."))
            pos_label.config(text=f"Cursor: line {line}, col {col}")

        text_widget.bind("<ButtonRelease-1>", update_pos_label)
        text_widget.bind("<KeyRelease>", update_pos_label)

        # Buttons
        btn_frame = tk.Frame(dlg, pady=10)
        btn_frame.pack(fill=tk.X)

        def on_replace_selection():
            try:
                sel = text_widget.selection_get()
                if sel.strip():
                    result[0] = sel
                    self._serial_insert_pos = None
                else:
                    raise tk.TclError("empty")
            except tk.TclError:
                messagebox.showwarning(
                    "No Selection",
                    "Please highlight the placeholder text first.",
                    parent=dlg
                )
                return
            dlg.destroy()

        def on_insert_here():
            # Convert tkinter text index to flat character offset
            idx = text_widget.index(tk.INSERT)
            line, col = map(int, idx.split("."))
            # Count characters up to that line and column
            flat_pos = 0
            for i in range(1, line):
                line_text = text_widget.get(f"{i}.0", f"{i}.end")
                flat_pos += len(line_text) + 1  # +1 for newline
            flat_pos += col
            self._serial_insert_pos = flat_pos
            result[0] = ""  # empty string signals insertion mode
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        tk.Button(
            btn_frame, text="Replace Selection",
            command=on_replace_selection, bg="#27AE60", fg="white",
            font=("Arial", 10, "bold"), padx=15
        ).pack(side=tk.LEFT, padx=(20, 10))

        tk.Button(
            btn_frame, text="Insert Here",
            command=on_insert_here, bg="#3498DB", fg="white",
            font=("Arial", 10, "bold"), padx=15
        ).pack(side=tk.LEFT, padx=(0, 10))

        tk.Button(
            btn_frame, text="Cancel",
            command=on_cancel, bg="#E74C3C", fg="white",
            font=("Arial", 10), padx=15
        ).pack(side=tk.LEFT)

        self.root.wait_window(dlg)
        return result[0]

    def _on_auto_paste_toggle(self):
        """Start or stop clipboard polling."""
        if self.auto_paste_var.get():
            self._last_clipboard = ""
            self._check_clipboard()
            self.auto_print_check.config(state=tk.NORMAL)
        else:
            if self._clipboard_poll_id:
                self.root.after_cancel(self._clipboard_poll_id)
                self._clipboard_poll_id = None
            self.auto_print_var.set(False)
            self.auto_print_check.config(state=tk.DISABLED)

    def _check_clipboard(self):
        """Poll clipboard for serial number patterns."""
        if not self.auto_paste_var.get():
            return
        try:
            clip = self.root.clipboard_get().strip()
        except tk.TclError:
            clip = ""

        if clip and clip != self._last_clipboard and clip.isalnum():
            self._last_clipboard = clip
            serial = None
            if len(clip) == 8:
                serial = f"{clip[:4]}-{clip[4:]}"
            elif len(clip) == 6:
                serial = clip

            if serial and not self.serial_entry.get().strip():
                self.serial_entry.delete(0, tk.END)
                self.serial_entry.insert(0, serial)
                self.status_label.config(text=f"Auto-pasted: {serial}", fg="blue")

                if (self.auto_print_var.get()
                        and serial != self._last_printed_serial):
                    self._auto_print_one(serial)

        self._clipboard_poll_id = self.root.after(500, self._check_clipboard)

    def _auto_print_one(self, serial: str):
        """Auto-print a single label for the given serial."""
        text_only = self.text_only_var.get()
        if not text_only and not self.zpl_template:
            self.status_label.config(
                text="Auto-print skipped: no template loaded", fg="orange"
            )
            return

        printer = self.printer_var.get()
        if not printer or printer in _INVALID_PRINTERS:
            self.status_label.config(
                text="Auto-print skipped: no printer selected", fg="orange"
            )
            return

        try:
            if text_only:
                zpl_code = self._generate_text_only_zpl(serial)
            else:
                zpl_code = self.replace_serial_number(self.zpl_template, serial)

            self.send_to_usb_printer(zpl_code, printer)
            self._last_printed_serial = serial
            self.status_label.config(
                text=f"Auto-printed 1 label: {serial}", fg="green"
            )

            # Store in recent labels
            mode = "Text" if text_only else "Template"
            display = f"{mode}: {serial[:30]}"
            self._recent_labels.insert(0, {"display": display, "zpl": zpl_code})
            self._recent_labels = self._recent_labels[:5]
            self.reprint_dropdown['values'] = [
                e["display"] for e in self._recent_labels
            ]
            self.reprint_dropdown.current(0)
            self.reprint_btn.config(state=tk.NORMAL)

            # Clear entry for next serial
            self.serial_entry.delete(0, tk.END)
        except Exception as e:
            self.status_label.config(
                text=f"Auto-print error: {e}", fg="red"
            )
    
    def send_to_usb_printer(self, zpl_code, printer_name):
        """Send ZPL code to printer (Windows via win32print, Linux via CUPS)"""
        if not self.usb_available:
            raise Exception(
                "No printing backend available. "
                "Windows: pip install pywin32 | Linux: pip install pycups"
            )

        if _PLATFORM == "windows":
            assert win32print is not None
            try:
                hPrinter = win32print.OpenPrinter(printer_name)
                try:
                    win32print.StartDocPrinter(hPrinter, 1, ("Label", "", "RAW"))
                    win32print.StartPagePrinter(hPrinter)
                    win32print.WritePrinter(hPrinter, zpl_code.encode('utf-8'))
                    win32print.EndPagePrinter(hPrinter)
                    win32print.EndDocPrinter(hPrinter)
                    return True
                finally:
                    win32print.ClosePrinter(hPrinter)
            except Exception as e:
                raise Exception(f"Print error: {str(e)}")

        elif _PLATFORM == "linux":
            assert cups is not None
            try:
                conn = cups.Connection()
                # Write ZPL to temp file — CUPS printFile needs a path
                with tempfile.NamedTemporaryFile(
                    mode='w', suffix='.zpl', delete=False
                ) as tmp:
                    tmp.write(zpl_code)
                    tmp_path = tmp.name
                try:
                    conn.printFile(
                        printer_name, tmp_path, "Label", {"raw": ""}
                    )
                    return True
                finally:
                    os.unlink(tmp_path)
            except Exception as e:
                raise Exception(f"CUPS print error: {str(e)}")
    
    def print_label(self):
        """Main print function"""
        text_only = self.text_only_var.get()
        
        # Validate template is loaded (only in template mode)
        if not text_only and not self.zpl_template:
            messagebox.showerror("Error", "No template loaded! Please load a PRN file.")
            return
        
        # Validate text input
        serial = self.serial_entry.get().strip()
        if not serial:
            label = "label text" if text_only else "serial number"
            messagebox.showerror("Error", f"Please enter a {label}!")
            self.serial_entry.focus()
            return
        
        # Validate printer selected
        printer = self.printer_var.get()
        if not printer or printer in _INVALID_PRINTERS:
            messagebox.showerror("Error", "Please select a valid printer!")
            return
        
        # Get quantity
        quantity = self.quantity_var.get()
        
        # Disable print button during printing
        self.print_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Printing...", fg="blue")
        self.root.update()
        
        try:
            # Generate ZPL
            if text_only:
                zpl_code = self._generate_text_only_zpl(serial)
            else:
                zpl_code = self.replace_serial_number(self.zpl_template, serial)
            
            # Print labels
            for i in range(quantity):
                self.status_label.config(text=f"Printing {i+1}/{quantity}...", fg="blue")
                self.root.update()
                
                self.send_to_usb_printer(zpl_code, printer)
            
            # Success
            self.status_label.config(
                text=f"✓ Printed {quantity} label(s) with serial: {serial}",
                fg="green"
            )

            # Store in recent labels
            mode = "Text" if text_only else "Template"
            display = f"{mode}: {serial[:30]}"
            self._recent_labels.insert(0, {"display": display, "zpl": zpl_code})
            self._recent_labels = self._recent_labels[:5]
            self.reprint_dropdown['values'] = [e["display"] for e in self._recent_labels]
            self.reprint_dropdown.current(0)
            self.reprint_btn.config(state=tk.NORMAL)
            
            # Clear serial number for next print
            self.serial_entry.delete(0, tk.END)
            self.serial_entry.focus()
        
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg="red")
            messagebox.showerror("Print Error", str(e))
        
        finally:
            # Re-enable print button
            self.print_btn.config(state=tk.NORMAL)
    
    def show_help(self):
        """Show help dialog"""
        help_text = """MEMSYS Zebra Label Printer — Help

── LABEL TYPES ──
• EWS labels: 38×19mm — use "Memsys EWS Zebra label.prn"
• Gateway labels: 57×32mm — use "Memsys GW Zebra label.prn"
• Text-only: any text, auto-sized (no template needed)

── HOW TO USE ──
Template mode:
  1. Click Browse and select a .prn template
  2. Confirm the correct label size is loaded
  3. Serial placeholder is auto-detected from S/N: pattern
  4. Enter the serial number
  5. Set quantity and click PRINT LABELS (or Enter)

Text-only mode:
  1. Check "Text-only label" checkbox
  2. Type your text — each word prints on its own line
  3. Set quantity and click PRINT LABELS (or Enter)

── AUTO-PASTE & AUTO-PRINT ──
  1. Enable "Auto-paste from clipboard"
  2. Copy a serial in any app:
     • 8 chars (ABCD1234) → formatted as ABCD-1234 (EWS)
     • 6 chars (ABC123) → pasted as-is (Gateway)
  3. Enable "Auto-print on paste" to print 1 label
     per new serial automatically

Reprint: Select a recent label from the dropdown
and click Reprint for a single copy.

── SWAPPING LABEL ROLLS ──
  1. Open the printer cover (pull latch forward)
  2. Open the yellow spring-loaded media guides
  3. Remove the current roll, insert the new one
  4. Release guides — they self-center the roll
  5. Feed labels under the printhead
  6. Verify the yellow feed sensor is centered
     on the labels (under the printhead)
  7. Close the cover
  8. Calibrate: hold PAUSE + CANCEL for 2 sec
     — printer feeds labels and auto-detects size
  9. Print a test label to confirm alignment

── TROUBLESHOOTING ──
• No printers: Install pywin32 (Win) or pycups (Linux)
• Template not found: Click Browse to select one
• Print fails: Check printer is on and connected
• Wrong size: Swap labels and recalibrate (see above)
• No placeholder found: A viewer opens to select it

For more details, see README_GUI.md
"""
        messagebox.showinfo("Help", help_text)
    
    def _on_mode_change(self):
        """Toggle between template mode and text-only mode"""
        if self.text_only_var.get():
            self.template_frame.pack_forget()
            self.serial_frame.config(text="Label Text")
        else:
            # Re-insert template frame before the serial frame
            self.template_frame.pack(before=self.serial_frame, fill=tk.X, pady=(0, 15))
            self.serial_frame.config(text="Serial Number")
    
    def _reprint_label(self):
        """Reprint the selected recent label (1 copy)"""
        idx = self.reprint_dropdown.current()
        if idx < 0 or idx >= len(self._recent_labels):
            self.status_label.config(text="No label selected to reprint", fg="red")
            return

        printer = self.printer_var.get()
        if not printer or printer in _INVALID_PRINTERS:
            messagebox.showerror("Error", "Please select a valid printer!")
            return

        entry = self._recent_labels[idx]
        self.reprint_btn.config(state=tk.DISABLED)
        self.status_label.config(text="Reprinting...", fg="blue")
        self.root.update()

        try:
            self.send_to_usb_printer(entry["zpl"], printer)
            self.status_label.config(
                text=f"✓ Reprinted 1 label: {entry['display']}",
                fg="green"
            )
        except Exception as e:
            self.status_label.config(text=f"Error: {str(e)}", fg="red")
            messagebox.showerror("Print Error", str(e))
        finally:
            self.reprint_btn.config(state=tk.NORMAL)

    def _generate_text_only_zpl(self, text):
        """Generate ZPL for a plain text label (38x19mm at 300 dpi)"""
        # 38x19mm at 300 dpi = 449 wide x 224 tall
        pw, ll = 449, 224

        words = text.split()
        longest_word = max(len(w) for w in words) if words else 0

        # Font size based on longest single word
        if longest_word <= 10:
            font_h = font_w = 70
        elif longest_word <= 12:
            font_h = font_w = 50
        else:
            font_h = font_w = 36

        # One word per line, capped at 3
        max_lines = min(len(words), 3) if words else 1

        # Explicit line breaks so each word gets its own centered line
        # Use separate field commands per word for maximum compatibility
        block_height = max_lines * font_h
        y_offset = max((ll - block_height) // 2, 0)

        field_commands = ""
        for i, word in enumerate(words[:max_lines]):
            y = y_offset + i * font_h
            field_commands += (
                f"^FO0,{y}^A0N,{font_h},{font_w}"
                f"^FB{pw},1,0,C^FD{word}^FS\n"
            )

        return (
            "^XA\n"
            "^MMT\n"
            f"^PW{pw}\n"
            f"^LL0{ll}\n"
            "^LS0\n"
            f"{field_commands}"
            "^PQ1,0,1,Y^XZ"
        )


def main():
    """Main entry point"""
    root = tk.Tk()
    app = ZebraLabelPrinter(root)
    root.mainloop()


if __name__ == "__main__":
    main()
