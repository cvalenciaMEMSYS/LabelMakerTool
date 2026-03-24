#!/usr/bin/env python3
"""
Zebra Label Printer - GUI Version with USB Support
Simple graphical interface for printing labels with custom serial numbers
"""

import tkinter as tk
from tkinter import ttk, messagebox, filedialog
import os
import sys
from pathlib import Path

try:
    import win32print  # type: ignore[import-untyped]
    _USB_AVAILABLE = True
except ImportError:
    win32print = None  # type: ignore[assignment]
    _USB_AVAILABLE = False


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
        self.usb_available = _USB_AVAILABLE
        self._recent_labels: list[dict] = []
        
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
            self.template_label.config(text=display_name, fg="#27AE60")
            self.status_label.config(text=f"Template loaded: {self.prn_file}", fg="green")
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
            self.load_template()
    
    def refresh_printers(self):
        """Refresh the list of available USB printers"""
        if not self.usb_available:
            self.printer_dropdown['values'] = ["pywin32 not installed - see README"]
            self.status_label.config(
                text="Install pywin32: pip install pywin32",
                fg="red"
            )
            return
        
        assert win32print is not None
        try:
            # Get list of all printers
            printers = []
            for printer_info in win32print.EnumPrinters(2):
                printers.append(printer_info[2])
            
            if not printers:
                self.printer_dropdown['values'] = ["No printers found"]
                self.status_label.config(text="No printers found", fg="red")
            else:
                self.printer_dropdown['values'] = printers
                # Select default printer
                try:
                    default = win32print.GetDefaultPrinter()
                    self.printer_var.set(default)
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
        # Replace XXXX-XXXX placeholder with the new serial number
        modified_zpl = zpl_code.replace('XXXX-XXXX', new_serial)
        return modified_zpl
    
    def send_to_usb_printer(self, zpl_code, printer_name):
        """Send ZPL code to USB printer"""
        if not self.usb_available:
            raise Exception("pywin32 not installed. Install with: pip install pywin32")
        
        assert win32print is not None
        try:
            # Open printer
            hPrinter = win32print.OpenPrinter(printer_name)
            
            try:
                # Start a print job
                hJob = win32print.StartDocPrinter(hPrinter, 1, ("Label", "", "RAW"))
                win32print.StartPagePrinter(hPrinter)
                
                # Send ZPL
                win32print.WritePrinter(hPrinter, zpl_code.encode('utf-8'))
                
                # End the job
                win32print.EndPagePrinter(hPrinter)
                win32print.EndDocPrinter(hPrinter)
                
                return True
                
            finally:
                win32print.ClosePrinter(hPrinter)
        
        except Exception as e:
            raise Exception(f"Print error: {str(e)}")
    
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
        if not printer or printer in ["No printers found", "pywin32 not installed - see README"]:
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
        help_text = """MEMSYS Zebra Label Printer - Help

How to Use:
1. Select your Zebra printer from the dropdown
2. Enter the serial number
3. Set the quantity (default: 1)
4. Click "PRINT LABELS" or press Enter

Keyboard Shortcuts:
- Enter: Print label
- +/-: Increase/decrease quantity

Requirements:
- Python 3.6+
- pywin32 (install with: pip install pywin32)
- Zebra printer installed in Windows
- PRN template file

Troubleshooting:
- No printers listed: Install pywin32
- Template not found: Click "Load Template File"
- Print fails: Check printer is on and connected

For more help, see README.md
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
        if not printer or printer in ["No printers found", "pywin32 not installed - see README"]:
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
