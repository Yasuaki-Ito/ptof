"""
PtoF - PPTX to Figures - GUI Interface

Module providing GUI interface using CustomTkinter.
"""

import os
import sys
import threading
from pathlib import Path
from tkinter import filedialog, messagebox, colorchooser

import customtkinter as ctk
from tkinterdnd2 import DND_FILES, TkinterDnD

from .core import process_pptx, parse_color, COLOR_NAMES


def get_resource_path(relative_path):
    """Get path for resources bundled with PyInstaller"""
    if hasattr(sys, '_MEIPASS'):
        # Built with PyInstaller
        return Path(sys._MEIPASS) / relative_path
    # Development environment
    return Path(__file__).parent.parent / relative_path


class CTkDnD(ctk.CTk, TkinterDnD.DnDWrapper):
    """CustomTkinter with Drag and Drop support"""
    def __init__(self, *args, **kwargs):
        super().__init__(*args, **kwargs)
        self.TkdndVersion = TkinterDnD._require(self)


class App(CTkDnD):
    def __init__(self):
        super().__init__()

        # Window settings
        self.title("PtoF - PPTX to Figures")
        self.geometry("700x620")
        self.minsize(600, 550)

        # Icon settings
        icon_path = get_resource_path("icon.ico")
        if icon_path.exists():
            self.iconbitmap(str(icon_path))

        # Theme settings
        ctk.set_appearance_mode("dark")
        ctk.set_default_color_theme("blue")

        # Variables
        self.input_files = []
        self.processing = False

        # Build UI
        self._create_widgets()

    def _create_widgets(self):
        # Main frame
        self.main_frame = ctk.CTkFrame(self)
        self.main_frame.pack(fill="both", expand=True, padx=20, pady=20)

        # Title
        self.title_label = ctk.CTkLabel(
            self.main_frame,
            text="PPTX to PDF/PNG/SVG",
            font=ctk.CTkFont(size=24, weight="bold")
        )
        self.title_label.pack(pady=(0, 20))

        # Input files section
        self._create_input_section()

        # Output directory section
        self._create_output_section()

        # Options section
        self._create_options_section()

        # Action buttons
        self._create_action_section()

        # Log section
        self._create_log_section()

    def _create_input_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=(0, 10))

        label = ctk.CTkLabel(frame, text="Input Files:", font=ctk.CTkFont(weight="bold"))
        label.pack(anchor="w", padx=10, pady=(10, 5))

        # Drop zone
        self.drop_zone = ctk.CTkFrame(
            frame,
            height=80,
            border_width=2,
            border_color="gray",
            fg_color=("gray90", "gray20")
        )
        self.drop_zone.pack(fill="x", padx=10, pady=(0, 10))
        self.drop_zone.pack_propagate(False)

        self.drop_label = ctk.CTkLabel(
            self.drop_zone,
            text="Drop PPTX files here\nor click Browse",
            text_color="gray",
            font=ctk.CTkFont(size=13)
        )
        self.drop_label.pack(expand=True)

        # Register drop events
        self.drop_zone.drop_target_register(DND_FILES)
        self.drop_zone.dnd_bind('<<Drop>>', self._on_drop)
        self.drop_zone.dnd_bind('<<DragEnter>>', self._on_drag_enter)
        self.drop_zone.dnd_bind('<<DragLeave>>', self._on_drag_leave)

        # Click drop zone to browse
        self.drop_zone.bind("<Button-1>", lambda e: self._browse_input())
        self.drop_label.bind("<Button-1>", lambda e: self._browse_input())

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.input_label = ctk.CTkLabel(
            btn_frame,
            text="No files selected",
            text_color="gray",
            anchor="w"
        )
        self.input_label.pack(side="left", fill="x", expand=True)

        self.browse_btn = ctk.CTkButton(
            btn_frame,
            text="Browse...",
            width=100,
            command=self._browse_input
        )
        self.browse_btn.pack(side="right")

        self.clear_btn = ctk.CTkButton(
            btn_frame,
            text="Clear",
            width=60,
            fg_color="gray",
            hover_color="darkgray",
            command=self._clear_input
        )
        self.clear_btn.pack(side="right", padx=(0, 10))

    def _create_output_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=(0, 10))

        label = ctk.CTkLabel(frame, text="Output Directory:", font=ctk.CTkFont(weight="bold"))
        label.pack(anchor="w", padx=10, pady=(10, 5))

        btn_frame = ctk.CTkFrame(frame, fg_color="transparent")
        btn_frame.pack(fill="x", padx=10, pady=(0, 10))

        self.output_var = ctk.StringVar(value=os.getcwd())
        self.output_entry = ctk.CTkEntry(btn_frame, textvariable=self.output_var)
        self.output_entry.pack(side="left", fill="x", expand=True)

        self.output_btn = ctk.CTkButton(
            btn_frame,
            text="Browse...",
            width=100,
            command=self._browse_output
        )
        self.output_btn.pack(side="right", padx=(10, 0))

    def _create_options_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="x", pady=(0, 10))

        # Header (click to expand/collapse)
        header_frame = ctk.CTkFrame(frame, fg_color="transparent")
        header_frame.pack(fill="x", padx=10, pady=(10, 5))

        self.options_expanded = False
        self.options_toggle_btn = ctk.CTkButton(
            header_frame,
            text="▶ Options",
            font=ctk.CTkFont(weight="bold"),
            fg_color="transparent",
            hover_color=("gray80", "gray30"),
            text_color=("black", "white"),
            anchor="w",
            width=120,
            command=self._toggle_options
        )
        self.options_toggle_btn.pack(side="left")

        # Options grid (initially hidden)
        self.options_grid = ctk.CTkFrame(frame, fg_color="transparent")
        # Not packed initially

        # Color selection
        color_frame = ctk.CTkFrame(self.options_grid, fg_color="transparent")
        color_frame.pack(fill="x", pady=2)

        ctk.CTkLabel(color_frame, text="Marker Color:", width=120, anchor="w").pack(side="left")

        # Preset selection
        self.color_var = ctk.StringVar(value="cyan")
        self.color_menu = ctk.CTkOptionMenu(
            color_frame,
            values=list(COLOR_NAMES.keys()),
            variable=self.color_var,
            command=self._on_color_preset_change,
            width=100
        )
        self.color_menu.pack(side="left", padx=(10, 0))

        # HEX input field
        self.color_entry_var = ctk.StringVar(value="#00FFFF")
        self.color_entry = ctk.CTkEntry(
            color_frame,
            textvariable=self.color_entry_var,
            width=80,
            placeholder_text="#RRGGBB"
        )
        self.color_entry.pack(side="left", padx=(10, 0))
        self.color_entry.bind("<Return>", self._on_color_entry_change)
        self.color_entry.bind("<FocusOut>", self._on_color_entry_change)

        # Color preview
        self.color_preview = ctk.CTkLabel(
            color_frame,
            text="",
            width=30,
            height=30,
            fg_color="#00FFFF",
            corner_radius=5
        )
        self.color_preview.pack(side="left", padx=(10, 0))

        # Color picker button
        self.color_picker_btn = ctk.CTkButton(
            color_frame,
            text="...",
            width=30,
            command=self._pick_color
        )
        self.color_picker_btn.pack(side="left", padx=(5, 0))

        # DPI
        dpi_frame = ctk.CTkFrame(self.options_grid, fg_color="transparent")
        dpi_frame.pack(fill="x", pady=2)

        ctk.CTkLabel(dpi_frame, text="DPI (for PNG):", width=120, anchor="w").pack(side="left")
        self.dpi_var = ctk.StringVar(value="300")
        self.dpi_entry = ctk.CTkEntry(dpi_frame, textvariable=self.dpi_var, width=150)
        self.dpi_entry.pack(side="left", padx=(10, 0))

        # Margin
        margin_frame = ctk.CTkFrame(self.options_grid, fg_color="transparent")
        margin_frame.pack(fill="x", pady=2)

        ctk.CTkLabel(margin_frame, text="Margin (pt):", width=120, anchor="w").pack(side="left")
        self.margin_var = ctk.StringVar(value="0")
        self.margin_entry = ctk.CTkEntry(margin_frame, textvariable=self.margin_var, width=150)
        self.margin_entry.pack(side="left", padx=(10, 0))

        # Checkboxes
        checkbox_frame = ctk.CTkFrame(self.options_grid, fg_color="transparent")
        checkbox_frame.pack(fill="x", pady=(10, 0))

        self.embed_fonts_var = ctk.BooleanVar(value=False)
        self.embed_fonts_cb = ctk.CTkCheckBox(
            checkbox_frame,
            text="Embed Fonts (PDF/A)",
            variable=self.embed_fonts_var
        )
        self.embed_fonts_cb.pack(side="left", padx=(0, 20))

        self.include_background_var = ctk.BooleanVar(value=False)
        self.include_background_cb = ctk.CTkCheckBox(
            checkbox_frame,
            text="Include Slide Background",
            variable=self.include_background_var
        )
        self.include_background_cb.pack(side="left", padx=(0, 20))

    def _create_action_section(self):
        frame = ctk.CTkFrame(self.main_frame, fg_color="transparent")
        frame.pack(fill="x", pady=(0, 10))

        self.dry_run_btn = ctk.CTkButton(
            frame,
            text="Dry Run",
            width=120,
            fg_color="gray",
            hover_color="darkgray",
            command=self._dry_run
        )
        self.dry_run_btn.pack(side="left")

        self.convert_btn = ctk.CTkButton(
            frame,
            text="Convert",
            width=120,
            command=self._convert
        )
        self.convert_btn.pack(side="right")

        # Progress bar
        self.progress = ctk.CTkProgressBar(frame)
        self.progress.pack(side="left", fill="x", expand=True, padx=20)
        self.progress.set(0)

    def _create_log_section(self):
        frame = ctk.CTkFrame(self.main_frame)
        frame.pack(fill="both", expand=True)

        label = ctk.CTkLabel(frame, text="Log:", font=ctk.CTkFont(weight="bold"))
        label.pack(anchor="w", padx=10, pady=(10, 5))

        self.log_text = ctk.CTkTextbox(frame, height=150)
        self.log_text.pack(fill="both", expand=True, padx=10, pady=(0, 10))

    def _toggle_options(self):
        """Expand/collapse options section"""
        if self.options_expanded:
            self.options_grid.pack_forget()
            self.options_toggle_btn.configure(text="▶ Options")
            self.options_expanded = False
        else:
            self.options_grid.pack(fill="x", padx=10, pady=(0, 10))
            self.options_toggle_btn.configure(text="▼ Options")
            self.options_expanded = True

    def _browse_input(self):
        files = filedialog.askopenfilenames(
            title="Select PPTX files",
            filetypes=[("PowerPoint files", "*.pptx"), ("All files", "*.*")]
        )
        if files:
            self._set_input_files(list(files))

    def _set_input_files(self, files):
        """Set input files"""
        # Filter .pptx files only
        pptx_files = [f for f in files if f.lower().endswith('.pptx')]
        if pptx_files:
            self.input_files = pptx_files
            if len(self.input_files) == 1:
                self.input_label.configure(text=Path(self.input_files[0]).name)
            else:
                self.input_label.configure(text=f"{len(self.input_files)} files selected")
            self.input_label.configure(text_color="#00D4AA")  # 明るいシアングリーン
            self.drop_label.configure(
                text="\n".join(Path(f).name for f in self.input_files[:3]) +
                     ("\n..." if len(self.input_files) > 3 else ""),
                text_color="#00D4AA"
            )

            # Set output directory to input file's directory + output_dir
            self.output_var.set(str(Path(self.input_files[0]).parent / "output_dir"))

    def _clear_input(self):
        """Clear input files"""
        self.input_files = []
        self.input_label.configure(text="No files selected", text_color="gray")
        self.drop_label.configure(text="Drop PPTX files here\nor click Browse", text_color="gray")

    def _on_drop(self, event):
        """When files are dropped"""
        # Parse dropped file paths
        files = self._parse_drop_data(event.data)
        if files:
            self._set_input_files(files)
        self._on_drag_leave(event)

    def _on_drag_enter(self, event):
        """When drag enters the zone"""
        self.drop_zone.configure(border_color=("#3B8ED0", "#1F6AA5"), fg_color=("gray85", "gray25"))

    def _on_drag_leave(self, event):
        """When drag leaves the zone"""
        self.drop_zone.configure(border_color="gray", fg_color=("gray90", "gray20"))

    def _parse_drop_data(self, data):
        """Convert drop data to list of file paths"""
        files = []
        # On Windows, paths may be wrapped in {} or space-separated
        if '{' in data:
            # {path1} {path2} format
            import re
            files = re.findall(r'\{([^}]+)\}', data)
        else:
            # Space-separated (for paths without spaces)
            files = data.split()

        # Normalize paths
        result = []
        for f in files:
            f = f.strip()
            if f and os.path.isfile(f):
                result.append(f)
        return result

    def _browse_output(self):
        directory = filedialog.askdirectory(title="Select output directory")
        if directory:
            self.output_var.set(directory)

    def _on_color_preset_change(self, color_name):
        """Update preview and HEX input when preset color is selected"""
        if color_name in COLOR_NAMES:
            r, g, b = COLOR_NAMES[color_name]
            hex_color = f"#{r:02X}{g:02X}{b:02X}"
            self.color_preview.configure(fg_color=hex_color)
            self.color_entry_var.set(hex_color)
            self.color_var.set(color_name)

    def _on_color_entry_change(self, event=None):
        """Update preview when HEX input field changes"""
        hex_value = self.color_entry_var.get().strip()
        try:
            rgb = parse_color(hex_value)
            hex_color = f"#{rgb[0]:02X}{rgb[1]:02X}{rgb[2]:02X}"
            self.color_preview.configure(fg_color=hex_color)
            self.color_var.set(hex_value)  # Update internal variable
        except ValueError:
            pass  # Ignore invalid colors

    def _pick_color(self):
        """Open color picker"""
        # Get current color
        try:
            current_rgb = parse_color(self.color_entry_var.get())
            initial_color = f"#{current_rgb[0]:02X}{current_rgb[1]:02X}{current_rgb[2]:02X}"
        except ValueError:
            initial_color = "#00FFFF"

        color = colorchooser.askcolor(color=initial_color, title="Select Marker Color")
        if color[1]:  # color is a tuple ((R, G, B), "#RRGGBB")
            hex_color = color[1].upper()
            self.color_var.set(hex_color)
            self.color_entry_var.set(hex_color)
            self.color_preview.configure(fg_color=hex_color)

    def _log(self, message):
        self.log_text.insert("end", message + "\n")
        self.log_text.see("end")

    def _clear_log(self):
        self.log_text.delete("1.0", "end")

    def _set_ui_state(self, enabled):
        state = "normal" if enabled else "disabled"
        self.browse_btn.configure(state=state)
        self.clear_btn.configure(state=state)
        self.output_btn.configure(state=state)
        self.output_entry.configure(state=state)
        self.color_menu.configure(state=state)
        self.color_entry.configure(state=state)
        self.color_picker_btn.configure(state=state)
        self.dpi_entry.configure(state=state)
        self.margin_entry.configure(state=state)
        self.embed_fonts_cb.configure(state=state)
        self.include_background_cb.configure(state=state)
        self.dry_run_btn.configure(state=state)
        self.convert_btn.configure(state=state)

    def _validate_inputs(self):
        if not self.input_files:
            messagebox.showerror("Error", "Please select input files")
            return False

        if not self.output_var.get():
            messagebox.showerror("Error", "Please select output directory")
            return False

        try:
            parse_color(self.color_entry_var.get())
        except ValueError:
            messagebox.showerror("Error", "Invalid color. Use color name or HEX (#RRGGBB)")
            return False

        try:
            int(self.dpi_var.get())
        except ValueError:
            messagebox.showerror("Error", "DPI must be a number")
            return False

        try:
            float(self.margin_var.get())
        except ValueError:
            messagebox.showerror("Error", "Margin must be a number")
            return False

        return True

    def _dry_run(self):
        self._run_conversion(dry_run=True)

    def _convert(self):
        self._run_conversion(dry_run=False)

    def _run_conversion(self, dry_run=False):
        if not self._validate_inputs():
            return

        self._clear_log()
        self._set_ui_state(False)
        self.progress.set(0)
        self.processing = True

        # Run in background thread
        thread = threading.Thread(target=self._process_files, args=(dry_run,))
        thread.daemon = True
        thread.start()

    def _process_files(self, dry_run):
        try:
            marker_color = parse_color(self.color_entry_var.get())
            dpi = int(self.dpi_var.get())
            margin = float(self.margin_var.get())
            output_dir = self.output_var.get()
            embed_fonts = self.embed_fonts_var.get()
            include_background = self.include_background_var.get()

            total_files = len(self.input_files)
            all_output_files = []

            for file_idx, pptx_file in enumerate(self.input_files):
                self.after(0, lambda m=f"Processing: {Path(pptx_file).name}": self._log(m))

                def progress_callback(msg, current, total):
                    if total > 0:
                        file_progress = file_idx / total_files
                        item_progress = current / total / total_files
                        overall = file_progress + item_progress
                        self.after(0, lambda p=overall: self.progress.set(p))
                    self.after(0, lambda m=f"  {msg}": self._log(m))
                    return self.processing

                try:
                    output_files = process_pptx(
                        pptx_file,
                        output_dir,
                        embed_fonts,
                        marker_color,
                        dpi,
                        margin,
                        dry_run,
                        quiet=True,
                        no_overwrite=False,
                        progress_callback=progress_callback,
                        include_background=include_background
                    )
                    all_output_files.extend(output_files)

                except Exception as e:
                    self.after(0, lambda m=f"Error: {e}": self._log(m))

            # Completion
            self.after(0, lambda: self.progress.set(1))
            if dry_run:
                self.after(0, lambda: self._log(f"\n[Dry-run] Would create {len(all_output_files)} file(s)"))
            else:
                self.after(0, lambda: self._log(f"\nSuccessfully created {len(all_output_files)} file(s)"))

            for f in all_output_files:
                self.after(0, lambda m=f"  - {f}": self._log(m))

        except Exception as e:
            self.after(0, lambda: self._log(f"Error: {e}"))

        finally:
            self.processing = False
            self.after(0, lambda: self._set_ui_state(True))


def main():
    app = App()
    app.mainloop()


if __name__ == '__main__':
    main()
