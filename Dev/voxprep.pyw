"""
VoxPrep - PowerPoint Deck Utilities
Windows-only CustomTkinter UI

Features:
    - Split deck by sections (chunking)
    - Export/Import speaker notes
    - Find/Replace in notes
    - Media extraction and import
"""

import os
import sys
import json
import re
import threading
import time
import tempfile
import traceback
import tkinter as tk
from tkinter import filedialog, messagebox, ttk
import customtkinter as ctk
import logging
from logging.handlers import RotatingFileHandler
from datetime import datetime

# --- Feature modules (same pattern as Voxsmith) ---
import voxsplit
import voxnotes
import voxreplace
import voxmedia
import voxmisc

# --- App Configuration ---
APP_NAME = "VoxPrep"
APP_VERSION = "v1.0"

# --- Logging Setup (matches Voxsmith pattern) ---
LOG_PREFIX = "[voxprep]"

def log(msg: str):
    """Simple logging helper."""
    print(f"{LOG_PREFIX} {msg}", flush=True)


def _get_app_paths():
    """Get standard app directories."""
    base = os.getenv("LOCALAPPDATA") or os.getenv("APPDATA") or tempfile.gettempdir()
    root = os.path.join(base, APP_NAME)
    logs = os.path.join(root, "logs")
    tempd = os.path.join(root, "temp")
    for d in (root, logs, tempd):
        try:
            os.makedirs(d, exist_ok=True)
        except Exception:
            pass
    return {"root": root, "logs": logs, "temp": tempd}


def _configure_file_logger(logger_name="voxprep"):
    """Set up rotating file logger."""
    logger = logging.getLogger(logger_name)
    logger.setLevel(logging.INFO)
    if any(isinstance(h, RotatingFileHandler) for h in logger.handlers):
        return logger
    try:
        paths = _get_app_paths()
        log_path = os.path.join(paths["logs"], "voxprep.log")
        fh = RotatingFileHandler(log_path, maxBytes=512*1024, backupCount=3, encoding="utf-8")
        fmt = logging.Formatter("[%(asctime)s] %(levelname)s %(name)s: %(message)s")
        fh.setFormatter(fmt)
        logger.addHandler(fh)
    except Exception:
        pass
    return logger


# --- Settings Persistence ---
def _settings_path():
    return os.path.join(_get_app_paths()["root"], "settings.json")


def load_settings() -> dict:
    """Load saved settings."""
    try:
        with open(_settings_path(), "r", encoding="utf-8") as f:
            return json.load(f)
    except Exception:
        return {}


def save_settings(**kwargs):
    """Save settings (merges with existing)."""
    try:
        current = load_settings()
        current.update(kwargs)
        with open(_settings_path(), "w", encoding="utf-8") as f:
            json.dump(current, f, indent=2)
    except Exception as e:
        log(f"Failed to save settings: {e}")


# --- Single Instance Check (matches Voxsmith) ---
def _check_single_instance():
    """Prevent multiple instances."""
    import ctypes
    mutex_name = "VoxPrep_SingleInstance_Mutex"
    try:
        mutex = ctypes.windll.kernel32.CreateMutexW(None, False, mutex_name)
        last_error = ctypes.windll.kernel32.GetLastError()
        if last_error == 183:  # ERROR_ALREADY_EXISTS
            return True, mutex
        return False, mutex
    except Exception:
        return False, None


def _release_single_instance(mutex):
    """Release mutex on exit."""
    try:
        if mutex:
            import ctypes
            ctypes.windll.kernel32.ReleaseMutex(mutex)
            ctypes.windll.kernel32.CloseHandle(mutex)
    except Exception:
        pass


# --- Tooltip Helper (matches Voxsmith) ---
class ToolTip:
    """Simple tooltip for widgets."""
    def __init__(self, widget, text):
        self.widget = widget
        self.text = text
        self.tipwindow = None
        widget.bind("<Enter>", self.show)
        widget.bind("<Leave>", self.hide)

    def show(self, event=None):
        if self.tipwindow:
            return
        x = self.widget.winfo_rootx() + 20
        y = self.widget.winfo_rooty() + self.widget.winfo_height() + 5
        self.tipwindow = tw = tk.Toplevel(self.widget)
        tw.wm_overrideredirect(True)
        tw.wm_geometry(f"+{x}+{y}")
        label = tk.Label(tw, text=self.text, justify="left",
                         background="#ffffe0", relief="solid", borderwidth=1,
                         font=("Open Sans", 10))
        label.pack()

    def hide(self, event=None):
        if self.tipwindow:
            self.tipwindow.destroy()
            self.tipwindow = None


# --- Module Wrappers (add logging, error handling) ---
def do_split_deck(deck_path: str, output_dir: str, sections: list = None, log_callback=None):
    """Wrapper for voxsplit with logging."""
    logger = logging.getLogger("voxprep")
    try:
        # Get sections if not provided
        if sections is None:
            sections = voxsplit.get_powerpoint_sections(deck_path)
        
        files = voxsplit.split_deck_by_sections(deck_path, sections, output_dir, log_callback)
        logger.info(f"SPLIT deck={os.path.basename(deck_path)} sections={len(sections)} files={len(files)}")
        return {"success": True, "files": files, "sections": len(sections)}
    except Exception as e:
        logger.error(f"SPLIT failed: {e}")
        traceback.print_exc()
        if log_callback:
            log_callback(f"Error: {e}")
        return {"success": False, "error": str(e)}


def do_export_notes(deck_path: str, output_path: str, format: str, log_callback=None):
    """Wrapper for voxnotes.export_notes with logging."""
    logger = logging.getLogger("voxprep")
    try:
        result = voxnotes.export_notes(deck_path, output_path, format, log_callback=log_callback)
        logger.info(f"EXPORT deck={os.path.basename(deck_path)} format={format} output={os.path.basename(result)}")
        return {"success": True, "path": result}
    except Exception as e:
        logger.error(f"EXPORT failed: {e}")
        traceback.print_exc()
        if log_callback:
            log_callback(f"Error: {e}")
        return {"success": False, "error": str(e)}


def do_import_notes(deck_path: str, notes_file: str, preview_only: bool = False, log_callback=None):
    """Wrapper for voxnotes.import_notes with logging."""
    logger = logging.getLogger("voxprep")
    try:
        result = voxnotes.import_notes(deck_path, notes_file, preview_only, log_callback=log_callback)
        logger.info(f"IMPORT deck={os.path.basename(deck_path)} notes={os.path.basename(notes_file)} preview={preview_only}")
        
        if preview_only:
            # Normalize change format for UI
            changes = result.get("changes", [])
            normalized = []
            for c in changes:
                # Count word difference
                orig_words = len(c.get("original_notes", "").split())
                new_words = len(c.get("edited_notes", "").split())
                word_diff = new_words - orig_words
                diff_str = f"+{word_diff}" if word_diff > 0 else str(word_diff)
                
                normalized.append({
                    "slide": c["slide_number"],
                    "slide_title": c.get("slide_title", ""),
                    "change_type": c.get("change_type", "modified"),
                    "word_diff": diff_str
                })
            return {"changes": normalized, "preview": True}
        else:
            # Normalize apply result for UI
            applied = result.get("applied", [])
            return {
                "success": True,
                "slides_updated": applied,
                "errors": result.get("errors", [])
            }
    except Exception as e:
        logger.error(f"IMPORT failed: {e}")
        traceback.print_exc()
        if log_callback:
            log_callback(f"Error: {e}")
        return {"success": False, "error": str(e)}


def do_find_replace(deck_path: str, find_text: str, replace_text: str,
                    case_sensitive: bool = False, preview_only: bool = False, log_callback=None):
    """Wrapper for voxreplace with logging."""
    logger = logging.getLogger("voxprep")
    try:
        if preview_only:
            results = voxreplace.preview_replace(deck_path, find_text, replace_text, case_sensitive, log_callback=log_callback)
            logger.info(f"FIND deck={os.path.basename(deck_path)} find='{find_text}' matches={len(results)}")
            # Normalize for UI
            normalized = []
            for r in results:
                normalized.append({
                    "slide": r["slide_number"],
                    "slide_title": r.get("slide_title", ""),
                    "match_count": r.get("match_count", 0)
                })
            return normalized
        else:
            result = voxreplace.replace_in_notes(deck_path, find_text, replace_text, case_sensitive, log_callback=log_callback)
            logger.info(f"REPLACE deck={os.path.basename(deck_path)} find='{find_text}' replace='{replace_text}' count={result.get('total_replacements', 0)}")
            return {
                "success": True,
                "total_replacements": result.get("total_replacements", 0),
                "slides_modified": result.get("slides_modified", []),
                "errors": result.get("errors", [])
            }
    except Exception as e:
        logger.error(f"REPLACE failed: {e}")
        traceback.print_exc()
        if log_callback:
            log_callback(f"Error: {e}")
        return {"success": False, "error": str(e)}


# --- Main Application ---
def main():
    _configure_file_logger()
    logger = logging.getLogger("voxprep")
    logger.info(f"=== {APP_NAME} {APP_VERSION} started ===")

    # --- Window Setup ---
    ctk.set_appearance_mode("light")
    ctk.set_default_color_theme("blue")

    # Consistent background color
    BG_COLOR = "#f1f1f1"

    root = ctk.CTk()
    root.title(f"{APP_NAME} {APP_VERSION}")
    root.geometry("900x720")
    root.minsize(800, 550)
    root.configure(fg_color=BG_COLOR)

    # --- Variables ---
    pptx_var = tk.StringVar()
    
    # Split tab output folder
    split_output_var = tk.StringVar()
    
    # Notes export
    export_format_var = tk.StringVar(value="docx")
    
    # Notes import
    import_file_var = tk.StringVar()
    
    # Find/Replace
    find_var = tk.StringVar()
    replace_var = tk.StringVar()
    case_sensitive_var = tk.BooleanVar(value=False)

    # Status
    status_var = tk.StringVar(value="Ready")

    # Track action buttons for disable/enable during operations
    action_buttons = []

    def set_busy(message="Working..."):
        """Disable all action buttons and show wait cursor."""
        status_var.set(message)
        root.configure(cursor="wait")
        for btn in action_buttons:
            try:
                btn.configure(state="disabled")
            except:
                pass
        root.update_idletasks()

    def set_ready(message="Ready"):
        """Re-enable all action buttons and restore cursor."""
        status_var.set(message)
        root.configure(cursor="")
        for btn in action_buttons:
            try:
                btn.configure(state="normal")
            except:
                pass
        root.update_idletasks()

    # --- Menu Bar Frame ---
    menu_frame = ctk.CTkFrame(root, height=40, corner_radius=0, fg_color=BG_COLOR)
    menu_frame.pack(fill="x", side="top")
    menu_frame.pack_propagate(False)

    def show_about_dialog():
        messagebox.showinfo(
            f"About {APP_NAME}",
            f"{APP_NAME} {APP_VERSION}\n\n"
            "PowerPoint deck utilities for L&D professionals.\n\n"
            "• Split decks by sections\n"
            "• Export/Import speaker notes\n"
            "• Find/Replace in notes\n\n"
            "© 2024 Don Burnside\n"
            "donburnside.com"
        )

    about_btn = ctk.CTkButton(
        menu_frame, text="About", width=70, height=30,
        fg_color="transparent", text_color=("gray10", "gray90"),
        hover_color=("gray80", "gray30"), font=("Open Sans", 13),
        command=show_about_dialog
    )
    about_btn.pack(side="right", padx=10, pady=5)

    # --- Main Content Frame ---
    main_frame = ctk.CTkFrame(root, fg_color=BG_COLOR)
    main_frame.pack(fill="both", expand=True, padx=15, pady=10)

    # --- Hero Area: Deck Info (Always Visible) ---
    hero_frame = ctk.CTkFrame(main_frame, fg_color="white", corner_radius=8)
    hero_frame.pack(fill="x", padx=10, pady=(5, 10))
    
    # Empty state widgets
    empty_state_frame = ctk.CTkFrame(hero_frame, fg_color="transparent")
    empty_state_frame.pack(fill="x", padx=20, pady=15)
    
    # Loaded state widgets (hidden initially)
    loaded_state_frame = ctk.CTkFrame(hero_frame, fg_color="transparent")
    
    deck_name_label = ctk.CTkLabel(loaded_state_frame, text="", font=("Open Sans", 16, "bold"))
    deck_name_label.pack(anchor="w")
    
    deck_path_label = ctk.CTkLabel(loaded_state_frame, text="", font=("Open Sans", 11), text_color="gray50")
    deck_path_label.pack(anchor="w", pady=(2, 0))
    
    deck_modified_label = ctk.CTkLabel(loaded_state_frame, text="", font=("Open Sans", 11), text_color="gray50")
    deck_modified_label.pack(anchor="w", pady=(2, 0))
    
    def get_file_modified_date(filepath):
        """Get human-readable modified date for a file."""
        try:
            mtime = os.path.getmtime(filepath)
            from datetime import datetime
            dt = datetime.fromtimestamp(mtime)
            return dt.strftime("%b %d, %Y %I:%M %p")
        except:
            return "Unknown"
    
    # Store sections for use by split (declared here so browse_pptx can reset it)
    current_sections = []
    
    def update_hero_display():
        """Update hero area based on whether a deck is loaded."""
        deck_path = pptx_var.get()
        
        if deck_path and os.path.isfile(deck_path):
            # Show loaded state
            empty_state_frame.pack_forget()
            loaded_state_frame.pack(fill="x", padx=20, pady=15)
            
            filename = os.path.basename(deck_path)
            folder = os.path.dirname(deck_path)
            modified = get_file_modified_date(deck_path)
            
            deck_name_label.configure(text=filename)
            deck_path_label.configure(text=folder)
            deck_modified_label.configure(text=f"Modified: {modified}")
            browse_btn.configure(text="Change...")
        else:
            # Show empty state
            loaded_state_frame.pack_forget()
            empty_state_frame.pack(fill="x", padx=20, pady=15)
            browse_btn.configure(text="Browse...")
    
    def browse_pptx():
        path = filedialog.askopenfilename(
            title="Select PowerPoint File",
            filetypes=[("PowerPoint", "*.pptx"), ("All files", "*.*")]
        )
        if path:
            pptx_var.set(path)
            # Auto-set split output dir to same folder
            split_output_var.set(os.path.dirname(path))
            update_hero_display()
            # Clear any previous analysis
            nonlocal current_sections
            current_sections = []
    
    # Empty state content
    ctk.CTkLabel(
        empty_state_frame, 
        text="Select a PowerPoint file to begin", 
        font=("Open Sans", 13),
        text_color="gray50"
    ).pack(side="left", padx=(0, 15))
    
    browse_btn = ctk.CTkButton(empty_state_frame, text="Browse...", width=100, command=browse_pptx)
    browse_btn.pack(side="left")
    
    # Change button for loaded state (same command)
    change_btn = ctk.CTkButton(loaded_state_frame, text="Change...", width=100, command=browse_pptx)
    change_btn.pack(anchor="e", pady=(0, 5))

    # --- Tabbed Interface for Features ---
    tabview = ctk.CTkTabview(main_frame, width=850, height=390, fg_color=BG_COLOR, segmented_button_selected_color="#3b8ed0")
    tabview.pack(fill="both", expand=True, padx=10, pady=(5, 10))

    tab_split = tabview.add("Split Deck")
    tab_export = tabview.add("Export Notes")
    tab_import = tabview.add("Import Notes")
    tab_media = tabview.add("Media")
    tab_replace = tabview.add("Find/Replace")
    tab_misc = tabview.add("Misc")

    # ========== SPLIT TAB ==========
    split_frame = ctk.CTkFrame(tab_split, fg_color=BG_COLOR)
    split_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ctk.CTkLabel(
        split_frame,
        text="Split a master deck into separate files by PowerPoint sections.",
        font=("Open Sans", 13)
    ).pack(anchor="w", pady=(5, 10))

    # Output folder for split (tab-specific)
    split_output_frame = ctk.CTkFrame(split_frame, fg_color="transparent")
    split_output_frame.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(split_output_frame, text="Output Folder:", font=("Open Sans", 12)).pack(side="left", padx=(0, 10))
    split_output_entry = ctk.CTkEntry(split_output_frame, textvariable=split_output_var, width=450, font=("Open Sans", 12))
    split_output_entry.pack(side="left", padx=(0, 10))
    
    def browse_split_output():
        path = filedialog.askdirectory(title="Select Output Folder")
        if path:
            split_output_var.set(path)
    
    ctk.CTkButton(split_output_frame, text="Browse...", width=80, command=browse_split_output).pack(side="left", padx=(0, 5))
    
    def open_split_output():
        folder = split_output_var.get()
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("Open Folder", "Choose a valid output folder first.")
    
    ctk.CTkButton(split_output_frame, text="Open", width=60, command=open_split_output).pack(side="left")

    # Sections list (will be populated when deck is analyzed)
    sections_label = ctk.CTkLabel(split_frame, text="Sections found: (none)", font=("Open Sans", 12))
    sections_label.pack(anchor="w", pady=5)

    sections_listbox = tk.Listbox(split_frame, height=6, width=80, font=("Consolas", 11))
    sections_listbox.pack(fill="x", pady=5)

    split_btn_frame = ctk.CTkFrame(split_frame, fg_color="transparent")
    split_btn_frame.pack(fill="x", pady=10)

    def on_analyze_sections():
        nonlocal current_sections
        deck = pptx_var.get()
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        
        sections_listbox.delete(0, tk.END)
        set_busy("Analyzing sections...")
        
        try:
            current_sections = voxsplit.get_powerpoint_sections(deck)
            
            if not current_sections:
                sections_label.configure(text="Sections found: 0")
                sections_listbox.insert(tk.END, "(No sections found in deck)")
            else:
                sections_label.configure(text=f"Sections found: {len(current_sections)}")
                for i, (name, start, count) in enumerate(current_sections, 1):
                    display_name = name if name else "(unnamed)"
                    end_slide = start + count - 1
                    sections_listbox.insert(tk.END, f"{i}. {display_name} -- Slides {start}-{end_slide} ({count} slides)")
            
            set_ready()
        except Exception as e:
            current_sections = []
            sections_label.configure(text="Sections found: (error)")
            sections_listbox.insert(tk.END, f"Error: {e}")
            set_ready("Analysis failed")
            traceback.print_exc()

    def on_split_deck():
        nonlocal current_sections
        deck = pptx_var.get()
        output = split_output_var.get()
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not output:
            messagebox.showerror("Error", "Select an output folder first.")
            return
        
        # If sections not analyzed yet, do it now
        if not current_sections:
            set_busy("Analyzing sections first...")
            try:
                current_sections = voxsplit.get_powerpoint_sections(deck)
            except Exception as e:
                set_ready("Analysis failed")
                messagebox.showerror("Error", f"Failed to read sections: {e}")
                return
        
        if not current_sections:
            set_ready()
            messagebox.showinfo("No Sections", "No sections found in deck. Nothing to split.")
            return

        def log_line(msg):
            sections_listbox.insert(tk.END, msg)
            sections_listbox.see(tk.END)
            root.update_idletasks()

        set_busy("Splitting deck...")
        sections_listbox.delete(0, tk.END)

        result = do_split_deck(deck, output, sections=current_sections, log_callback=log_line)
        
        if result.get("success"):
            file_count = len(result.get('files', []))
            set_ready(f"Split complete: {file_count} files created")
            log_line(f"\n✓ Success! Created {file_count} files.")
            messagebox.showinfo("Success", f"Created {file_count} files in:\n{output}")
        else:
            set_ready("Split failed")
            messagebox.showerror("Error", result.get("error", "Unknown error"))

    btn_analyze = ctk.CTkButton(split_btn_frame, text="Analyze Sections", width=150, command=on_analyze_sections)
    btn_analyze.pack(side="left", padx=5)
    action_buttons.append(btn_analyze)
    
    btn_split = ctk.CTkButton(split_btn_frame, text="Split Deck", width=150, command=on_split_deck)
    btn_split.pack(side="left", padx=5)
    action_buttons.append(btn_split)

    # ========== EXPORT NOTES TAB ==========
    export_frame = ctk.CTkFrame(tab_export, fg_color=BG_COLOR)
    export_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ctk.CTkLabel(
        export_frame,
        text="Export speaker notes to a document for editing or review.",
        font=("Open Sans", 13)
    ).pack(anchor="w", pady=10)

    format_frame = ctk.CTkFrame(export_frame, fg_color="transparent")
    format_frame.pack(anchor="w", pady=10)

    ctk.CTkLabel(format_frame, text="Export Format:", font=("Open Sans", 12)).pack(side="left", padx=5)
    
    format_menu = ctk.CTkOptionMenu(
        format_frame,
        variable=export_format_var,
        values=["docx", "txt", "md"],
        width=120
    )
    format_menu.pack(side="left", padx=10)

    export_info = ctk.CTkLabel(
        export_frame,
        text="* docx: Word document with VO-friendly formatting (14pt, 1.5 spacing)\n"
             "* txt: Plain text file\n"
             "* md: Markdown (good for version control)",
        font=("Open Sans", 11),
        justify="left"
    )
    export_info.pack(anchor="w", pady=10)

    def on_export_notes():
        deck = pptx_var.get()
        fmt = export_format_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return

        # Build default filename
        base_name = os.path.splitext(os.path.basename(deck))[0]
        default_name = f"{base_name}_notes.{fmt}"
        default_dir = os.path.dirname(deck)
        
        # Format-specific file type
        filetypes = {
            "docx": [("Word Document", "*.docx"), ("All files", "*.*")],
            "txt": [("Text File", "*.txt"), ("All files", "*.*")],
            "md": [("Markdown", "*.md"), ("All files", "*.*")]
        }
        
        # Show Save As dialog
        output_file = filedialog.asksaveasfilename(
            title="Export Notes As",
            initialdir=default_dir,
            initialfile=default_name,
            defaultextension=f".{fmt}",
            filetypes=filetypes.get(fmt, [("All files", "*.*")])
        )
        
        if not output_file:
            return  # User cancelled

        set_busy(f"Exporting notes to {fmt}...")

        result = do_export_notes(deck, output_file, fmt)
        
        if result.get("success"):
            set_ready("Export complete")
            messagebox.showinfo("Success", f"Notes exported to:\n{output_file}")
        else:
            set_ready("Export failed")
            messagebox.showerror("Error", result.get("error", "Unknown error"))

    btn_export = ctk.CTkButton(export_frame, text="Export Notes...", width=150, command=on_export_notes)
    btn_export.pack(anchor="w", pady=10)
    action_buttons.append(btn_export)

    # ========== IMPORT NOTES TAB ==========
    import_frame = ctk.CTkFrame(tab_import, fg_color=BG_COLOR)
    import_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ctk.CTkLabel(
        import_frame,
        text="Import edited notes back into PowerPoint.",
        font=("Open Sans", 13)
    ).pack(anchor="w", pady=10)

    import_file_frame = ctk.CTkFrame(import_frame, fg_color="transparent")
    import_file_frame.pack(fill="x", pady=10)

    ctk.CTkLabel(import_file_frame, text="Notes File:", font=("Open Sans", 12)).pack(side="left", padx=5)
    import_entry = ctk.CTkEntry(import_file_frame, textvariable=import_file_var, width=400, font=("Open Sans", 12))
    import_entry.pack(side="left", padx=5)

    def browse_import_file():
        path = filedialog.askopenfilename(
            title="Select Notes File",
            filetypes=[
                ("All supported", "*.docx;*.txt;*.md"),
                ("Word", "*.docx"),
                ("Text", "*.txt"),
                ("Markdown", "*.md")
            ]
        )
        if path:
            import_file_var.set(path)

    ctk.CTkButton(import_file_frame, text="Browse...", width=100, command=browse_import_file).pack(side="left", padx=5)

    # Changes preview
    changes_label = ctk.CTkLabel(import_frame, text="Changes to apply: (none)", font=("Open Sans", 12))
    changes_label.pack(anchor="w", pady=5)

    changes_listbox = tk.Listbox(import_frame, height=6, width=80, font=("Consolas", 11))
    changes_listbox.pack(fill="x", pady=5)

    import_btn_frame = ctk.CTkFrame(import_frame, fg_color="transparent")
    import_btn_frame.pack(fill="x", pady=10)

    def on_preview_import():
        deck = pptx_var.get()
        notes_file = import_file_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not notes_file or not os.path.isfile(notes_file):
            messagebox.showerror("Error", "Select a valid notes file first.")
            return

        changes_listbox.delete(0, tk.END)
        set_busy("Analyzing changes...")

        result = do_import_notes(deck, notes_file, preview_only=True)
        
        if result.get("changes"):
            changes_label.configure(text=f"Changes to apply: {len(result['changes'])} slides")
            for change in result["changes"]:
                changes_listbox.insert(tk.END, f"Slide {change['slide']}: {change.get('word_diff', '?')} words changed")
        else:
            changes_label.configure(text="Changes to apply: (none)")
            changes_listbox.insert(tk.END, "(No changes detected or preview not available)")
        
        set_ready()

    def on_apply_import():
        deck = pptx_var.get()
        notes_file = import_file_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not notes_file or not os.path.isfile(notes_file):
            messagebox.showerror("Error", "Select a valid notes file first.")
            return

        if not messagebox.askyesno("Confirm", f"Apply changes to:\n{os.path.basename(deck)}\n\nThis will modify the PowerPoint file."):
            return

        def log_line(msg):
            changes_listbox.insert(tk.END, msg)
            changes_listbox.see(tk.END)
            root.update_idletasks()

        set_busy("Applying changes...")
        changes_listbox.delete(0, tk.END)

        result = do_import_notes(deck, notes_file, preview_only=False, log_callback=log_line)
        
        if result.get("success"):
            set_ready("Import complete")
            messagebox.showinfo("Success", f"Updated {len(result.get('slides_updated', []))} slides.")
        else:
            set_ready("Import failed")
            messagebox.showerror("Error", result.get("error", "Unknown error"))

    btn_preview_import = ctk.CTkButton(import_btn_frame, text="Preview Changes", width=150, command=on_preview_import)
    btn_preview_import.pack(side="left", padx=5)
    action_buttons.append(btn_preview_import)
    
    btn_apply_import = ctk.CTkButton(import_btn_frame, text="Apply Changes", width=150, command=on_apply_import)
    btn_apply_import.pack(side="left", padx=5)
    action_buttons.append(btn_apply_import)

    # ========== MEDIA TAB ==========
    media_frame = ctk.CTkFrame(tab_media, fg_color=BG_COLOR)
    media_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ctk.CTkLabel(
        media_frame,
        text="Extract, strip, or import audio/video from slides.",
        font=("Open Sans", 13)
    ).pack(anchor="w", pady=(5, 10))

    # Media folder picker
    media_folder_var = tk.StringVar()
    
    media_folder_frame = ctk.CTkFrame(media_frame, fg_color="transparent")
    media_folder_frame.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(media_folder_frame, text="Media Folder:", font=("Open Sans", 12)).pack(side="left", padx=(0, 10))
    media_folder_entry = ctk.CTkEntry(media_folder_frame, textvariable=media_folder_var, width=450, font=("Open Sans", 12))
    media_folder_entry.pack(side="left", padx=(0, 10))
    
    def browse_media_folder():
        path = filedialog.askdirectory(title="Select Media Folder")
        if path:
            media_folder_var.set(path)
    
    ctk.CTkButton(media_folder_frame, text="Browse...", width=80, command=browse_media_folder).pack(side="left", padx=(0, 5))
    
    def open_media_folder():
        folder = media_folder_var.get()
        if folder and os.path.isdir(folder):
            os.startfile(folder)
        else:
            messagebox.showinfo("Open Folder", "Choose a valid media folder first.")
    
    ctk.CTkButton(media_folder_frame, text="Open", width=60, command=open_media_folder).pack(side="left")

    # Log output
    media_log_label = ctk.CTkLabel(media_frame, text="Output:", font=("Open Sans", 12))
    media_log_label.pack(anchor="w", pady=(5, 2))

    media_listbox = tk.Listbox(media_frame, height=6, width=80, font=("Consolas", 11))
    media_listbox.pack(fill="x", pady=5)

    # Buttons frame
    media_btn_frame = ctk.CTkFrame(media_frame, fg_color="transparent")
    media_btn_frame.pack(fill="x", pady=10)

    def on_strip_audio():
        deck = pptx_var.get()
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        
        if not messagebox.askyesno("Confirm", f"Remove ALL audio from:\n{os.path.basename(deck)}\n\nThis will modify the PowerPoint file."):
            return
        
        def log_line(msg):
            media_listbox.insert(tk.END, msg)
            media_listbox.see(tk.END)
            root.update_idletasks()
        
        media_listbox.delete(0, tk.END)
        set_busy("Stripping audio...")
        
        try:
            result = voxmedia.strip_all_audio(deck, log_callback=log_line)
            if result.get("success"):
                count = result.get("audio_removed", 0)
                set_ready(f"Stripped {count} audio shape(s)")
                if count > 0:
                    messagebox.showinfo("Success", f"Removed {count} audio shape(s) from {len(result.get('slides_modified', []))} slide(s).")
                else:
                    messagebox.showinfo("No Audio", "No audio found in deck.")
            else:
                set_ready("Strip failed")
                messagebox.showerror("Error", "Strip operation failed.")
        except Exception as e:
            set_ready("Strip failed")
            media_listbox.insert(tk.END, f"Error: {e}")
            messagebox.showerror("Error", str(e))
            traceback.print_exc()

    def on_export_media():
        deck = pptx_var.get()
        media_folder = media_folder_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not media_folder:
            messagebox.showerror("Error", "Select a media folder first.")
            return
        
        def log_line(msg):
            media_listbox.insert(tk.END, msg)
            media_listbox.see(tk.END)
            root.update_idletasks()
        
        media_listbox.delete(0, tk.END)
        set_busy("Exporting media...")
        
        try:
            result = voxmedia.export_media(deck, media_folder, log_callback=log_line)
            if result.get("success"):
                count = result.get("files_exported", 0)
                set_ready(f"Exported {count} file(s)")
                if count > 0:
                    messagebox.showinfo("Success", f"Exported {count} file(s) to:\n{media_folder}")
                else:
                    messagebox.showinfo("No Media", "No embedded media found in deck.")
            else:
                set_ready("Export failed")
                messagebox.showerror("Error", "Export operation failed.")
        except Exception as e:
            set_ready("Export failed")
            media_listbox.insert(tk.END, f"Error: {e}")
            messagebox.showerror("Error", str(e))
            traceback.print_exc()

    def on_import_audio():
        deck = pptx_var.get()
        media_folder = media_folder_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not media_folder:
            messagebox.showerror("Error", "Select a media folder first.")
            return
        
        if not messagebox.askyesno("Confirm", f"Import audio from:\n{media_folder}\n\ninto:\n{os.path.basename(deck)}\n\nThis will modify the PowerPoint file."):
            return
        
        def log_line(msg):
            media_listbox.insert(tk.END, msg)
            media_listbox.see(tk.END)
            root.update_idletasks()
        
        media_listbox.delete(0, tk.END)
        set_busy("Importing audio...")
        
        try:
            result = voxmedia.import_audio(deck, media_folder, log_callback=log_line)
            if result.get("success"):
                count = result.get("files_imported", 0)
                set_ready(f"Imported {count} file(s)")
                if count > 0:
                    messagebox.showinfo("Success", f"Imported {count} audio file(s) to {len(result.get('slides_updated', []))} slide(s).")
                else:
                    messagebox.showinfo("No Audio", "No slideXX.wav files found in media folder.")
            else:
                set_ready("Import failed")
                messagebox.showerror("Error", "Import operation failed.")
        except Exception as e:
            set_ready("Import failed")
            media_listbox.insert(tk.END, f"Error: {e}")
            messagebox.showerror("Error", str(e))
            traceback.print_exc()

    btn_strip = ctk.CTkButton(media_btn_frame, text="Strip All Audio", width=130, command=on_strip_audio)
    btn_strip.pack(side="left", padx=5)
    action_buttons.append(btn_strip)
    
    btn_export_media = ctk.CTkButton(media_btn_frame, text="Export Media", width=130, command=on_export_media)
    btn_export_media.pack(side="left", padx=5)
    action_buttons.append(btn_export_media)
    
    btn_import_audio = ctk.CTkButton(media_btn_frame, text="Import Audio", width=130, command=on_import_audio)
    btn_import_audio.pack(side="left", padx=5)
    action_buttons.append(btn_import_audio)
    
    # Disable Import Audio if voxattach not available
    if not voxmedia.HAS_VOXATTACH:
        btn_import_audio.configure(state="disabled")
        ctk.CTkLabel(
            media_frame,
            text="(Import Audio requires voxattach.py)",
            font=("Open Sans", 10),
            text_color="gray50"
        ).pack(anchor="w")

    # ========== FIND/REPLACE TAB ==========
    replace_frame = ctk.CTkFrame(tab_replace, fg_color=BG_COLOR)
    replace_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ctk.CTkLabel(
        replace_frame,
        text="Find and replace text in speaker notes across all slides.",
        font=("Open Sans", 13)
    ).pack(anchor="w", pady=10)

    # Find field
    find_frame = ctk.CTkFrame(replace_frame, fg_color="transparent")
    find_frame.pack(fill="x", pady=5)
    ctk.CTkLabel(find_frame, text="Find:", font=("Open Sans", 12), width=80).pack(side="left", padx=5)
    find_entry = ctk.CTkEntry(find_frame, textvariable=find_var, width=500, font=("Open Sans", 12))
    find_entry.pack(side="left", padx=5)

    # Replace field
    replace_field_frame = ctk.CTkFrame(replace_frame, fg_color="transparent")
    replace_field_frame.pack(fill="x", pady=5)
    ctk.CTkLabel(replace_field_frame, text="Replace:", font=("Open Sans", 12), width=80).pack(side="left", padx=5)
    replace_entry = ctk.CTkEntry(replace_field_frame, textvariable=replace_var, width=500, font=("Open Sans", 12))
    replace_entry.pack(side="left", padx=5)

    # Options
    options_frame = ctk.CTkFrame(replace_frame, fg_color="transparent")
    options_frame.pack(fill="x", pady=5)
    case_check = ctk.CTkCheckBox(options_frame, text="Case sensitive", variable=case_sensitive_var)
    case_check.pack(side="left", padx=10)

    # Results
    results_label = ctk.CTkLabel(replace_frame, text="Matches: (none)", font=("Open Sans", 12))
    results_label.pack(anchor="w", pady=5)

    results_listbox = tk.Listbox(replace_frame, height=6, width=80, font=("Consolas", 11))
    results_listbox.pack(fill="x", pady=5)

    replace_btn_frame = ctk.CTkFrame(replace_frame, fg_color="transparent")
    replace_btn_frame.pack(fill="x", pady=10)

    def on_preview_replace():
        deck = pptx_var.get()
        find_text = find_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not find_text:
            messagebox.showerror("Error", "Enter text to find.")
            return

        results_listbox.delete(0, tk.END)
        set_busy("Searching...")

        result = do_find_replace(deck, find_text, replace_var.get(), 
                                  case_sensitive=case_sensitive_var.get(), 
                                  preview_only=True)
        
        if isinstance(result, list) and result:
            total_matches = sum(r.get("match_count", 0) for r in result)
            results_label.configure(text=f"Matches: {total_matches} in {len(result)} slides")
            for r in result:
                results_listbox.insert(tk.END, f"Slide {r['slide']}: {r.get('match_count', '?')} matches")
        else:
            results_label.configure(text="Matches: (none)")
            results_listbox.insert(tk.END, "(No matches found or preview not available)")
        
        set_ready()

    def on_apply_replace():
        deck = pptx_var.get()
        find_text = find_var.get()
        replace_text = replace_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        if not find_text:
            messagebox.showerror("Error", "Enter text to find.")
            return

        if not messagebox.askyesno("Confirm", f"Replace all occurrences of:\n'{find_text}'\nwith:\n'{replace_text}'\n\nThis will modify the PowerPoint file."):
            return

        def log_line(msg):
            results_listbox.insert(tk.END, msg)
            results_listbox.see(tk.END)
            root.update_idletasks()

        set_busy("Replacing...")
        results_listbox.delete(0, tk.END)

        result = do_find_replace(deck, find_text, replace_text,
                                  case_sensitive=case_sensitive_var.get(),
                                  preview_only=False,
                                  log_callback=log_line)
        
        if result.get("success"):
            set_ready("Replace complete")
            messagebox.showinfo("Success", f"Made {result.get('total_replacements', '?')} replacements in {len(result.get('slides_modified', []))} slides.")
        else:
            set_ready("Replace failed")
            messagebox.showerror("Error", result.get("error", "Unknown error"))

    btn_preview_replace = ctk.CTkButton(replace_btn_frame, text="Preview", width=150, command=on_preview_replace)
    btn_preview_replace.pack(side="left", padx=5)
    action_buttons.append(btn_preview_replace)
    
    btn_apply_replace = ctk.CTkButton(replace_btn_frame, text="Replace All", width=150, command=on_apply_replace)
    btn_apply_replace.pack(side="left", padx=5)
    action_buttons.append(btn_apply_replace)

    # ========== MISC TAB ==========
    misc_frame = ctk.CTkFrame(tab_misc, fg_color=BG_COLOR)
    misc_frame.pack(fill="both", expand=True, padx=10, pady=10)

    ctk.CTkLabel(
        misc_frame,
        text="Miscellaneous deck utilities. Use with caution - these modify your deck!",
        font=("Open Sans", 13),
        text_color="gray40"
    ).pack(anchor="w", pady=(5, 15))

    # --- Strip Animations Section ---
    anim_section = ctk.CTkFrame(misc_frame, fg_color="transparent")
    anim_section.pack(fill="x", pady=(0, 15))
    
    ctk.CTkLabel(
        anim_section,
        text="Strip All Animations",
        font=("Open Sans", 13, "bold")
    ).pack(anchor="w")
    
    ctk.CTkLabel(
        anim_section,
        text="Remove every animation effect from all slides. Useful for cleaning up SME decks.",
        font=("Open Sans", 11),
        text_color="gray50"
    ).pack(anchor="w", pady=(2, 8))
    
    def on_strip_animations():
        deck = pptx_var.get()
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        
        if not messagebox.askyesno("Confirm", f"Remove ALL animations from:\n{os.path.basename(deck)}\n\nThis cannot be undone!"):
            return
        
        def log_line(msg):
            misc_listbox.insert(tk.END, msg)
            misc_listbox.see(tk.END)
            root.update_idletasks()
        
        misc_listbox.delete(0, tk.END)
        set_busy("Stripping animations...")
        
        try:
            result = voxmisc.strip_all_animations(deck, log_callback=log_line)
            if result.get("success"):
                count = result.get("effects_removed", 0)
                set_ready(f"Stripped {count} animation(s)")
                if count > 0:
                    messagebox.showinfo("Success", f"Removed {count} animation(s) from {len(result.get('slides_modified', []))} slide(s).")
                else:
                    messagebox.showinfo("No Animations", "No animations found in deck.")
            else:
                set_ready("Strip failed")
                messagebox.showerror("Error", "Strip operation failed.")
        except Exception as e:
            set_ready("Strip failed")
            misc_listbox.insert(tk.END, f"Error: {e}")
            messagebox.showerror("Error", str(e))
            traceback.print_exc()
    
    btn_strip_anim = ctk.CTkButton(anim_section, text="Strip All Animations", width=160, command=on_strip_animations)
    btn_strip_anim.pack(anchor="w")
    action_buttons.append(btn_strip_anim)

    # --- Font Normalization Section ---
    font_section = ctk.CTkFrame(misc_frame, fg_color="transparent")
    font_section.pack(fill="x", pady=(0, 10))
    
    ctk.CTkLabel(
        font_section,
        text="Font Normalization",
        font=("Open Sans", 13, "bold")
    ).pack(anchor="w")
    
    ctk.CTkLabel(
        font_section,
        text="Force all text to a single font. Preserves size, bold, italic, and other formatting.",
        font=("Open Sans", 11),
        text_color="gray50"
    ).pack(anchor="w", pady=(2, 8))
    
    font_control_frame = ctk.CTkFrame(font_section, fg_color="transparent")
    font_control_frame.pack(fill="x")
    
    ctk.CTkLabel(font_control_frame, text="Target Font:", font=("Open Sans", 12)).pack(side="left", padx=(0, 10))
    
    target_font_var = tk.StringVar(value="Arial")
    font_options = [
        # System / Windows defaults
        "Arial",
        "Calibri", 
        "Segoe UI",
        "Tahoma",
        "Verdana",
        "Consolas",
        # Roboto family
        "Roboto",
        "Roboto Condensed",
        "Roboto Slab",
        "Roboto Mono",
        "Roboto Light",
        # Google / Web fonts
        "Open Sans",
        "Lato",
        "Montserrat",
        "Source Sans Pro",
        "Poppins",
        "Nunito",
        "Raleway",
        "Oswald",
        "Inter",
        # Classic serif
        "Times New Roman",
        "Georgia",
        "Cambria",
        "Garamond",
        # Corporate / Clean
        "Helvetica",
        "Helvetica Neue",
        "Century Gothic",
        "Franklin Gothic",
        "Gill Sans",
        "Futura",
        "Avenir",
        "Proxima Nova",
    ]
    font_menu = ctk.CTkOptionMenu(font_control_frame, variable=target_font_var, values=font_options, width=180)
    font_menu.pack(side="left", padx=(0, 10))
    
    def on_analyze_fonts():
        deck = pptx_var.get()
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        
        def log_line(msg):
            misc_listbox.insert(tk.END, msg)
            misc_listbox.see(tk.END)
            root.update_idletasks()
        
        misc_listbox.delete(0, tk.END)
        set_busy("Analyzing fonts...")
        
        try:
            result = voxmisc.analyze_fonts(deck, log_callback=log_line)
            if result.get("success"):
                set_ready(f"Found {len(result.get('fonts', {}))} font(s)")
            else:
                set_ready("Analysis failed")
        except Exception as e:
            set_ready("Analysis failed")
            misc_listbox.insert(tk.END, f"Error: {e}")
            traceback.print_exc()
    
    def on_normalize_fonts():
        deck = pptx_var.get()
        target = target_font_var.get()
        
        if not deck or not os.path.isfile(deck):
            messagebox.showerror("Error", "Select a valid PowerPoint file first.")
            return
        
        if not messagebox.askyesno("Confirm", f"Change ALL fonts to '{target}' in:\n{os.path.basename(deck)}\n\nThis cannot be undone!"):
            return
        
        def log_line(msg):
            misc_listbox.insert(tk.END, msg)
            misc_listbox.see(tk.END)
            root.update_idletasks()
        
        misc_listbox.delete(0, tk.END)
        set_busy(f"Normalizing to {target}...")
        
        try:
            result = voxmisc.normalize_fonts(deck, target, log_callback=log_line)
            if result.get("success"):
                count = result.get("runs_changed", 0)
                set_ready(f"Changed {count} text run(s)")
                if count > 0:
                    messagebox.showinfo("Success", f"Changed {count} text run(s) to '{target}' in {len(result.get('slides_modified', []))} slide(s).")
                else:
                    messagebox.showinfo("No Changes", f"All text already using '{target}' or no text found.")
            else:
                set_ready("Normalize failed")
                messagebox.showerror("Error", "Normalize operation failed.")
        except Exception as e:
            set_ready("Normalize failed")
            misc_listbox.insert(tk.END, f"Error: {e}")
            messagebox.showerror("Error", str(e))
            traceback.print_exc()
    
    btn_analyze_fonts = ctk.CTkButton(font_control_frame, text="Analyze", width=100, command=on_analyze_fonts)
    btn_analyze_fonts.pack(side="left", padx=(0, 5))
    action_buttons.append(btn_analyze_fonts)
    
    btn_normalize = ctk.CTkButton(font_control_frame, text="Normalize", width=100, command=on_normalize_fonts)
    btn_normalize.pack(side="left")
    action_buttons.append(btn_normalize)

    # Log output for Misc tab
    misc_log_label = ctk.CTkLabel(misc_frame, text="Output:", font=("Open Sans", 12))
    misc_log_label.pack(anchor="w", pady=(10, 2))

    misc_listbox = tk.Listbox(misc_frame, height=5, width=80, font=("Consolas", 11))
    misc_listbox.pack(fill="x", pady=5)

    # --- Status Bar ---
    status_frame = ctk.CTkFrame(root, height=30, corner_radius=0, fg_color=BG_COLOR)
    status_frame.pack(fill="x", side="bottom")
    status_label = ctk.CTkLabel(status_frame, textvariable=status_var, font=("Open Sans", 11))
    status_label.pack(side="left", padx=15, pady=5)

    # --- Window Close Handler ---
    def on_close():
        # No persistence - fresh start each launch
        try:
            _release_single_instance(_SINGLE_LOCK)
        except Exception:
            pass
        root.destroy()

    root.protocol("WM_DELETE_WINDOW", on_close)

    # --- Start ---
    root.mainloop()


# --- Entry Point ---
if __name__ == "__main__":
    try:
        already_open, _SINGLE_LOCK = _check_single_instance()
    except Exception as e:
        already_open, _SINGLE_LOCK = False, None
        print(f"Single-instance check failed: {e}")

    if already_open:
        try:
            r = tk.Tk()
            r.withdraw()
            messagebox.showinfo("Already running", f"{APP_NAME} is already open.")
            r.destroy()
        except Exception:
            pass
        sys.exit(0)

    try:
        main()
    except Exception:
        try:
            log_dir = os.path.join(os.getenv("APPDATA") or tempfile.gettempdir(), APP_NAME)
            os.makedirs(log_dir, exist_ok=True)
            log_path = os.path.join(log_dir, "crash.log")
            with open(log_path, "w", encoding="utf-8") as f:
                f.write(traceback.format_exc())
            r = tk.Tk()
            r.withdraw()
            messagebox.showerror(f"{APP_NAME} - Error", f"An error occurred.\n\nDetails written to:\n{log_path}")
            r.destroy()
        finally:
            try:
                _release_single_instance(_SINGLE_LOCK)
            except Exception:
                pass
