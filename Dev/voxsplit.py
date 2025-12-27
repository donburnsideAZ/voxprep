"""
voxsplit.py
PowerPoint deck splitting by sections for Voxsmith.

Reads native PowerPoint sections and creates separate .pptx files for each section.
Named sections use their title as filename; unnamed sections get sequential numbering.

Example:
    Section 1: "Introduction" â†’ Introduction.pptx
    Section 2: (unnamed) â†’ Deck-1.pptx
    Section 3: "Chapter 1" â†’ Chapter-1.pptx
"""

import os
import re
import shutil
from pathlib import Path

try:
    from win32com.client import Dispatch
    import pywintypes
    HAS_COM = True
except Exception:
    HAS_COM = False

try:
    import win32api
    HAS_WIN32API = True
except Exception:
    HAS_WIN32API = False

LOG_PREFIX = "[voxsplit]"


def log(msg: str):
    """Simple logging helper."""
    print(f"{LOG_PREFIX} {msg}", flush=True)


def get_short_path(long_path: str) -> str:
    """
    Convert long path with spaces to Windows 8.3 short path format.
    This is needed because PowerPoint COM doesn't like spaces in paths.
    
    Args:
        long_path: Path that may contain spaces
    
    Returns:
        Short path without spaces (if win32api available), otherwise original
    
    Example:
        "C:/test/Chunking test/file.pptx" â†’ "C:/test/CHUNKI~1/file.pptx"
    """
    if not HAS_WIN32API:
        return long_path
    
    try:
        # Convert forward slashes to backslashes
        long_path = long_path.replace('/', '\\')
        # Get short path
        short_path = win32api.GetShortPathName(long_path)
        return short_path
    except Exception:
        # If conversion fails, return original
        return long_path


def sanitize_filename(name: str) -> str:
    """
    Sanitize section name for use as filename.
    
    Args:
        name: Section name from PowerPoint
    
    Returns:
        Safe filename (without extension)
    
    Examples:
        "Chapter 1: Overview" â†’ "Chapter-1-Overview"
        "What's Next?" â†’ "Whats-Next"
        "  Intro  " â†’ "Intro"
    """
    # Remove/replace forbidden Windows filename characters
    forbidden = r'[<>:"/\\|?*]'
    safe = re.sub(forbidden, '', name)
    
    # Replace spaces with dashes
    safe = safe.replace(' ', '-')
    
    # Collapse multiple dashes
    safe = re.sub(r'-+', '-', safe)
    
    # Remove leading/trailing dashes
    safe = safe.strip('-')
    
    # Truncate if too long (Windows max path component is 255)
    if len(safe) > 50:
        safe = safe[:50].rstrip('-')
    
    return safe or "Untitled"


def is_unnamed_section(name: str) -> bool:
    """
    Check if section name is empty or default.
    
    Args:
        name: Section name from PowerPoint
    
    Returns:
        True if section should be considered unnamed
    """
    if not name or not name.strip():
        return True
    
    # PowerPoint's default section names
    defaults = [
        "default section",
        "untitled section",
        "section"
    ]
    
    return name.strip().lower() in defaults


def get_powerpoint_sections(pptx_path: str) -> list:
    """
    Read PowerPoint sections via COM API.
    
    Args:
        pptx_path: Path to .pptx file
    
    Returns:
        List of tuples: (section_name, start_slide_index, slide_count)
        Example: [("Introduction", 1, 5), ("Chapter 1", 6, 10)]
    
    Raises:
        RuntimeError: If COM not available or file can't be opened
    """
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available (pywin32 not installed)")
    
    # Convert to short path to avoid spaces
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    try:
        app = Dispatch("PowerPoint.Application")
        app.Visible = True  # Must be visible - COM restriction
        
        # Open presentation (read-only)
        pres = app.Presentations.Open(pptx_path, True, False, False)  # True = read-only
        
        sections = []
        section_props = pres.SectionProperties
        section_count = section_props.Count
        
        if section_count == 0:
            # No sections defined - treat entire deck as one section
            total_slides = pres.Slides.Count
            sections.append(("", 1, total_slides))
        else:
            # Read each section
            for i in range(1, section_count + 1):
                name = section_props.Name(i)
                first_slide = section_props.FirstSlide(i)
                slides_in_section = section_props.SlidesCount(i)
                
                sections.append((name, first_slide, slides_in_section))
        
        # Close presentation without saving
        pres.Close()
        app.Quit()
        
        return sections
        
    except Exception as e:
        try:
            pres.Close()
            app.Quit()
        except:
            pass
        raise RuntimeError(f"Failed to read sections: {e}")


def split_deck_by_sections(pptx_path: str, sections: list, output_dir: str = None, log_callback=None) -> list:
    """
    Create separate .pptx files for each section.
    
    Args:
        pptx_path: Path to master deck
        sections: List of (name, start_slide, slide_count) from get_powerpoint_sections()
        output_dir: Where to save chunks (default: same folder as master)
        log_callback: Optional function to call for logging (e.g., log_line from UI)
    
    Returns:
        List of created file paths
    
    Examples:
        sections = [("Introduction", 1, 5), ("", 6, 10), ("Chapter 1", 11, 15)]
        files = split_deck_by_sections("Master.pptx", sections)
        # Creates: Introduction.pptx, Deck-1.pptx, Chapter-1.pptx
    """
    def _log(msg):
        """Log to callback if provided, otherwise print."""
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available (pywin32 not installed)")
    
    # Convert to short path to avoid spaces
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    # Default output dir = same folder as source
    if output_dir is None:
        output_dir = os.path.dirname(pptx_path)
    
    # Also convert output dir to short path
    output_dir = get_short_path(output_dir)
    
    os.makedirs(output_dir, exist_ok=True)
    
    try:
        # Open PowerPoint (must be visible - COM doesn't allow hiding)
        # PowerPoint.Application.Visible = False raises COM error -2147352567
        # This is a PowerPoint-specific limitation (unlike Excel)
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        # Open master presentation
        master = app.Presentations.Open(pptx_path, False, False, False)
        
        created_files = []
        unnamed_counter = 1
        
        for section_name, first_slide, slide_count in sections:
            try:
                # Determine output filename
                if is_unnamed_section(section_name):
                    filename = f"Deck-{unnamed_counter}.pptx"
                    unnamed_counter += 1
                else:
                    safe_name = sanitize_filename(section_name)
                    filename = f"{safe_name}.pptx"
                
                output_path = os.path.join(output_dir, filename)
                
                # Handle filename conflicts
                if os.path.exists(output_path):
                    base = os.path.splitext(filename)[0]
                    counter = 2
                    while os.path.exists(output_path):
                        filename = f"{base}_{counter}.pptx"
                        output_path = os.path.join(output_dir, filename)
                        counter += 1
                
                # Convert final output path to short path for COM
                output_path = get_short_path(output_path)
                
                _log(f"Creating {filename} (slides {first_slide}-{first_slide + slide_count - 1})...")
                
                # Method: Use Python to copy file, then open and edit
                # This avoids PowerPoint's SaveCopyAs which is problematic
                try:
                    # Use Python to copy the file (works with spaces, no COM issues)
                    shutil.copy2(pptx_path, output_path)
                    _log(f"  Copied master deck")
                    
                    # Convert to short path for opening
                    output_path_short = get_short_path(output_path)
                    
                    # Open the copy
                    chunk_pres = app.Presentations.Open(output_path_short, False, False, False)
                    _log(f"  Opened copy for editing")
                    
                    # Delete slides OUTSIDE the section range
                    # Work backwards to avoid index shifting
                    last_slide = first_slide + slide_count - 1
                    total_slides = chunk_pres.Slides.Count
                    
                    # Delete slides after the section
                    for i in range(total_slides, last_slide, -1):
                        try:
                            chunk_pres.Slides.Item(i).Delete()
                        except:
                            pass
                    
                    # Delete slides before the section
                    for i in range(first_slide - 1, 0, -1):
                        try:
                            chunk_pres.Slides.Item(i).Delete()
                        except:
                            pass
                    
                    # Save and close
                    chunk_pres.Save()
                    final_count = chunk_pres.Slides.Count
                    chunk_pres.Close()
                    
                    created_files.append(output_path)
                    _log(f"  Saved: {filename} ({final_count} slides)")
                    
                except Exception as save_error:
                    _log(f"  Error: {save_error}")
                    # Try to clean up partial file
                    try:
                        if os.path.exists(output_path):
                            os.remove(output_path)
                    except:
                        pass
                    continue
                
            except Exception as e:
                _log(f"  Error creating {filename}: {e}")
                try:
                    new_pres.Close()
                except:
                    pass
                continue
        
        # Close master presentation
        master.Close()
        app.Quit()
        
        return created_files
        
    except Exception as e:
        try:
            master.Close()
            app.Quit()
        except:
            pass
        raise RuntimeError(f"Split operation failed: {e}")


# --- CLI for standalone testing ---

def _usage():
    print("Usage: python voxsplit.py <deck.pptx> [output_dir]")
    print()
    print("Splits a PowerPoint deck by sections into separate files.")
    print("Output files saved in same folder as source (or output_dir if specified).")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        _usage()
        sys.exit(64)
    
    deck_path = sys.argv[1]
    output_dir = sys.argv[2] if len(sys.argv) > 2 else None
    
    try:
        log(f"Reading sections from: {deck_path}")
        sections = get_powerpoint_sections(deck_path)
        
        if not sections:
            log("No sections found in deck")
            sys.exit(1)
        
        log(f"Found {len(sections)} section(s):")
        for name, start, count in sections:
            display_name = f'"{name}"' if name else "(unnamed)"
            log(f"  {display_name}: slides {start}-{start + count - 1}")
        
        log("\nSplitting deck...")
        created = split_deck_by_sections(deck_path, sections, output_dir)
        
        log(f"\nSuccess! Created {len(created)} file(s):")
        for path in created:
            log(f"  {os.path.basename(path)}")
        
        sys.exit(0)
        
    except Exception as e:
        log(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)
