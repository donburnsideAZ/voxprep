"""
voxmisc.py
Miscellaneous PowerPoint deck utilities for VoxPrep.

Dangerous but useful operations:
- Strip all animations from a deck
- Normalize all fonts to a single family

Example:
    # Remove all animations
    strip_all_animations("Training.pptx")
    
    # Force all text to Arial
    normalize_fonts("Training.pptx", "Arial")
"""

import os
import time
from pathlib import Path
from typing import Dict, Optional, Callable, List

try:
    from win32com.client import Dispatch, gencache
    import pywintypes
    HAS_COM = True
except Exception:
    HAS_COM = False

try:
    import win32api
    HAS_WIN32API = True
except Exception:
    HAS_WIN32API = False

LOG_PREFIX = "[voxmisc]"

# COM retry settings
COM_RETRY_ATTEMPTS = 3
COM_RETRY_DELAY = 1.5


def log(msg: str):
    """Simple logging helper."""
    print(f"{LOG_PREFIX} {msg}", flush=True)


def open_presentation_with_retry(app, pptx_path: str, read_only: bool = True, max_attempts: int = COM_RETRY_ATTEMPTS):
    """Open a PowerPoint presentation with retry logic."""
    last_error = None
    
    for attempt in range(1, max_attempts + 1):
        try:
            pres = app.Presentations.Open(pptx_path, read_only, False, False)
            return pres
        except Exception as e:
            last_error = e
            if attempt < max_attempts:
                log(f"  Open failed (attempt {attempt}/{max_attempts}), retrying in {COM_RETRY_DELAY}s...")
                time.sleep(COM_RETRY_DELAY)
            else:
                log(f"  Open failed after {max_attempts} attempts")
    
    raise RuntimeError(f"Failed to open presentation after {max_attempts} attempts: {last_error}")


def get_short_path(long_path: str) -> str:
    """Convert path to Windows 8.3 short path format if possible."""
    if not HAS_WIN32API:
        return long_path
    try:
        long_path = long_path.replace('/', '\\')
        return win32api.GetShortPathName(long_path)
    except Exception:
        return long_path


# ============================================================
# STRIP ALL ANIMATIONS
# ============================================================

def strip_all_animations(pptx_path: str, log_callback: Optional[Callable] = None) -> Dict:
    """
    Remove ALL animation effects from all slides in a PowerPoint deck.
    
    This removes:
    - Entrance/exit/emphasis/motion path animations
    - Trigger animations
    - All effects from the main sequence and interactive sequences
    
    Args:
        pptx_path: Path to the PowerPoint file
        log_callback: Optional function for progress logging
    
    Returns:
        Dict with 'success', 'slides_modified', 'effects_removed' counts
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available (pywin32 not installed)")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=False)
        slide_count = pres.Slides.Count
        
        _log(f"Scanning {slide_count} slides for animations...")
        
        total_removed = 0
        slides_modified = []
        
        for i in range(1, slide_count + 1):
            slide = pres.Slides.Item(i)
            removed_this_slide = 0
            
            try:
                # Get the timeline
                timeline = slide.TimeLine
                
                # Clear main sequence (entrance, exit, emphasis, motion animations)
                main_seq = timeline.MainSequence
                effect_count = main_seq.Count
                
                if effect_count > 0:
                    # Delete all effects (iterate backwards)
                    for j in range(effect_count, 0, -1):
                        try:
                            main_seq.Item(j).Delete()
                            removed_this_slide += 1
                        except Exception:
                            pass
                
                # Clear interactive sequences (trigger-based animations)
                try:
                    interactive_seqs = timeline.InteractiveSequences
                    seq_count = interactive_seqs.Count
                    
                    if seq_count > 0:
                        # Delete all interactive sequences
                        for k in range(seq_count, 0, -1):
                            try:
                                seq = interactive_seqs.Item(k)
                                eff_count = seq.Count
                                removed_this_slide += eff_count
                                # Delete the entire sequence
                                interactive_seqs.Item(k).Delete()
                            except Exception:
                                pass
                except Exception:
                    # InteractiveSequences might not exist
                    pass
                
            except Exception as e:
                _log(f"  Slide {i}: error accessing timeline - {e}")
            
            if removed_this_slide > 0:
                slides_modified.append(i)
                total_removed += removed_this_slide
                _log(f"  Slide {i}: removed {removed_this_slide} animation(s)")
        
        if total_removed > 0:
            pres.Save()
            _log(f"Saved. Removed {total_removed} animation(s) from {len(slides_modified)} slide(s).")
        else:
            _log("No animations found in deck.")
        
        return {
            "success": True,
            "effects_removed": total_removed,
            "slides_modified": slides_modified
        }
        
    except Exception as e:
        raise RuntimeError(f"Strip animations failed: {e}")
    
    finally:
        if pres is not None:
            try:
                pres.Close()
            except:
                pass
        if app is not None:
            try:
                app.Quit()
            except:
                pass


# ============================================================
# ANALYZE FONTS
# ============================================================

def analyze_fonts(pptx_path: str, log_callback: Optional[Callable] = None) -> Dict:
    """
    Analyze all fonts used in a PowerPoint deck.
    
    Args:
        pptx_path: Path to the PowerPoint file
        log_callback: Optional function for progress logging
    
    Returns:
        Dict with 'success', 'fonts' (dict of font_name -> count), 'total_runs'
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available (pywin32 not installed)")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=True)
        slide_count = pres.Slides.Count
        
        _log(f"Analyzing fonts in {slide_count} slides...")
        
        fonts = {}
        total_runs = 0
        
        for i in range(1, slide_count + 1):
            slide = pres.Slides.Item(i)
            
            # Iterate through all shapes
            for j in range(1, slide.Shapes.Count + 1):
                try:
                    shape = slide.Shapes.Item(j)
                    _analyze_shape_fonts(shape, fonts)
                except Exception:
                    pass
        
        # Count total text runs
        total_runs = sum(fonts.values())
        
        _log(f"Found {len(fonts)} unique font(s) across {total_runs} text run(s)")
        for font_name, count in sorted(fonts.items(), key=lambda x: -x[1]):
            _log(f"  {font_name}: {count}")
        
        return {
            "success": True,
            "fonts": fonts,
            "total_runs": total_runs
        }
        
    except Exception as e:
        raise RuntimeError(f"Analyze fonts failed: {e}")
    
    finally:
        if pres is not None:
            try:
                pres.Close()
            except:
                pass
        if app is not None:
            try:
                app.Quit()
            except:
                pass


def _analyze_shape_fonts(shape, fonts: Dict):
    """Recursively analyze fonts in a shape and its children."""
    try:
        # Check if shape has a text frame
        if shape.HasTextFrame:
            text_frame = shape.TextFrame
            if text_frame.HasText:
                text_range = text_frame.TextRange
                
                # Iterate through runs (text with consistent formatting)
                for k in range(1, text_range.Runs().Count + 1):
                    try:
                        run = text_range.Runs(k)
                        font_name = run.Font.Name
                        if font_name:
                            fonts[font_name] = fonts.get(font_name, 0) + 1
                    except Exception:
                        pass
        
        # Check if shape is a group
        if shape.Type == 6:  # msoGroup
            for k in range(1, shape.GroupItems.Count + 1):
                try:
                    _analyze_shape_fonts(shape.GroupItems.Item(k), fonts)
                except Exception:
                    pass
        
        # Check if shape is a table
        if shape.HasTable:
            table = shape.Table
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count + 1):
                    try:
                        cell = table.Cell(row, col)
                        if cell.Shape.HasTextFrame:
                            text_range = cell.Shape.TextFrame.TextRange
                            for k in range(1, text_range.Runs().Count + 1):
                                try:
                                    run = text_range.Runs(k)
                                    font_name = run.Font.Name
                                    if font_name:
                                        fonts[font_name] = fonts.get(font_name, 0) + 1
                                except Exception:
                                    pass
                    except Exception:
                        pass
                        
    except Exception:
        pass


# ============================================================
# NORMALIZE FONTS
# ============================================================

def normalize_fonts(pptx_path: str, target_font: str, log_callback: Optional[Callable] = None) -> Dict:
    """
    Change all fonts in a PowerPoint deck to a single target font.
    
    Preserves other formatting (size, bold, italic, color, etc.).
    
    Args:
        pptx_path: Path to the PowerPoint file
        target_font: Name of the font to apply (e.g., "Arial", "Calibri")
        log_callback: Optional function for progress logging
    
    Returns:
        Dict with 'success', 'slides_modified', 'runs_changed'
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available (pywin32 not installed)")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=False)
        slide_count = pres.Slides.Count
        
        _log(f"Normalizing fonts to '{target_font}' across {slide_count} slides...")
        
        total_changed = 0
        slides_modified = set()
        
        for i in range(1, slide_count + 1):
            slide = pres.Slides.Item(i)
            changed_this_slide = 0
            
            # Iterate through all shapes
            for j in range(1, slide.Shapes.Count + 1):
                try:
                    shape = slide.Shapes.Item(j)
                    changed = _normalize_shape_fonts(shape, target_font)
                    changed_this_slide += changed
                except Exception:
                    pass
            
            if changed_this_slide > 0:
                slides_modified.add(i)
                total_changed += changed_this_slide
        
        if total_changed > 0:
            pres.Save()
            _log(f"Saved. Changed {total_changed} text run(s) to '{target_font}' in {len(slides_modified)} slide(s).")
        else:
            _log(f"All text already using '{target_font}' or no text found.")
        
        return {
            "success": True,
            "runs_changed": total_changed,
            "slides_modified": list(slides_modified)
        }
        
    except Exception as e:
        raise RuntimeError(f"Normalize fonts failed: {e}")
    
    finally:
        if pres is not None:
            try:
                pres.Close()
            except:
                pass
        if app is not None:
            try:
                app.Quit()
            except:
                pass


def _normalize_shape_fonts(shape, target_font: str) -> int:
    """Recursively normalize fonts in a shape and its children. Returns count of changes."""
    changed = 0
    
    try:
        # Check if shape has a text frame
        if shape.HasTextFrame:
            text_frame = shape.TextFrame
            if text_frame.HasText:
                text_range = text_frame.TextRange
                
                # Iterate through runs
                for k in range(1, text_range.Runs().Count + 1):
                    try:
                        run = text_range.Runs(k)
                        current_font = run.Font.Name
                        if current_font and current_font != target_font:
                            run.Font.Name = target_font
                            changed += 1
                    except Exception:
                        pass
        
        # Check if shape is a group
        if shape.Type == 6:  # msoGroup
            for k in range(1, shape.GroupItems.Count + 1):
                try:
                    changed += _normalize_shape_fonts(shape.GroupItems.Item(k), target_font)
                except Exception:
                    pass
        
        # Check if shape is a table
        if shape.HasTable:
            table = shape.Table
            for row in range(1, table.Rows.Count + 1):
                for col in range(1, table.Columns.Count + 1):
                    try:
                        cell = table.Cell(row, col)
                        if cell.Shape.HasTextFrame:
                            text_range = cell.Shape.TextFrame.TextRange
                            for k in range(1, text_range.Runs().Count + 1):
                                try:
                                    run = text_range.Runs(k)
                                    current_font = run.Font.Name
                                    if current_font and current_font != target_font:
                                        run.Font.Name = target_font
                                        changed += 1
                                except Exception:
                                    pass
                    except Exception:
                        pass
                        
    except Exception:
        pass
    
    return changed


# ============================================================
# CLI
# ============================================================

def _usage():
    print("Usage:")
    print("  python voxmisc.py strip-animations <deck.pptx>")
    print("  python voxmisc.py analyze-fonts <deck.pptx>")
    print("  python voxmisc.py normalize-fonts <deck.pptx> <target_font>")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        _usage()
        sys.exit(64)
    
    command = sys.argv[1].lower()
    deck_path = sys.argv[2]
    
    try:
        if command == "strip-animations":
            result = strip_all_animations(deck_path)
            print(f"Removed {result['effects_removed']} animation(s)")
            
        elif command == "analyze-fonts":
            result = analyze_fonts(deck_path)
            print(f"Found {len(result['fonts'])} unique font(s)")
            
        elif command == "normalize-fonts":
            if len(sys.argv) < 4:
                print("Error: target_font required")
                sys.exit(64)
            target_font = sys.argv[3]
            result = normalize_fonts(deck_path, target_font)
            print(f"Changed {result['runs_changed']} text run(s) to '{target_font}'")
            
        else:
            print(f"Unknown command: {command}")
            _usage()
            sys.exit(64)
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)
