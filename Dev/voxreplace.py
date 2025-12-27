"""
voxreplace.py
Find and replace in PowerPoint speaker notes for VoxPrep.

Searches across all slide notes and performs bulk replacements.
Supports case-sensitive and regex matching.

Example:
    # Preview changes
    matches = find_in_notes("Training.pptx", "Acme Corp")
    print(f"Found {len(matches)} matches")
    
    # Preview replacement
    preview = preview_replace("Training.pptx", "Acme Corp", "Acme Industries")
    
    # Apply replacement
    result = replace_in_notes("Training.pptx", "Acme Corp", "Acme Industries")
"""

import os
import re
import time
from pathlib import Path
from typing import List, Dict, Optional, Callable, Tuple

try:
    from win32com.client import Dispatch, gencache
    HAS_COM = True
except Exception:
    HAS_COM = False

try:
    import win32api
    HAS_WIN32API = True
except Exception:
    HAS_WIN32API = False

LOG_PREFIX = "[voxreplace]"

# COM retry settings
COM_RETRY_ATTEMPTS = 3
COM_RETRY_DELAY = 1.5  # seconds


def log(msg: str):
    """Simple logging helper."""
    print(f"{LOG_PREFIX} {msg}", flush=True)


def open_presentation_with_retry(app, pptx_path: str, read_only: bool = True, max_attempts: int = COM_RETRY_ATTEMPTS):
    """
    Open a PowerPoint presentation with retry logic.
    
    PowerPoint COM can fail if a previous instance hasn't fully released.
    This retries with a delay to handle that race condition.
    """
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


def sanitize_text(text: str) -> str:
    """
    Remove control characters that break XML/docx.
    
    Keeps: tab (0x09), newline (0x0A), carriage return (0x0D)
    Removes: NULL bytes, other control chars (0x00-0x08, 0x0B-0x0C, 0x0E-0x1F)
    """
    if not text:
        return text
    
    delete_chars = ''.join(
        chr(i) for i in range(32) 
        if i not in (9, 10, 13)
    )
    delete_chars += chr(127)
    
    return text.translate(str.maketrans('', '', delete_chars))


def get_short_path(long_path: str) -> str:
    """Convert long path with spaces to Windows 8.3 short path format."""
    if not HAS_WIN32API:
        return long_path
    
    try:
        long_path = long_path.replace('/', '\\')
        short_path = win32api.GetShortPathName(long_path)
        return short_path
    except Exception:
        return long_path


def _get_slide_notes(pres, slide_num: int) -> str:
    """Get notes text from a slide."""
    try:
        slide = pres.Slides.Item(slide_num)
        notes_page = slide.NotesPage
        for shape in notes_page.Shapes:
            if shape.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                if shape.HasTextFrame:
                    return shape.TextFrame.TextRange.Text or ""
    except:
        pass
    return ""


def _set_slide_notes(pres, slide_num: int, text: str) -> bool:
    """Set notes text on a slide. Returns True on success."""
    try:
        slide = pres.Slides.Item(slide_num)
        notes_page = slide.NotesPage
        for shape in notes_page.Shapes:
            if shape.PlaceholderFormat.Type == 2:  # ppPlaceholderBody
                if shape.HasTextFrame:
                    shape.TextFrame.TextRange.Text = text
                    return True
    except:
        pass
    return False


def _get_slide_title(pres, slide_num: int) -> str:
    """Get title from a slide."""
    try:
        slide = pres.Slides.Item(slide_num)
        for shape in slide.Shapes:
            if shape.HasTextFrame:
                if shape.PlaceholderFormat.Type == 1:  # ppPlaceholderTitle
                    return shape.TextFrame.TextRange.Text.strip()
    except:
        pass
    return ""


# =============================================================================
# FIND FUNCTIONS
# =============================================================================

def find_in_notes(pptx_path: str, search_term: str,
                  case_sensitive: bool = False, use_regex: bool = False,
                  log_callback: Callable = None) -> List[Dict]:
    """
    Find all occurrences of search term in speaker notes.
    
    Args:
        pptx_path: Path to PowerPoint file
        search_term: Text to search for
        case_sensitive: Whether to match case (default: False)
        use_regex: Whether to treat search_term as regex (default: False)
        log_callback: Optional function for progress logging
    
    Returns:
        List of matches:
        [
            {
                "slide_number": 3,
                "slide_title": "Configuration",
                "match_count": 2,
                "matches": [
                    {"start": 45, "end": 54, "context": "...the Acme Corp system..."},
                    {"start": 120, "end": 129, "context": "...contact Acme Corp for..."}
                ],
                "notes_preview": "First 200 chars of notes..."
            },
            ...
        ]
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available")
    
    if not search_term:
        raise ValueError("Search term cannot be empty")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    # Build regex pattern
    if use_regex:
        try:
            flags = 0 if case_sensitive else re.IGNORECASE
            pattern = re.compile(search_term, flags)
        except re.error as e:
            raise ValueError(f"Invalid regex pattern: {e}")
    else:
        # Escape special regex chars for literal search
        escaped = re.escape(search_term)
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(escaped, flags)
    
    results = []
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=True)
        slide_count = pres.Slides.Count
        
        _log(f"Searching {slide_count} slides for: {search_term}")
        
        total_matches = 0
        
        for i in range(1, slide_count + 1):
            notes_text = sanitize_text(_get_slide_notes(pres, i))
            
            if not notes_text:
                continue
            
            # Find all matches
            matches = list(pattern.finditer(notes_text))
            
            if matches:
                title = sanitize_text(_get_slide_title(pres, i))
                
                match_details = []
                for m in matches:
                    # Build context snippet (50 chars before/after)
                    start = max(0, m.start() - 50)
                    end = min(len(notes_text), m.end() + 50)
                    context = notes_text[start:end]
                    if start > 0:
                        context = "..." + context
                    if end < len(notes_text):
                        context = context + "..."
                    
                    match_details.append({
                        "start": m.start(),
                        "end": m.end(),
                        "matched_text": m.group(),
                        "context": context
                    })
                
                results.append({
                    "slide_number": i,
                    "slide_title": title,
                    "match_count": len(matches),
                    "matches": match_details,
                    "notes_preview": notes_text[:200] + ("..." if len(notes_text) > 200 else "")
                })
                
                total_matches += len(matches)
                _log(f"  Slide {i}: {len(matches)} match(es)")
        
        _log(f"Found {total_matches} match(es) across {len(results)} slide(s)")
        
        return results
        
    except Exception as e:
        raise RuntimeError(f"Search failed: {e}")
    
    finally:
        # Always clean up COM objects
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


# =============================================================================
# REPLACE FUNCTIONS
# =============================================================================

def preview_replace(pptx_path: str, search_term: str, replace_term: str,
                    case_sensitive: bool = False, use_regex: bool = False,
                    log_callback: Callable = None) -> List[Dict]:
    """
    Preview what replacements would be made without applying them.
    
    Args:
        pptx_path: Path to PowerPoint file
        search_term: Text to search for
        replace_term: Text to replace with
        case_sensitive: Whether to match case (default: False)
        use_regex: Whether to treat search_term as regex (default: False)
        log_callback: Optional function for progress logging
    
    Returns:
        List of previews:
        [
            {
                "slide_number": 3,
                "slide_title": "Configuration",
                "match_count": 2,
                "original_notes": "...original text...",
                "preview_notes": "...text after replacement..."
            },
            ...
        ]
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available")
    
    if not search_term:
        raise ValueError("Search term cannot be empty")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    # Build regex pattern
    if use_regex:
        try:
            flags = 0 if case_sensitive else re.IGNORECASE
            pattern = re.compile(search_term, flags)
        except re.error as e:
            raise ValueError(f"Invalid regex pattern: {e}")
    else:
        escaped = re.escape(search_term)
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(escaped, flags)
    
    results = []
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=True)
        slide_count = pres.Slides.Count
        
        _log(f"Previewing replacement: '{search_term}' -> '{replace_term}'")
        
        total_replacements = 0
        
        for i in range(1, slide_count + 1):
            notes_text = sanitize_text(_get_slide_notes(pres, i))
            
            if not notes_text:
                continue
            
            # Check for matches
            matches = list(pattern.finditer(notes_text))
            
            if matches:
                title = sanitize_text(_get_slide_title(pres, i))
                
                # Perform replacement (preview only)
                new_text = pattern.sub(replace_term, notes_text)
                
                results.append({
                    "slide_number": i,
                    "slide_title": title,
                    "match_count": len(matches),
                    "original_notes": notes_text,
                    "preview_notes": new_text
                })
                
                total_replacements += len(matches)
                _log(f"  Slide {i}: {len(matches)} replacement(s)")
        
        _log(f"Preview: {total_replacements} replacement(s) across {len(results)} slide(s)")
        
        return results
        
    except Exception as e:
        raise RuntimeError(f"Preview failed: {e}")
    
    finally:
        # Always clean up COM objects
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


def replace_in_notes(pptx_path: str, search_term: str, replace_term: str,
                     case_sensitive: bool = False, use_regex: bool = False,
                     slides_to_apply: List[int] = None,
                     log_callback: Callable = None) -> Dict:
    """
    Perform find/replace across speaker notes.
    
    Args:
        pptx_path: Path to PowerPoint file
        search_term: Text to search for
        replace_term: Text to replace with
        case_sensitive: Whether to match case (default: False)
        use_regex: Whether to treat search_term as regex (default: False)
        slides_to_apply: Optional list of slide numbers to apply (default: all)
        log_callback: Optional function for progress logging
    
    Returns:
        Result dict:
        {
            "slides_modified": [3, 7, 12],
            "total_replacements": 5,
            "errors": []
        }
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available")
    
    if not search_term:
        raise ValueError("Search term cannot be empty")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    # Build regex pattern
    if use_regex:
        try:
            flags = 0 if case_sensitive else re.IGNORECASE
            pattern = re.compile(search_term, flags)
        except re.error as e:
            raise ValueError(f"Invalid regex pattern: {e}")
    else:
        escaped = re.escape(search_term)
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(escaped, flags)
    
    result = {
        "slides_modified": [],
        "total_replacements": 0,
        "errors": []
    }
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        # Open for editing (not read-only) with retry
        pres = open_presentation_with_retry(app, pptx_path, read_only=False)
        slide_count = pres.Slides.Count
        
        _log(f"Replacing: '{search_term}' -> '{replace_term}'")
        
        for i in range(1, slide_count + 1):
            # Skip if not in filter list
            if slides_to_apply is not None and i not in slides_to_apply:
                continue
            
            notes_text = sanitize_text(_get_slide_notes(pres, i))
            
            if not notes_text:
                continue
            
            # Check for matches
            matches = list(pattern.finditer(notes_text))
            
            if matches:
                # Perform replacement
                new_text = pattern.sub(replace_term, notes_text)
                
                # Write back
                if _set_slide_notes(pres, i, new_text):
                    result["slides_modified"].append(i)
                    result["total_replacements"] += len(matches)
                    _log(f"  Slide {i}: {len(matches)} replacement(s)")
                else:
                    result["errors"].append(f"Slide {i}: Failed to write notes")
                    _log(f"  Slide {i}: Error writing notes")
        
        # Save changes
        pres.Save()
        
        _log(f"Completed: {result['total_replacements']} replacement(s) across {len(result['slides_modified'])} slide(s)")
        
        return result
        
    except Exception as e:
        raise RuntimeError(f"Replace failed: {e}")
    
    finally:
        # Always clean up COM objects
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


# =============================================================================
# BATCH REPLACE
# =============================================================================

def batch_replace(pptx_path: str, replacements: List[Tuple[str, str]],
                  case_sensitive: bool = False,
                  log_callback: Callable = None) -> Dict:
    """
    Perform multiple find/replace operations in a single pass.
    
    Args:
        pptx_path: Path to PowerPoint file
        replacements: List of (search, replace) tuples
        case_sensitive: Whether to match case (default: False)
        log_callback: Optional function for progress logging
    
    Returns:
        Result dict:
        {
            "slides_modified": [3, 7, 12],
            "replacements_by_term": {"old1": 3, "old2": 5},
            "total_replacements": 8,
            "errors": []
        }
    
    Example:
        replacements = [
            ("Acme Corp", "Acme Industries"),
            ("v1.0", "v2.0"),
            ("2023", "2024")
        ]
        result = batch_replace("Training.pptx", replacements)
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available")
    
    if not replacements:
        raise ValueError("Replacements list cannot be empty")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    # Build patterns for all replacements
    patterns = []
    for search_term, replace_term in replacements:
        if not search_term:
            continue
        escaped = re.escape(search_term)
        flags = 0 if case_sensitive else re.IGNORECASE
        pattern = re.compile(escaped, flags)
        patterns.append((pattern, search_term, replace_term))
    
    result = {
        "slides_modified": [],
        "replacements_by_term": {search: 0 for search, _ in replacements},
        "total_replacements": 0,
        "errors": []
    }
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=False)
        slide_count = pres.Slides.Count
        
        _log(f"Batch replacing {len(patterns)} term(s) across {slide_count} slides...")
        
        for i in range(1, slide_count + 1):
            notes_text = sanitize_text(_get_slide_notes(pres, i))
            
            if not notes_text:
                continue
            
            modified = False
            current_text = notes_text
            
            # Apply each replacement
            for pattern, search_term, replace_term in patterns:
                matches = list(pattern.finditer(current_text))
                if matches:
                    current_text = pattern.sub(replace_term, current_text)
                    result["replacements_by_term"][search_term] += len(matches)
                    result["total_replacements"] += len(matches)
                    modified = True
            
            # Write back if any changes
            if modified:
                if _set_slide_notes(pres, i, current_text):
                    result["slides_modified"].append(i)
                    _log(f"  Slide {i}: Modified")
                else:
                    result["errors"].append(f"Slide {i}: Failed to write notes")
        
        pres.Save()
        
        _log(f"Completed: {result['total_replacements']} total replacement(s)")
        
        return result
        
    except Exception as e:
        raise RuntimeError(f"Batch replace failed: {e}")
    
    finally:
        # Always clean up COM objects
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


# =============================================================================
# STATS
# =============================================================================

def get_notes_stats(pptx_path: str, log_callback: Callable = None) -> Dict:
    """
    Get statistics about speaker notes in a deck.
    
    Args:
        pptx_path: Path to PowerPoint file
        log_callback: Optional function for progress logging
    
    Returns:
        Stats dict:
        {
            "total_slides": 49,
            "slides_with_notes": 46,
            "slides_without_notes": 3,
            "total_characters": 15234,
            "total_words": 2847,
            "avg_words_per_slide": 62
        }
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_COM:
        raise RuntimeError("Windows COM API not available")
    
    pptx_path = get_short_path(str(Path(pptx_path).resolve()))
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    stats = {
        "total_slides": 0,
        "slides_with_notes": 0,
        "slides_without_notes": 0,
        "total_characters": 0,
        "total_words": 0,
        "avg_words_per_slide": 0
    }
    
    app = None
    pres = None
    
    try:
        gencache.EnsureDispatch("PowerPoint.Application")
        app = Dispatch("PowerPoint.Application")
        app.Visible = True
        
        pres = open_presentation_with_retry(app, pptx_path, read_only=True)
        slide_count = pres.Slides.Count
        
        stats["total_slides"] = slide_count
        
        _log(f"Analyzing {slide_count} slides...")
        
        for i in range(1, slide_count + 1):
            notes_text = sanitize_text(_get_slide_notes(pres, i))
            
            if notes_text and notes_text.strip():
                stats["slides_with_notes"] += 1
                stats["total_characters"] += len(notes_text)
                # Simple word count (split on whitespace)
                words = len(notes_text.split())
                stats["total_words"] += words
            else:
                stats["slides_without_notes"] += 1
        
        # Calculate average
        if stats["slides_with_notes"] > 0:
            stats["avg_words_per_slide"] = round(
                stats["total_words"] / stats["slides_with_notes"]
            )
        
        _log(f"Stats: {stats['slides_with_notes']} slides with notes, {stats['total_words']} words total")
        
        return stats
        
    except Exception as e:
        raise RuntimeError(f"Stats failed: {e}")
    
    finally:
        # Always clean up COM objects
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


# =============================================================================
# CLI for standalone testing
# =============================================================================

def _usage():
    print("Usage: python voxreplace.py <command> <args>")
    print()
    print("Commands:")
    print("  find <deck.pptx> <search_term>              Find matches")
    print("  preview <deck.pptx> <search> <replace>      Preview replacements")
    print("  replace <deck.pptx> <search> <replace>      Apply replacements")
    print("  stats <deck.pptx>                           Show notes statistics")
    print()
    print("Options:")
    print("  -c, --case-sensitive    Match case exactly")
    print("  -r, --regex             Treat search as regex pattern")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 2:
        _usage()
        sys.exit(64)
    
    command = sys.argv[1].lower()
    
    # Parse options
    case_sensitive = '-c' in sys.argv or '--case-sensitive' in sys.argv
    use_regex = '-r' in sys.argv or '--regex' in sys.argv
    
    # Remove options from argv for positional args
    args = [a for a in sys.argv[2:] if not a.startswith('-')]
    
    try:
        if command == "find":
            if len(args) < 2:
                print("Usage: python voxreplace.py find <deck.pptx> <search_term>")
                sys.exit(64)
            
            results = find_in_notes(args[0], args[1], case_sensitive, use_regex)
            
            if not results:
                print("\nNo matches found.")
            else:
                print(f"\nFound matches in {len(results)} slide(s):")
                for r in results:
                    print(f"\n  Slide {r['slide_number']}: {r['slide_title'] or '(no title)'}")
                    print(f"    {r['match_count']} match(es)")
                    for m in r['matches'][:3]:  # Show first 3
                        print(f"      \"{m['context']}\"")
                    if len(r['matches']) > 3:
                        print(f"      ... and {len(r['matches']) - 3} more")
                        
        elif command == "preview":
            if len(args) < 3:
                print("Usage: python voxreplace.py preview <deck.pptx> <search> <replace>")
                sys.exit(64)
            
            results = preview_replace(args[0], args[1], args[2], case_sensitive, use_regex)
            
            if not results:
                print("\nNo matches found. Nothing to replace.")
            else:
                print(f"\nWould modify {len(results)} slide(s):")
                for r in results:
                    print(f"  Slide {r['slide_number']}: {r['match_count']} replacement(s)")
                    
        elif command == "replace":
            if len(args) < 3:
                print("Usage: python voxreplace.py replace <deck.pptx> <search> <replace>")
                sys.exit(64)
            
            result = replace_in_notes(args[0], args[1], args[2], case_sensitive, use_regex)
            
            print(f"\nReplaced {result['total_replacements']} occurrence(s)")
            print(f"Modified slides: {result['slides_modified']}")
            if result['errors']:
                print(f"Errors: {result['errors']}")
                
        elif command == "stats":
            if len(args) < 1:
                print("Usage: python voxreplace.py stats <deck.pptx>")
                sys.exit(64)
            
            stats = get_notes_stats(args[0])
            
            print(f"\nDeck Statistics:")
            print(f"  Total slides: {stats['total_slides']}")
            print(f"  Slides with notes: {stats['slides_with_notes']}")
            print(f"  Slides without notes: {stats['slides_without_notes']}")
            print(f"  Total words: {stats['total_words']}")
            print(f"  Total characters: {stats['total_characters']}")
            print(f"  Avg words/slide: {stats['avg_words_per_slide']}")
                
        else:
            print(f"Unknown command: {command}")
            _usage()
            sys.exit(64)
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)
