"""
voxmedia.py
Media extraction, stripping, and import for VoxPrep.

Handles audio/video embedded in PowerPoint decks:
- Strip all audio from slides
- Export media with slide-based naming (slide01.wav, slide02.mp4, etc.)
- Import audio back using voxattach

Example:
    # Remove all audio
    strip_all_audio("Training.pptx")
    
    # Export media to folder
    export_media("Training.pptx", "media_folder")
    # Creates: media_folder/slide01.wav, slide03.mp4, etc.
    
    # Import cleaned audio back
    import_audio("Training.pptx", "media_folder")
"""

import os
import re
import shutil
import time
import zipfile
from pathlib import Path
from typing import List, Dict, Optional, Callable
from xml.etree import ElementTree as ET

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

# Try to import voxattach for audio import functionality
try:
    import voxattach
    HAS_VOXATTACH = True
except ImportError:
    HAS_VOXATTACH = False

LOG_PREFIX = "[voxmedia]"

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
# STRIP ALL AUDIO
# ============================================================

def strip_all_audio(pptx_path: str, log_callback: Optional[Callable] = None) -> Dict:
    """
    Remove ALL audio shapes from all slides in a PowerPoint deck.
    
    Args:
        pptx_path: Path to the PowerPoint file
        log_callback: Optional function for progress logging
    
    Returns:
        Dict with 'success', 'slides_modified', 'audio_removed' counts
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
        
        _log(f"Scanning {slide_count} slides for audio...")
        
        total_removed = 0
        slides_modified = []
        
        for i in range(1, slide_count + 1):
            slide = pres.Slides.Item(i)
            removed_this_slide = 0
            
            # Iterate backwards to safely delete
            shape_count = slide.Shapes.Count
            for j in range(shape_count, 0, -1):
                try:
                    shape = slide.Shapes.Item(j)
                    # Type 16 = msoMedia, MediaType 2 = audio
                    if int(shape.Type) == 16 and hasattr(shape, "MediaType"):
                        if int(shape.MediaType) == 2:  # ppMediaTypeSound
                            shape.Delete()
                            removed_this_slide += 1
                except Exception:
                    pass
            
            if removed_this_slide > 0:
                slides_modified.append(i)
                total_removed += removed_this_slide
                _log(f"  Slide {i}: removed {removed_this_slide} audio shape(s)")
        
        if total_removed > 0:
            pres.Save()
            _log(f"Saved. Removed {total_removed} audio shape(s) from {len(slides_modified)} slide(s).")
        else:
            _log("No audio found in deck.")
        
        return {
            "success": True,
            "audio_removed": total_removed,
            "slides_modified": slides_modified
        }
        
    except Exception as e:
        raise RuntimeError(f"Strip audio failed: {e}")
    
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
# EXPORT MEDIA
# ============================================================

def _parse_pptx_media_relationships(pptx_path: str) -> Dict[int, List[Dict]]:
    """
    Parse PPTX (ZIP) to map slides to their embedded media files.
    
    Returns:
        Dict mapping slide_number -> list of {filename, media_type, internal_path}
    """
    slide_media = {}
    
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        # Find all slide relationship files
        slide_rels = [n for n in zf.namelist() if re.match(r'ppt/slides/_rels/slide\d+\.xml\.rels', n)]
        
        for rel_path in slide_rels:
            # Extract slide number from path
            match = re.search(r'slide(\d+)\.xml\.rels', rel_path)
            if not match:
                continue
            slide_num = int(match.group(1))
            
            # Parse the relationships XML
            rel_content = zf.read(rel_path).decode('utf-8')
            root = ET.fromstring(rel_content)
            
            media_files = []
            ns = {'r': 'http://schemas.openxmlformats.org/package/2006/relationships'}
            
            for rel in root.findall('.//r:Relationship', ns):
                rel_type = rel.get('Type', '')
                target = rel.get('Target', '')
                
                # Look for audio or video relationships
                if 'audio' in rel_type.lower() or 'video' in rel_type.lower():
                    # Target is relative path like ../media/media1.m4a
                    if target.startswith('../media/'):
                        media_filename = target.replace('../media/', '')
                        internal_path = f'ppt/media/{media_filename}'
                        
                        # Determine media type from extension
                        ext = os.path.splitext(media_filename)[1].lower()
                        if ext in ['.m4a', '.mp3', '.wav', '.wma', '.aiff']:
                            media_type = 'audio'
                        elif ext in ['.mp4', '.m4v', '.mov', '.wmv', '.avi']:
                            media_type = 'video'
                        else:
                            media_type = 'unknown'
                        
                        media_files.append({
                            'filename': media_filename,
                            'media_type': media_type,
                            'internal_path': internal_path,
                            'extension': ext
                        })
            
            if media_files:
                slide_media[slide_num] = media_files
    
    return slide_media


def export_media(pptx_path: str, output_folder: str, log_callback: Optional[Callable] = None) -> Dict:
    """
    Export all embedded audio/video from a PowerPoint deck.
    
    Files are named by slide number: slide01.wav, slide02.mp4, etc.
    Audio is extracted in its native format (usually m4a) - user converts as needed.
    
    Args:
        pptx_path: Path to the PowerPoint file
        output_folder: Folder to save extracted media
        log_callback: Optional function for progress logging
    
    Returns:
        Dict with 'success', 'files_exported', 'manifest' (list of exported files)
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    pptx_path = str(Path(pptx_path).resolve())
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    os.makedirs(output_folder, exist_ok=True)
    
    _log("Analyzing deck for embedded media...")
    
    try:
        slide_media = _parse_pptx_media_relationships(pptx_path)
    except Exception as e:
        raise RuntimeError(f"Failed to parse PPTX structure: {e}")
    
    if not slide_media:
        _log("No embedded media found in deck.")
        return {
            "success": True,
            "files_exported": 0,
            "manifest": []
        }
    
    _log(f"Found media on {len(slide_media)} slide(s)")
    
    manifest = []
    files_exported = 0
    
    with zipfile.ZipFile(pptx_path, 'r') as zf:
        for slide_num in sorted(slide_media.keys()):
            media_list = slide_media[slide_num]
            
            for idx, media_info in enumerate(media_list):
                internal_path = media_info['internal_path']
                media_type = media_info['media_type']
                original_ext = media_info['extension']
                
                # Build output filename: slide01.m4a, slide01.mp4, etc.
                # If multiple media on same slide, add suffix: slide01_2.m4a
                if len(media_list) == 1:
                    out_filename = f"slide{slide_num:02d}{original_ext}"
                else:
                    out_filename = f"slide{slide_num:02d}_{idx + 1}{original_ext}"
                
                out_path = os.path.join(output_folder, out_filename)
                
                try:
                    # Extract from ZIP
                    media_data = zf.read(internal_path)
                    with open(out_path, 'wb') as f:
                        f.write(media_data)
                    
                    _log(f"  Slide {slide_num}: {out_filename} ({media_type})")
                    
                    manifest.append({
                        'slide': slide_num,
                        'filename': out_filename,
                        'media_type': media_type,
                        'path': out_path
                    })
                    files_exported += 1
                    
                except KeyError:
                    _log(f"  Slide {slide_num}: media file not found in archive")
                except Exception as e:
                    _log(f"  Slide {slide_num}: export failed - {e}")
    
    _log(f"Exported {files_exported} file(s) to {output_folder}")
    
    return {
        "success": True,
        "files_exported": files_exported,
        "manifest": manifest
    }


# ============================================================
# IMPORT AUDIO
# ============================================================

def import_audio(pptx_path: str, media_folder: str, log_callback: Optional[Callable] = None) -> Dict:
    """
    Import audio files back into PowerPoint slides.
    
    Looks for files named slideXX.wav in the media folder and attaches
    them to the corresponding slides using voxattach.
    
    Args:
        pptx_path: Path to the PowerPoint file
        media_folder: Folder containing slideXX.wav files
        log_callback: Optional function for progress logging
    
    Returns:
        Dict with 'success', 'files_imported', 'slides_updated'
    """
    def _log(msg):
        if log_callback:
            log_callback(msg)
        else:
            log(msg)
    
    if not HAS_VOXATTACH:
        raise RuntimeError("voxattach module not available. Cannot import audio.")
    
    pptx_path = str(Path(pptx_path).resolve())
    
    if not os.path.isfile(pptx_path):
        raise FileNotFoundError(f"PowerPoint file not found: {pptx_path}")
    
    if not os.path.isdir(media_folder):
        raise FileNotFoundError(f"Media folder not found: {media_folder}")
    
    # Find all slideXX.wav files
    audio_files = []
    pattern = re.compile(r'^slide(\d+)\.wav$', re.IGNORECASE)
    
    for filename in os.listdir(media_folder):
        match = pattern.match(filename)
        if match:
            slide_num = int(match.group(1))
            audio_files.append({
                'slide': slide_num,
                'filename': filename,
                'path': os.path.join(media_folder, filename)
            })
    
    if not audio_files:
        _log("No slideXX.wav files found in media folder.")
        return {
            "success": True,
            "files_imported": 0,
            "slides_updated": []
        }
    
    # Sort by slide number
    audio_files.sort(key=lambda x: x['slide'])
    
    _log(f"Found {len(audio_files)} audio file(s) to import")
    
    # Reset voxattach for new run
    voxattach.reset_for_new_run()
    
    files_imported = 0
    slides_updated = []
    
    for audio_info in audio_files:
        slide_num = audio_info['slide']
        audio_path = audio_info['path']
        
        _log(f"  Slide {slide_num}: {audio_info['filename']}")
        
        try:
            # voxattach expects src and dst audio paths
            # For import, src and dst are the same (audio already processed)
            result = voxattach.attach_or_skip(
                pptx_path,
                slide_num,
                audio_path,
                audio_path  # dst same as src - no processing needed
            )
            
            if result.get('attached'):
                files_imported += 1
                slides_updated.append(slide_num)
            elif result.get('reason') == 'open_process_only':
                _log(f"    Skipped (deck open in PowerPoint)")
            else:
                _log(f"    Attachment skipped: {result.get('reason', 'unknown')}")
                
        except Exception as e:
            _log(f"    Failed: {e}")
    
    _log(f"Imported {files_imported} audio file(s) to {len(slides_updated)} slide(s)")
    
    return {
        "success": True,
        "files_imported": files_imported,
        "slides_updated": slides_updated
    }


# ============================================================
# CLI
# ============================================================

def _usage():
    print("Usage:")
    print("  python voxmedia.py strip <deck.pptx>")
    print("  python voxmedia.py export <deck.pptx> <output_folder>")
    print("  python voxmedia.py import <deck.pptx> <media_folder>")


if __name__ == "__main__":
    import sys
    
    if len(sys.argv) < 3:
        _usage()
        sys.exit(64)
    
    command = sys.argv[1].lower()
    deck_path = sys.argv[2]
    
    try:
        if command == "strip":
            result = strip_all_audio(deck_path)
            print(f"Removed {result['audio_removed']} audio shape(s)")
            
        elif command == "export":
            if len(sys.argv) < 4:
                print("Error: output_folder required for export")
                sys.exit(64)
            output_folder = sys.argv[3]
            result = export_media(deck_path, output_folder)
            print(f"Exported {result['files_exported']} file(s)")
            
        elif command == "import":
            if len(sys.argv) < 4:
                print("Error: media_folder required for import")
                sys.exit(64)
            media_folder = sys.argv[3]
            result = import_audio(deck_path, media_folder)
            print(f"Imported {result['files_imported']} file(s)")
            
        else:
            print(f"Unknown command: {command}")
            _usage()
            sys.exit(64)
            
    except Exception as e:
        print(f"Error: {e}")
        import traceback
        traceback.print_exc()
        sys.exit(2)
