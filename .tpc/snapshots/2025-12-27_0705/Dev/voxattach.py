"""
voxattach 1.1
Lightweight helper for attaching audio to specific slides in a PPTX using the
PowerPoint COM API. Designed for use inside Voxsmith v2.x.

Changes in 1.1:
- Fixed logic that skipped attachment after slide 1 by incorrectly short-circuiting
  when the deck was open. Run-mode is now decided once per deck and respected.
- General cleanup and small robustness tweaks.
"""
from __future__ import annotations

from pathlib import Path
import sys

__version__ = "1.1"

try:
    import win32com.client as win32
except Exception as e:
    win32 = None
    _import_error = e
else:
    _import_error = None

LOG_PREFIX = "[voxattach]"

def log(msg: str):
    print(f"{LOG_PREFIX} {msg}", flush=True)


# ---------- Audio staging ----------

def process_audio(src_audio: str, dst_audio: str):
    p_src = Path(src_audio)
    p_dst = Path(dst_audio)
    if not p_src.exists():
        raise FileNotFoundError(f"source audio missing: {p_src}")
    p_dst.parent.mkdir(parents=True, exist_ok=True)
    if p_src.resolve() != p_dst.resolve():
        p_dst.write_bytes(p_src.read_bytes())
    log(f"processed audio -> {p_dst}")


# ---------- COM helpers ----------

def _require_windows_com():
    if win32 is None:
        raise RuntimeError(f"pywin32/COM unavailable: {_import_error!r}")


def _ppt_running_app():
    try:
        return win32.GetActiveObject("PowerPoint.Application")
    except Exception:
        return None


def _ppt_new_app():
    app = win32.Dispatch("PowerPoint.Application")
    app.Visible = True  # keep visible to avoid weird focus issues
    return app


def is_deck_open(pptx_path: str) -> bool:
    """Return True if the exact presentation is already open (editable)."""
    app = _ppt_running_app()
    if not app:
        return False
    full = str(Path(pptx_path).resolve())
    for pres in app.Presentations:
        try:
            if str(pres.FullName).lower() == full.lower():
                return True
        except Exception:
            pass
    return False


# ---------- Single-session + run-mode ----------

_SESSION = {"app": None, "pres": None, "path": None, "opened_by_us": False}

# Run-mode decided once per deck path for the entire run:
#   "attach"        -> deck was closed at start
#   "process_only"  -> deck was open at start
_RUN_MODE = {"mode": None, "path": None}


def reset_for_new_run():
    """
    Reset the run-mode so the next attach_or_skip call checks deck status fresh.
    Call this at the START of each Voxsmith narration generation run.
    """
    global _RUN_MODE
    _RUN_MODE = {"mode": None, "path": None}
    log("RUN MODE: Reset for new run. Will check deck status on first slide.")


def _ensure_session(deck_path: str):
    """Open or reuse a single Presentation for the given deck_path. Leave it open."""
    _require_windows_com()
    global _SESSION
    full = str(Path(deck_path).resolve())

    # Reuse if valid and for same deck
    if _SESSION["pres"] is not None and _SESSION["path"] and _SESSION["path"].lower() == full.lower():
        try:
            _ = _SESSION["pres"].FullName  # touch to ensure COM object is alive
            return _SESSION["app"], _SESSION["pres"], _SESSION["opened_by_us"]
        except Exception:
            # stale; drop and reopen
            _SESSION.update({"app": None, "pres": None, "path": None, "opened_by_us": False})

    app = _ppt_running_app()
    opened_by_us = False
    if app is None:
        app = _ppt_new_app()
        opened_by_us = True

    # Find open pres or open it
    pres = None
    for p in app.Presentations:
        try:
            if str(p.FullName).lower() == full.lower():
                pres = p
                opened_by_us = False  # user had it open
                break
        except Exception:
            pass
    if pres is None:
        pres = app.Presentations.Open(full, WithWindow=True)
        opened_by_us = True

    _SESSION = {"app": app, "pres": pres, "path": full, "opened_by_us": opened_by_us}
    return app, pres, opened_by_us


# ---------- Slide-level helpers ----------

def _delete_existing_vox_audio(slide):
    """Only remove our prior audio (tagged AlternativeText == 'VOX_VO')."""
    try:
        count = int(slide.Shapes.Count)
        for i in range(count, 0, -1):
            s = slide.Shapes.Item(i)
            try:
                if int(s.Type) == 16 and hasattr(s, "MediaType") and int(s.MediaType) == 2:
                    alt = str(getattr(s, "AlternativeText", "") or "")
                    if alt.strip() == "VOX_VO":
                        s.Delete()
            except Exception:
                pass
    except Exception:
        pass


def _configure_play_settings(shape, *, hide=True):
    """Configure media playback flags on the shape without touching timelines."""
    try:
        ps = shape.AnimationSettings.PlaySettings
        ps.PlayOnEntry = False            # we'll rely on explicit effect
        ps.HideWhileNotPlaying = bool(hide)
        ps.LoopUntilStopped = False
        ps.RewindMovieWhenDone = False
        ps.StopPreviousSound = True
        shape.AnimationSettings.EntryEffect = 0
    except Exception:
        pass


def _append_media_play_after_previous(slide, shape):
    """Append a 'Play (Media)' effect to MainSequence with After Previous trigger."""
    try:
        seq = slide.TimeLine.MainSequence
        msoAnimEffectMediaPlay = 83      # Media Play
        msoAnimTriggerAfterPrevious = 3  # After Previous
        eff = seq.AddEffect(shape, msoAnimEffectMediaPlay)
        eff.Timing.TriggerType = msoAnimTriggerAfterPrevious
        eff.Timing.TriggerDelayTime = 0.0
    except Exception:
        pass


def _attach_on_open_presentation(pres, slide_index_1based: int, audio_path: str, *, left=20, top=20, width=32, height=32):
    """Attach using an already-open Presentation; save but DO NOT close. Leave deck open."""
    slide = pres.Slides(slide_index_1based)

    # Clean up only our shapes
    _delete_existing_vox_audio(slide)

    # Insert the media icon
    shape = slide.Shapes.AddMediaObject2(str(Path(audio_path).resolve()), False, True, float(left), float(top))
    try:
        shape.Width = float(width)
        shape.Height = float(height)
    except Exception:
        pass

    # Position off the slide to the right; bottom aligned with small margin
    try:
        W = pres.PageSetup.SlideWidth
        H = pres.PageSetup.SlideHeight
        shape.Left = W + 5
        shape.Top  = H - shape.Height - 5
    except Exception:
        pass

    # Tag as ours
    try:
        shape.AlternativeText = "VOX_VO"
    except Exception:
        pass

    # Configure visibility and playback
    _configure_play_settings(shape, hide=True)
    _append_media_play_after_previous(slide, shape)

    # Save after each slide
    try:
        pres.Save()
    except Exception:
        pass


# ---------- Public API ----------

def attach_or_skip(pptx_path: str, slide_index_1based: int, src_audio: str, out_audio: str, *, left=20, top=20, width=32, height=32):
    """
    Run-mode behavior:
      - Decide mode once per deck path at first call:
          * 'process_only' if deck is open at start
          * 'attach' if deck is closed at start
      - In 'process_only': copy audio only, never attach
      - In 'attach': open/reuse session, attach every slide, leave deck open
    """
    # Always stage the audio
    process_audio(src_audio, out_audio)

    full = str(Path(pptx_path).resolve())

    # Decide run-mode at FIRST slide only
    if _RUN_MODE["mode"] is None or (_RUN_MODE["path"] and _RUN_MODE["path"].lower() != full.lower()):
        current_deck_open = is_deck_open(pptx_path)
        if current_deck_open:
            _RUN_MODE.update({"mode": "process_only", "path": full})
            log("RUN MODE: deck open at start -> Process-only for entire run.")
        else:
            _RUN_MODE.update({"mode": "attach", "path": full})
            log("RUN MODE: deck closed at start -> Attach mode for entire run.")

    # If COM unavailable, degrade to process-only behavior
    if win32 is None:
        log("deck attach skipped: COM unavailable (pywin32 not installed)")
        return {"processed": True, "attached": False, "reason": "no_com", "out_audio": str(Path(out_audio).resolve())}

    if _RUN_MODE["mode"] == "process_only":
        # Hands-off: do NOT attach on any slide this run
        return {"processed": True, "attached": False, "reason": "open_process_only", "out_audio": str(Path(out_audio).resolve())}

    # Attach mode: open/reuse single session and attach to ALL slides; leave deck open
    try:
        app, pres, opened_by_us = _ensure_session(pptx_path)
        _attach_on_open_presentation(pres, slide_index_1based, out_audio, left=left, top=top, width=width, height=height)
        return {"processed": True, "attached": True, "reason": None, "out_audio": str(Path(out_audio).resolve())}
    except Exception as e:
        log(f"attach failed: {e}. processed audio already at {out_audio}")
        return {"processed": True, "attached": False, "reason": "exception", "error": str(e), "out_audio": str(Path(out_audio).resolve())}


# ---------- CLI (single-slide testing) ----------

def _usage():
    print("usage: python voxattach.py <deck.pptx> <slide_index> <src_audio> <out_audio> [left] [top] [width] [height]")

if __name__ == "__main__":
    if len(sys.argv) < 5:
        _usage()
        sys.exit(64)
    deck = sys.argv[1]
    try:
        slide_idx = int(sys.argv[2])
    except ValueError:
        print("slide_index must be an integer (1-based).")
        sys.exit(64)
    src = sys.argv[3]
    out = sys.argv[4]
    coords = [20, 20, 32, 32]
    if len(sys.argv) >= 9:
        try:
            coords = [float(sys.argv[5]), float(sys.argv[6]), float(sys.argv[7]), float(sys.argv[8])]
        except Exception:
            pass

    try:
        result = attach_or_skip(deck, slide_idx, src, out,
                                left=coords[0], top=coords[1], width=coords[2], height=coords[3])
        sys.exit(0 if (result.get("attached") or result.get("processed")) else 2)
    except Exception as e:
        log(f"fatal: {e}")
        sys.exit(2)