import ctypes
import os


def configure_windows_dpi_awareness():
    if os.name != "nt":
        return

    try:
        # Prefer per-monitor v2 so UIAutomation rectangles and mouse events use
        # the same physical coordinate space on mixed-DPI displays.
        if ctypes.windll.user32.SetProcessDpiAwarenessContext(ctypes.c_void_p(-4)):
            return
    except Exception:
        pass

    try:
        ctypes.windll.shcore.SetProcessDpiAwareness(2)
        return
    except Exception:
        pass

    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass
