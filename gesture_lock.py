
import cv2
import ctypes
import ctypes.wintypes
import threading
import time
import os
from PIL import Image, ImageDraw
import pystray
from pystray import MenuItem as item

from mediapipe.tasks import python as mp_python
from mediapipe.tasks.python.vision import GestureRecognizer, GestureRecognizerOptions, RunningMode
import urllib.request

GESTURE_HOLD_SECONDS = 1.0
CAMERA_INDEX         = 0
CHECK_INTERVAL_MS    = 30
MODEL_PATH           = "gesture_recognizer.task"
MODEL_URL            = "https://storage.googleapis.com/mediapipe-models/gesture_recognizer/gesture_recognizer/float16/1/gesture_recognizer.task"

# Global state
is_running        = True
screen_locked     = False
lock_armed        = False
fist_start        = None
status_text       = "Watching..."
tray_icon_ref     = None

# Shared camera reference so lock_windows() can release it instantly
_cap_lock         = threading.Lock()
_cap              = None          # The cv2.VideoCapture object


def get_cap():
    with _cap_lock:
        return _cap

def set_cap(new_cap):
    global _cap
    with _cap_lock:
        _cap = new_cap

def release_cap():
    """Release camera immediately — call before locking so light turns off."""
    global _cap
    with _cap_lock:
        if _cap is not None:
            _cap.release()
            _cap = None


# ── Lock with camera release first ────────

def lock_windows():
    """Release camera first, then lock — ensures webcam light is OFF on lock screen."""
    release_cap()
    time.sleep(0.15)   # Small pause so the OS registers the camera is free
    ctypes.windll.user32.LockWorkStation()


# ── Session watcher (lock / unlock events) ─

WTS_SESSION_LOCK   = 0x7
WTS_SESSION_UNLOCK = 0x8

def start_session_watcher():
    global screen_locked, status_text, tray_icon_ref

    WM_WTSSESSION_CHANGE = 0x02B1
    WM_DESTROY           = 0x0002
    user32               = ctypes.windll.user32
    kernel32             = ctypes.windll.kernel32

    WNDPROCTYPE = ctypes.WINFUNCTYPE(
        ctypes.c_long,
        ctypes.wintypes.HWND,
        ctypes.c_uint,
        ctypes.wintypes.WPARAM,
        ctypes.wintypes.LPARAM,
    )

    def wnd_proc(hwnd, msg, wparam, lparam):
        global screen_locked, status_text, tray_icon_ref
        if msg == WM_WTSSESSION_CHANGE:
            if wparam == WTS_SESSION_LOCK:
                screen_locked = True
                status_text   = "Screen locked — paused"
                release_cap()   # Belt-and-suspenders: ensure camera is free
                if tray_icon_ref:
                    try:
                        tray_icon_ref.icon  = create_tray_image("orange")
                        tray_icon_ref.title = "GestureLock — Screen Locked"
                    except Exception:
                        pass
            elif wparam == WTS_SESSION_UNLOCK:
                screen_locked = False
                status_text   = "Watching..."
                if tray_icon_ref:
                    try:
                        color = "green" if is_running else "red"
                        tray_icon_ref.icon  = create_tray_image(color)
                        tray_icon_ref.title = f"GestureLock — {'Active' if is_running else 'Paused'}"
                    except Exception:
                        pass
        elif msg == WM_DESTROY:
            user32.PostQuitMessage(0)
        return user32.DefWindowProcW(hwnd, msg, wparam, lparam)

    wnd_proc_cb = WNDPROCTYPE(wnd_proc)

    class WNDCLASS(ctypes.Structure):
        _fields_ = [
            ("style",         ctypes.c_uint),
            ("lpfnWndProc",   WNDPROCTYPE),
            ("cbClsExtra",    ctypes.c_int),
            ("cbWndExtra",    ctypes.c_int),
            ("hInstance",     ctypes.wintypes.HANDLE),
            ("hIcon",         ctypes.wintypes.HANDLE),
            ("hCursor",       ctypes.wintypes.HANDLE),
            ("hbrBackground", ctypes.wintypes.HANDLE),
            ("lpszMenuName",  ctypes.c_wchar_p),
            ("lpszClassName", ctypes.c_wchar_p),
        ]

    hinstance  = kernel32.GetModuleHandleW(None)
    class_name = "GestureLockSessionWatcher"
    wc               = WNDCLASS()
    wc.lpfnWndProc   = wnd_proc_cb
    wc.hInstance     = hinstance
    wc.lpszClassName = class_name
    user32.RegisterClassW(ctypes.byref(wc))

    HWND_MESSAGE = ctypes.wintypes.HWND(-3)
    hwnd = user32.CreateWindowExW(
        0, class_name, "GestureLockWatcher", 0,
        0, 0, 0, 0, HWND_MESSAGE, None, hinstance, None,
    )
    ctypes.windll.wtsapi32.WTSRegisterSessionNotification(hwnd, 0)

    class MSG(ctypes.Structure):
        _fields_ = [
            ("hwnd",    ctypes.wintypes.HWND),
            ("message", ctypes.c_uint),
            ("wParam",  ctypes.wintypes.WPARAM),
            ("lParam",  ctypes.wintypes.LPARAM),
            ("time",    ctypes.c_uint),
            ("pt",      ctypes.wintypes.POINT),
        ]

    msg = MSG()
    while user32.GetMessageW(ctypes.byref(msg), None, 0, 0) != 0:
        user32.TranslateMessage(ctypes.byref(msg))
        user32.DispatchMessageW(ctypes.byref(msg))

    ctypes.windll.wtsapi32.WTSUnRegisterSessionNotification(hwnd)


# ── Model ─────────────────────────────────

def download_model():
    if not os.path.exists(MODEL_PATH):
        print("Downloading gesture model (~25MB, first run only)...")
        urllib.request.urlretrieve(MODEL_URL, MODEL_PATH)
        print("Model downloaded.")


# ── Landmark fallback ─────────────────────

def finger_extended_score(lm):
    fingers  = [(8,6),(12,10),(16,14),(20,18)]
    extended = sum(1 for tip, pip in fingers if lm[tip].y < lm[pip].y)
    return extended / 4.0


# ── Open camera ───────────────────────────

def open_camera_fresh():
    """Try to open the webcam, retry up to 10 times (handles post-unlock delay)."""
    for _ in range(10):
        c = cv2.VideoCapture(CAMERA_INDEX)
        if c.isOpened():
            c.set(cv2.CAP_PROP_FRAME_WIDTH,  640)
            c.set(cv2.CAP_PROP_FRAME_HEIGHT, 480)
            return c
        c.release()
        time.sleep(1)
    return None


# ── Gesture logic ─────────────────────────

def process_gesture(gesture_name):
    global lock_armed, fist_start, status_text

    if gesture_name == "Open_Palm":
        if not lock_armed:
            lock_armed  = True
            fist_start  = None
            status_text = "Palm detected — close fist to lock"

    elif gesture_name == "Closed_Fist" and lock_armed:
        if fist_start is None:
            fist_start  = time.time()
            status_text = "Fist held — locking soon..."
        elif time.time() - fist_start >= GESTURE_HOLD_SECONDS:
            status_text = "Locking..."
            lock_armed  = False
            fist_start  = None
            lock_windows()   # ← releases camera first, then locks
    else:
        fist_start = None
        if not lock_armed:
            status_text = "Watching..."


# ── Detection loop ────────────────────────

def detection_loop():
    global is_running, lock_armed, fist_start, status_text

    use_legacy = False
    recognizer = None

    try:
        download_model()
        base_options = mp_python.BaseOptions(model_asset_path=MODEL_PATH)
        options      = GestureRecognizerOptions(
            base_options=base_options,
            running_mode=RunningMode.VIDEO,
        )
        recognizer = GestureRecognizer.create_from_options(options)
        print("Using mediapipe Gesture Recognizer (new API)")
    except Exception as e:
        print(f"New API failed ({e}), falling back...")
        use_legacy = True

    if use_legacy:
        detection_loop_legacy()
        return

    import mediapipe as mp

    frame_timestamp = 0
    failed_frames   = 0

    while True:
        # Paused or locked — make sure camera is released
        if screen_locked or not is_running:
            lock_armed = False
            fist_start = None
            if not screen_locked:
                status_text = "Paused"
            release_cap()
            time.sleep(0.3)
            continue

        # Camera not open — open it
        cap = get_cap()
        if cap is None or not cap.isOpened():
            status_text = "Starting camera..."
            release_cap()
            time.sleep(1)
            new_cap = open_camera_fresh()
            set_cap(new_cap)
            failed_frames = 0
            continue

        ret, frame = cap.read()
        if not ret:
            failed_frames += 1
            if failed_frames > 30:
                release_cap()
                failed_frames = 0
                status_text   = "Reconnecting camera..."
            time.sleep(0.05)
            continue

        failed_frames    = 0
        frame_timestamp += CHECK_INTERVAL_MS
        rgb              = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
        mp_image         = mp.Image(image_format=mp.ImageFormat.SRGB, data=rgb)
        result           = recognizer.recognize_for_video(mp_image, frame_timestamp)

        gesture_name = result.gestures[0][0].category_name if result.gestures else None
        process_gesture(gesture_name)
        time.sleep(CHECK_INTERVAL_MS / 1000.0)

    release_cap()
    recognizer.close()


def detection_loop_legacy():
    global is_running, lock_armed, fist_start, status_text

    import mediapipe as mp
    mp_hands      = mp.solutions.hands
    failed_frames = 0

    with mp_hands.Hands(
        static_image_mode=False,
        max_num_hands=1,
        min_detection_confidence=0.7,
        min_tracking_confidence=0.5,
    ) as hands:
        while True:
            if screen_locked or not is_running:
                lock_armed = False
                fist_start = None
                if not screen_locked:
                    status_text = "Paused"
                release_cap()
                time.sleep(0.3)
                continue

            cap = get_cap()
            if cap is None or not cap.isOpened():
                status_text = "Starting camera..."
                release_cap()
                time.sleep(1)
                set_cap(open_camera_fresh())
                failed_frames = 0
                continue

            ret, frame = cap.read()
            if not ret:
                failed_frames += 1
                if failed_frames > 30:
                    release_cap()
                    failed_frames = 0
                time.sleep(0.05)
                continue
            failed_frames = 0

            rgb     = cv2.cvtColor(frame, cv2.COLOR_BGR2RGB)
            results = hands.process(rgb)

            if results.multi_hand_landmarks:
                score   = finger_extended_score(results.multi_hand_landmarks[0].landmark)
                gesture = "Open_Palm" if score >= 0.6 else ("Closed_Fist" if score <= 0.25 else None)
            else:
                gesture = None

            process_gesture(gesture)
            time.sleep(CHECK_INTERVAL_MS / 1000.0)

    release_cap()


# System tray 

def create_tray_image(color="green"):
    size  = 64
    img   = Image.new("RGBA", (size, size), (0, 0, 0, 0))
    draw  = ImageDraw.Draw(img)
    fills = {
        "green":  (34,  197, 94),
        "red":    (239, 68,  68),
        "orange": (251, 146, 60),
    }
    draw.ellipse([4, 4, size-4, size-4], fill=fills.get(color, fills["green"]))
    return img


def on_pause_resume(icon, item_obj):
    global is_running
    is_running = not is_running
    icon.icon  = create_tray_image("green" if is_running else "red")
    icon.title = f"GestureLock — {'Active' if is_running else 'Paused'}"


def on_lock_now(icon, item_obj):
    lock_windows()


def on_quit(icon, item_obj):
    icon.stop()
    os._exit(0)


def status_title(item_obj):
    return f"Status: {status_text}"


def build_tray():
    global tray_icon_ref
    icon = pystray.Icon(
        "GestureLock",
        create_tray_image("green"),
        "GestureLock — Active",
        menu=pystray.Menu(
            item(status_title, None, enabled=False),
            pystray.Menu.SEPARATOR,
            item("Pause / Resume", on_pause_resume),
            item("Lock Now",       on_lock_now),
            pystray.Menu.SEPARATOR,
            item("Quit",           on_quit),
        )
    )
    tray_icon_ref = icon
    return icon


# ── Entry point ───────────────────────────

if __name__ == "__main__":
    threading.Thread(target=start_session_watcher, daemon=True).start()
    threading.Thread(target=detection_loop,        daemon=True).start()
    build_tray().run()