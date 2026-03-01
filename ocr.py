"""
OCR Tool - Region select -> llama.cpp vision OCR -> clipboard
Auto-downloads models and llama.cpp on first run.
"""

import subprocess
import sys
import os
import io
import json
import re
import tempfile
import ctypes
import threading
import zipfile
import tkinter as tk
from tkinter import ttk
from pathlib import Path

# ─── Auto-install Pillow ─────────────────────────────────────
try:
    from PIL import Image, ImageGrab, ImageTk, ImageEnhance, ImageFilter
except ImportError:
    subprocess.check_call(
        [sys.executable, "-m", "pip", "install", "Pillow"],
        creationflags=getattr(subprocess, "CREATE_NO_WINDOW", 0),
    )
    from PIL import Image, ImageGrab, ImageTk, ImageEnhance, ImageFilter

# ─── DPI Awareness ──────────────────────────────────────────
try:
    ctypes.windll.shcore.SetProcessDpiAwareness(2)
except Exception:
    try:
        ctypes.windll.user32.SetProcessDPIAware()
    except Exception:
        pass

# Fix pythonw
if len(sys.argv) <= 1:
    if sys.stdout is None:
        sys.stdout = io.StringIO()
    if sys.stderr is None:
        sys.stderr = io.StringIO()

# ─── Find the real script directory ──────────────────────────
# PyInstaller --onefile extracts to a temp folder
# We need the folder where the .exe actually lives
if getattr(sys, 'frozen', False):
    # Running as compiled .exe
    SCRIPT_DIR = Path(sys.executable).parent
else:
    # Running as .py script
    SCRIPT_DIR = Path(__file__).parent
LOG_FILE = SCRIPT_DIR / "ocr_debug.log"


# ─── Logging ─────────────────────────────────────────────────

def log(msg, settings=None):
    debug = settings.get("debug", False) if settings else False
    if debug:
        print(msg)
        with open(LOG_FILE, "a", encoding="utf-8") as f:
            f.write(msg + "\n")


def show_debug_popup(title, msg):
    ctypes.windll.user32.MessageBoxW(0, str(msg)[:1000], str(title), 0)


# ─── Settings ───────────────────────────────────────────────

DEFAULT_SETTINGS = {
    "hotkey": "ctrl+shift+p",
    "active_model": "glm-ocr",
    "models": {
        "glm-ocr": {
            "name": "GLM-OCR (Accurate)",
            "model_path": "models/GLM-OCR.-Q4_K_M.gguf",
            "mmproj_path": "models/GLM-OCR.mmproj-Q4_K_M.gguf",
            "model_url": "https://huggingface.co/nopesadly/GLM-OCR-Q4_K_M.gguf/resolve/main/GLM-OCR.-Q4_K_M.gguf",
            "mmproj_url": "https://huggingface.co/nopesadly/GLM-OCR-Q4_K_M.gguf/resolve/main/GLM-OCR.mmproj-Q4_K_M.gguf",
            "prompt": "<image>\nExtract only the exact text from this image. Output the text exactly as written, with no extra words.",
            "extra_args": [
                "--chat-template", "vicuna",
                "--top-k", "1",
                "--repeat-penalty", "1.0",
                "-c", "1024",
            ],
        },
        "smolvlm": {
            "name": "SmolVLM 500M (Fast)",
            "model_path": "models/SmolVLM-500M-Instruct-Q8_0.gguf",
            "mmproj_path": "models/mmproj-SmolVLM-500M-Instruct-Q8_0.gguf",
            "model_url": "https://huggingface.co/ggml-org/SmolVLM-500M-Instruct-GGUF/resolve/main/SmolVLM-500M-Instruct-Q8_0.gguf",
            "mmproj_url": "https://huggingface.co/ggml-org/SmolVLM-500M-Instruct-GGUF/resolve/main/mmproj-SmolVLM-500M-Instruct-Q8_0.gguf",
            "prompt": "OCR: Read every character in this image exactly.",
            "extra_args": ["--no-warmup"],
        },
    },
    "llama_dir": "llama",
    "llama_url": "https://github.com/ggml-org/llama.cpp/releases/download/b8183/llama-b8183-bin-win-vulkan-x64.zip",
    "ngl": 99,
    "preprocess": False,
    "show_toast": False,
    "play_sound": True,
    "debug": False,
}


def load_settings():
    settings_path = SCRIPT_DIR / "settings.json"
    try:
        with open(settings_path, "r") as f:
            user = json.load(f)
        if "models" not in user:
            raise ValueError("Old format")
        return {**DEFAULT_SETTINGS, **user}
    except (FileNotFoundError, json.JSONDecodeError, ValueError):
        with open(settings_path, "w") as f:
            json.dump(DEFAULT_SETTINGS, f, indent=2)
        return DEFAULT_SETTINGS.copy()


def save_settings(settings):
    with open(SCRIPT_DIR / "settings.json", "w") as f:
        json.dump(settings, f, indent=2)


def get_active_model(settings):
    active_id = settings.get("active_model", "glm-ocr")
    models = settings.get("models", {})
    if active_id not in models:
        active_id = next(iter(models))
    return active_id, models[active_id]


def resolve_path(path_str):
    p = Path(path_str)
    return p if p.is_absolute() else SCRIPT_DIR / p




# ─── CLI Helpers (for AHK) ──────────────────────────────────

def cli_list_models():
    settings = load_settings()
    active = settings.get("active_model", "")
    for model_id, cfg in settings.get("models", {}).items():
        name = cfg.get("name", model_id)
        is_active = "true" if model_id == active else "false"
        print(f"{model_id}|{name}|{is_active}")


def cli_set_model(model_id):
    settings = load_settings()
    if model_id in settings.get("models", {}):
        settings["active_model"] = model_id
        save_settings(settings)

# ─── Download / Setup ───────────────────────────────────────

def check_missing(settings):
    """Returns list of (type, label, url, dest) for missing files."""
    missing = []

    # Check llama
    llama_dir = resolve_path(settings.get("llama_dir", "llama"))
    cli_path = llama_dir / "llama-mtmd-cli.exe"
    if not cli_path.exists():
        url = settings.get("llama_url", "")
        missing.append(("llama", "llama.cpp (Vulkan)", url, str(llama_dir)))

    # Check ALL model files (not just active)
    for model_id, cfg in settings.get("models", {}).items():
        model_path = resolve_path(cfg["model_path"])
        mmproj_path = resolve_path(cfg["mmproj_path"])

        if not model_path.exists():
            url = cfg.get("model_url", "")
            missing.append(("model", f"{cfg.get('name', model_id)} - model", url, str(model_path)))

        if not mmproj_path.exists():
            url = cfg.get("mmproj_url", "")
            missing.append(("mmproj", f"{cfg.get('name', model_id)} - vision", url, str(mmproj_path)))

    return missing


def download_file(url, dest, progress_callback=None):
    """Download a file with progress reporting."""
    import urllib.request
    import urllib.error

    req = urllib.request.Request(url, headers={"User-Agent": "OCR-Tool/1.0"})
    response = urllib.request.urlopen(req, timeout=30)
    total = int(response.headers.get("Content-Length", 0))
    downloaded = 0
    block_size = 1024 * 64  # 64KB chunks

    dest_path = Path(dest)
    dest_path.parent.mkdir(parents=True, exist_ok=True)

    with open(dest_path, "wb") as f:
        while True:
            chunk = response.read(block_size)
            if not chunk:
                break
            f.write(chunk)
            downloaded += len(chunk)
            if progress_callback:
                progress_callback(downloaded, total)


def extract_llama_zip(zip_path, dest_dir):
    """Extract only needed files from llama.cpp release zip."""
    dest = Path(dest_dir)
    dest.mkdir(parents=True, exist_ok=True)

    with zipfile.ZipFile(zip_path, "r") as z:
        for member in z.namelist():
            filename = Path(member).name
            if not filename:
                continue
            if filename.endswith(".exe") or filename.endswith(".dll"):
                data = z.read(member)
                (dest / filename).write_bytes(data)

    Path(zip_path).unlink()


class SetupWindow:
    """First-run setup GUI with download progress."""

    def __init__(self, missing, settings):
        self.missing = missing
        self.settings = settings
        self.success = False
        self.error_msg = ""

        # Thread communication
        self.current_status = "Preparing..."
        self.current_detail = ""
        self.current_progress = 0
        self.download_done = False

        self.root = tk.Tk()
        self.root.title("OCR Tool — First Run Setup")
        self.root.geometry("500x220")
        self.root.resizable(False, False)
        self.root.attributes("-topmost", True)
        self.root.configure(bg="#1a1a2e")

        # Center
        self.root.update_idletasks()
        x = (self.root.winfo_screenwidth() - 500) // 2
        y = (self.root.winfo_screenheight() - 220) // 2
        self.root.geometry(f"+{x}+{y}")

        # Title
        tk.Label(
            self.root, text="🔧  First-Time Setup",
            font=("Segoe UI", 14, "bold"), fg="white", bg="#1a1a2e",
        ).pack(pady=(20, 10))

        # Status
        self.status_label = tk.Label(
            self.root, text="Preparing...",
            font=("Segoe UI", 10), fg="#cccccc", bg="#1a1a2e",
        )
        self.status_label.pack(pady=(5, 5))

        # Progress bar
        style = ttk.Style()
        style.theme_use("default")
        style.configure(
            "Custom.Horizontal.TProgressbar",
            troughcolor="#2a2a4a", background="#0078D7", thickness=20,
        )
        self.progress = ttk.Progressbar(
            self.root, length=420, mode="determinate",
            style="Custom.Horizontal.TProgressbar",
        )
        self.progress.pack(pady=(5, 5))

        # Detail
        self.detail_label = tk.Label(
            self.root, text="",
            font=("Segoe UI", 9), fg="#888888", bg="#1a1a2e",
        )
        self.detail_label.pack(pady=(5, 10))

        # Close handler
        self.root.protocol("WM_DELETE_WINDOW", self._on_close)

    def _on_close(self):
        self.download_done = True
        self.root.destroy()

    def run(self):
        thread = threading.Thread(target=self._download_thread, daemon=True)
        thread.start()
        self._poll()
        self.root.mainloop()
        return self.success

    def _poll(self):
        if self.download_done:
            try:
                if self.error_msg:
                    self.status_label.config(text=f"❌  {self.error_msg}", fg="#ff6666")
                    self.root.after(3000, self.root.destroy)
                else:
                    self.status_label.config(text="✅  Setup complete!", fg="#66ff66")
                    self.root.after(1500, self.root.destroy)
            except tk.TclError:
                pass
            return

        try:
            self.status_label.config(text=self.current_status)
            self.detail_label.config(text=self.current_detail)
            self.progress["value"] = self.current_progress
        except tk.TclError:
            return

        self.root.after(100, self._poll)

    def _progress_cb(self, downloaded, total):
        if total > 0:
            self.current_progress = (downloaded / total) * 100
            mb_down = downloaded / (1024 * 1024)
            mb_total = total / (1024 * 1024)
            self.current_detail = f"{mb_down:.1f} MB / {mb_total:.1f} MB"
        else:
            mb_down = downloaded / (1024 * 1024)
            self.current_detail = f"{mb_down:.1f} MB downloaded"

    def _download_thread(self):
        try:
            total_items = len(self.missing)

            for i, (file_type, label, url, dest) in enumerate(self.missing):
                self.current_status = f"Downloading {label}... ({i + 1}/{total_items})"
                self.current_progress = 0
                self.current_detail = "Starting..."

                if not url or "YOUR_USERNAME" in url:
                    self.error_msg = f"No download URL for {label}. Update settings.json"
                    self.download_done = True
                    return

                if file_type == "llama":
                    # Download zip, extract
                    zip_path = Path(dest) / "llama_temp.zip"
                    Path(dest).mkdir(parents=True, exist_ok=True)
                    download_file(url, str(zip_path), self._progress_cb)

                    self.current_status = f"Extracting {label}..."
                    self.current_detail = "Unpacking files..."
                    extract_llama_zip(str(zip_path), dest)
                else:
                    # Direct file download
                    download_file(url, dest, self._progress_cb)

            self.success = True
        except Exception as e:
            self.error_msg = str(e)[:200]
        finally:
            self.download_done = True


# ─── Virtual Screen ──────────────────────────────────────────

def get_virtual_screen():
    user32 = ctypes.windll.user32
    x = user32.GetSystemMetrics(76)
    y = user32.GetSystemMetrics(77)
    w = user32.GetSystemMetrics(78)
    h = user32.GetSystemMetrics(79)
    return x, y, w, h


# ─── Status Pill ─────────────────────────────────────────────

class StatusPill:
    def __init__(self):
        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.attributes("-alpha", 0.92)
        self.root.configure(bg="#1a1a2e")
        self.root.geometry("+20+20")

        self.frame = tk.Frame(
            self.root, bg="#1a1a2e", padx=16, pady=10,
            highlightbackground="#0078D7", highlightthickness=2,
        )
        self.frame.pack()

        self.label = tk.Label(
            self.frame, text="✂️  Captured",
            font=("Segoe UI", 11), fg="white", bg="#1a1a2e",
        )
        self.label.pack()
        self.root.update()

    def update_status(self, text):
        try:
            self.label.config(text=text)
            self.root.update()
        except tk.TclError:
            pass

    def close(self):
        try:
            self.root.destroy()
        except tk.TclError:
            pass


# ─── Image Preprocessing ────────────────────────────────────

def preprocess_image(image, settings=None):
    w, h = image.size
    if image.mode != "RGB":
        image = image.convert("RGB")

    min_dim = 512
    if w < min_dim or h < min_dim:
        scale = min(max(min_dim / w, min_dim / h, 2.0), 6.0)
        image = image.resize((int(w * scale), int(h * scale)), Image.LANCZOS)

    max_dim = 1024
    if image.width > max_dim or image.height > max_dim:
        ratio = min(max_dim / image.width, max_dim / image.height)
        image = image.resize((int(image.width * ratio), int(image.height * ratio)), Image.LANCZOS)

    image = image.filter(ImageFilter.SHARPEN)
    image = ImageEnhance.Contrast(image).enhance(1.4)

    pad = 40
    padded = Image.new("RGB", (image.width + pad * 2, image.height + pad * 2), (255, 255, 255))
    padded.paste(image, (pad, pad))
    return padded


# ─── Region Selector ────────────────────────────────────────

class RegionSelector:
    def __init__(self):
        self.region = None
        self.start_x = 0
        self.start_y = 0
        self.vx, self.vy, self.vw, self.vh = get_virtual_screen()

        self.screenshot = ImageGrab.grab(
            bbox=(self.vx, self.vy, self.vx + self.vw, self.vy + self.vh),
            all_screens=True,
        )
        self.dark = ImageEnhance.Brightness(self.screenshot).enhance(0.5)

        self.root = tk.Tk()
        self.root.overrideredirect(True)
        self.root.attributes("-topmost", True)
        self.root.geometry(f"{self.vw}x{self.vh}+{self.vx}+{self.vy}")
        self.root.configure(cursor="cross")

        self.canvas = tk.Canvas(
            self.root, width=self.vw, height=self.vh,
            highlightthickness=0, cursor="cross",
        )
        self.canvas.pack(fill="both", expand=True)

        self.dark_photo = ImageTk.PhotoImage(self.dark)
        self.canvas.create_image(0, 0, anchor="nw", image=self.dark_photo)

        self.canvas.bind("<ButtonPress-1>", self._on_press)
        self.canvas.bind("<B1-Motion>", self._on_drag)
        self.canvas.bind("<ButtonRelease-1>", self._on_release)
        self.root.bind("<Escape>", lambda e: self._cancel())

        self._bright_region_id = None
        self._bright_region_photo = None

    def _on_press(self, event):
        self.start_x = event.x
        self.start_y = event.y

    def _on_drag(self, event):
        self.canvas.delete("sel")
        x1, y1 = min(self.start_x, event.x), min(self.start_y, event.y)
        x2, y2 = max(self.start_x, event.x), max(self.start_y, event.y)
        if x2 - x1 < 5 or y2 - y1 < 5:
            return
        region = self.screenshot.crop((x1, y1, x2, y2))
        self._bright_region_photo = ImageTk.PhotoImage(region)
        if self._bright_region_id:
            self.canvas.delete(self._bright_region_id)
        self._bright_region_id = self.canvas.create_image(
            x1, y1, anchor="nw", image=self._bright_region_photo, tags="sel",
        )
        self.canvas.create_rectangle(x1, y1, x2, y2, outline="#0078D7", width=2, tags="sel")

    def _on_release(self, event):
        x1, y1 = min(self.start_x, event.x), min(self.start_y, event.y)
        x2, y2 = max(self.start_x, event.x), max(self.start_y, event.y)
        if (x2 - x1) > 10 and (y2 - y1) > 10:
            self.region = (x1, y1, x2, y2)
        self.root.destroy()

    def _cancel(self):
        self.region = None
        self.root.destroy()

    def run(self):
        self.root.mainloop()
        return self.screenshot.crop(self.region) if self.region else None


# ─── OCR ─────────────────────────────────────────────────────

def run_ocr(image_path, settings):
    model_id, model_cfg = get_active_model(settings)
    llama_dir = resolve_path(settings.get("llama_dir", "llama"))
    cli_path = llama_dir / "llama-mtmd-cli.exe"
    model_path = resolve_path(model_cfg["model_path"])
    mmproj_path = resolve_path(model_cfg["mmproj_path"])
    prompt = model_cfg.get("prompt", "Extract all text from this image.")
    extra_args = [str(a) for a in model_cfg.get("extra_args", [])]

    cmd = [
        str(cli_path),
        "-m", str(model_path),
        "--mmproj", str(mmproj_path),
        "--image", str(image_path),
        "-p", prompt,
        "-ngl", str(settings.get("ngl", 99)),
        "-n", "512",
        "--temp", "0",
    ] + extra_args

    log(f"[OCR] Model: {model_id}", settings)
    log(f"[OCR] Command: {' '.join(cmd)}", settings)

    for label, path in [("CLI", cli_path), ("Model", model_path),
                         ("MMProj", mmproj_path), ("Image", image_path)]:
        if not Path(path).exists():
            return f"ERROR: {label} not found at {path}"

    proc = subprocess.Popen(
        cmd, stdout=subprocess.PIPE, stderr=subprocess.PIPE,
        stdin=subprocess.PIPE, cwd=str(llama_dir),
        creationflags=subprocess.CREATE_NO_WINDOW,
    )

    try:
        stdout, stderr = proc.communicate(input=b"/exit\n", timeout=30)
    except subprocess.TimeoutExpired:
        proc.kill()
        proc.communicate()
        return "ERROR: Timeout"

    stdout_text = stdout.decode("utf-8", errors="replace")
    log(f"[OCR] Exit code: {proc.returncode}", settings)
    log(f"[OCR] Raw stdout: '{stdout_text}'", settings)

    if proc.returncode != 0:
        return f"ERROR: Exit code {proc.returncode}"

    return parse_output(stdout_text)


def parse_output(raw):
    text = re.sub(r"\x1b\[[0-9;]*m", "", raw)
    text = text.replace("/exit", "")
    lines = []
    for line in text.split("\n"):
        s = line.strip()
        if s.startswith("[") and ("t/s" in s or "Prompt:" in s):
            continue
        if s and all(c in "_-=" for c in s):
            continue
        if s.startswith(">"):
            continue
        lines.append(line)
    return clean_ocr_result("\n".join(lines).strip())


def clean_ocr_result(text):
    lines = text.split("\n")

    # Strip leading lines that are just 1-3 digit numbers (artifacts)
    while lines:
        s = lines[0].strip()
        if not s:
            lines.pop(0)
        elif re.match(r"^\d{1,3}$", s):
            lines.pop(0)
        else:
            break

    # Strip trailing empty lines
    while lines and not lines[-1].strip():
        lines.pop()

    result = "\n".join(lines).strip()

    # Remove chat template artifacts
    for prefix in ["ASSISTANT:", "ASSISTANT", "Assistant:", "Assistant",
                    "USER:", "USER", "User:", "User",
                    "A:", "Output:", "Result:", "Answer:"]:
        if result.upper().startswith(prefix.upper()):
            result = result[len(prefix):].strip()

    return result.strip()


# ─── Clipboard ───────────────────────────────────────────────

def copy_to_clipboard(text):
    root = tk.Tk()
    root.withdraw()
    root.clipboard_clear()
    root.clipboard_append(text)
    root.update()
    root.destroy()


# ─── Toast ───────────────────────────────────────────────────

def show_toast(text):
    safe = text[:200].replace("'", "`'").replace("\n", " ")
    ps = f"""
    Add-Type -AssemblyName System.Windows.Forms
    $n = New-Object System.Windows.Forms.NotifyIcon
    $n.Icon = [System.Drawing.SystemIcons]::Information
    $n.BalloonTipTitle = 'OCR Complete'
    $n.BalloonTipText = 'Copied: {safe}'
    $n.Visible = $true
    $n.ShowBalloonTip(3000)
    Start-Sleep 4
    $n.Dispose()
    """
    subprocess.Popen(
        ["powershell", "-WindowStyle", "Hidden", "-Command", ps],
        creationflags=subprocess.CREATE_NO_WINDOW,
    )


# ─── Main ────────────────────────────────────────────────────

def main():
    settings = load_settings()

    if settings.get("debug") and LOG_FILE.exists():
        LOG_FILE.unlink()

    log("=" * 50, settings)
    log("[START] OCR Tool", settings)

    # ── Check for missing files → auto download ──
    missing = check_missing(settings)
    if missing:
        log(f"[SETUP] Missing {len(missing)} files, starting download...", settings)
        setup = SetupWindow(missing, settings)
        ok = setup.run()
        if not ok:
            show_debug_popup("Setup Failed", "Could not download required files.\nCheck settings.json URLs.")
            sys.exit(1)

    # ── Region select ──
    model_id, model_cfg = get_active_model(settings)
    model_name = model_cfg.get("name", model_id)

    selector = RegionSelector()
    image = selector.run()

    if image is None:
        sys.exit(0)

    log(f"[STEP 1] Captured: {image.size}", settings)

    # ── Status pill ──
    pill = StatusPill()

    try:
        # Preprocess (if enabled)
        if settings.get("preprocess", False):
            pill.update_status("🔄  Processing image...")
            image = preprocess_image(image, settings)
        else:
            if image.mode != "RGB":
                image = image.convert("RGB")

        # Save temp
        temp = Path(tempfile.gettempdir()) / "ocr_capture.png"
        image.save(str(temp), "PNG")

        # OCR
        pill.update_status(f"🔍  {model_name}...")
        text = run_ocr(temp, settings)

        if text and not text.startswith("ERROR"):
            copy_to_clipboard(text)
            log(f"[DONE] Copied: '{text}'", settings)

            pill.update_status("✅  Copied!")

            if settings.get("play_sound", True):
                import winsound
                winsound.MessageBeep()
            if settings.get("show_toast", False):
                show_toast(text)
            if settings.get("debug", False):
                show_debug_popup("OCR Result", f"Copied:\n\n{text}")

            pill.root.after(1500, pill.close)
            pill.root.mainloop()
        else:
            log(f"[ERROR] {text}", settings)
            pill.update_status("❌  No text found")

            if settings.get("play_sound", True):
                import winsound
                winsound.MessageBeep(winsound.MB_ICONHAND)
            if settings.get("debug", False):
                show_debug_popup("OCR Failed", f"Result: {text}")

            pill.root.after(2000, pill.close)
            pill.root.mainloop()

    except Exception as e:
        log(f"[CRASH] {e}", settings)
        if settings.get("debug"):
            show_debug_popup("Crash", str(e))
        pill.close()

    finally:
        temp_path = Path(tempfile.gettempdir()) / "ocr_capture.png"
        if temp_path.exists():
            temp_path.unlink()


# ─── Entry Point ─────────────────────────────────────────────

if __name__ == "__main__":
    if len(sys.argv) > 1:
        if sys.argv[1] == "--list-models":
            cli_list_models()
        elif sys.argv[1] == "--set-model" and len(sys.argv) > 2:
            cli_set_model(sys.argv[2])
        sys.exit(0)
    main()
