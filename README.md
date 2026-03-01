# AI Screen OCR Tool

A powerful, privacy-friendly Screen OCR (Optical Character Recognition) tool powered by local Vision-Language Models (VLMs) and `llama.cpp`. 
Select any region on your screen, and the tool will use AI (like GLM-OCR or SmolVLM) to perfectly extract the text and copy it directly to your clipboard. No cloud APIs, no subscriptions, fully local!

## Features

- ** Local AI Powered:** Uses advanced Vision models (GLM-OCR & SmolVLM) for highly accurate text extraction, even on complex backgrounds or messy formatting.
- ** Auto-Setup:** Automatically downloads the required `llama.cpp` (Vulkan GPU-accelerated) binaries and AI models on the first run.
- ** Global Hotkeys & Tray Icon:** Managed by AutoHotkey v2 for minimal resource usage. Trigger OCR from anywhere, switch models on the fly.
- ** Visual Feedback:** Features a floating "Status Pill" and Windows notifications so you know exactly when the text is ready to paste.
- ** Image Preprocessing:** Optional image upscaling, sharpening, and contrast enhancement to help the AI read tiny or blurry text.

## Prerequisites

This tool is designed for **Windows** (uses Windows APIs, AHK, and Windows Vulkan binaries).

1. **[Python 3.8+](https://www.python.org/downloads/)** (Make sure to check "Add Python to PATH" during installation)
2. **[AutoHotkey v2](https://www.autohotkey.com/)** (To run the tray and hotkey manager)
3. *Note: The Python script will automatically install `Pillow` via pip if it's missing.*

##  Installation & Usage

### 1. Running from Source
1. Clone or download this repository.
2. Double-click the AutoHotkey script (`hotkey.ahk`). 
3. A tray icon will appear in your Windows taskbar.
4. Press **`Ctrl + Shift + P`** (default) to start your first capture!

### 2. First-Time Setup
On your very first run, a setup window will appear. It will automatically download:
- `llama.cpp` (Windows Vulkan backend for fast GPU inference)
- The default Vision Models (GLM-OCR and SmolVLM)
*Depending on your internet speed, this may take a few minutes. Once downloaded, it runs entirely offline.*

## Configuration (`settings.json`)

After the first run, a `settings.json` file will be generated in the app directory. You can access it easily by right-clicking the Tray Icon and selecting **Settings**.

### Key Settings:
*   `"hotkey"`: Change your global shortcut (e.g., `"win+shift+s"`, `"ctrl+alt+o"`).
*   `"active_model"`: Set your default model (`"glm-ocr"` for accuracy, `"smolvlm"` for speed).
*   `"preprocess"`: Set to `true` to enable automatic image upscaling and contrast enhancement before OCR.
*   `"play_sound"`: Play a Windows beep on success/failure (`true`/`false`).
*   `"show_toast"`: Show a Windows notification when text is copied to the clipboard (`true`/`false`).
*   `"debug"`: Set to `true` to generate an `ocr_debug.log` file and show popup errors for troubleshooting.

### Adding Custom Models
You can add other GGUF vision models to the `"models"` block in `settings.json`! Just provide the download URLs, local paths, and the required prompt structure for that specific model.

```json
"models": {
  "my-custom-model": {
    "name": "Custom Model Name",
    "model_path": "models/my-model-Q4.gguf",
    "mmproj_path": "models/my-model-mmproj.gguf",
    "model_url": "https://huggingface.co/...",
    "mmproj_url": "https://huggingface.co/...",
    "prompt": "Extract text:",
    "extra_args":["--top-k", "1", "-c", "1024"]
  }
}
```

## Tray Menu Options
Right-click the tool's icon in your system tray to:
*   **Scan Region:** Manually trigger an OCR scan without the hotkey.
*   **Models:** Switch seamlessly between downloaded AI models.
*   **Settings:** Opens the `settings.json` file.
*   **Start with Windows:** Toggles an autostart shortcut in your Windows Startup folder.
*   **Exit:** Closes the hotkey listener and background processes.

##  Troubleshooting

* **Nothing happens when I press the hotkey:** Ensure AutoHotkey v2 is installed and running. Check the tray icon.
* **OCR is returning garbage text:** Make sure you aren't selecting an area that is entirely blank. If the text is very small, enable `"preprocess": true` in the settings.
* **Crash / Silent Failure:** Open `settings.json`, set `"debug": true`, and try again. Check `ocr_debug.log` in the script directory for specific error messages regarding `llama.cpp` or file paths.

---
*Built with [Python](https://python.org/), [AutoHotkey](https://www.autohotkey.com/), and [llama.cpp](https://github.com/ggerganov/llama.cpp).*
