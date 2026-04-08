# JIT-OCR

Just-In-Time OCR — screen region select → instant text extraction → clipboard.

Runs entirely offline using llama.cpp vision models. No API keys, no cloud, no BS.

## Features

- **Resource efficient** — runs on potato hardware
- **Minimum 2GB VRAM or RAM** — works on integrated graphics
- **Fully offline** — no internet needed after first download
- **Auto-downloads everything** — just run it

## How It Works

1. Press hotkey (default: `Ctrl+Shift+P`)
2. Select a screen region
3. Text gets extracted and copied to clipboard
4. Done

## First Run/Installation

Run ocr.exe once and then the hotkey should work.

Everything downloads automatically on first launch:
- llama.cpp (Vulkan)
- OCR model files

Just run it and wait.

## Models

| Model | Speed | Accuracy | VRAM |
|---|---|---|---|
| GLM-OCR | Moderate | High | ~1.5GB |
| PaddleOCR | Moderate | High | ~1.5GB |
| SmolVLM 500M | Fast | Good enough | ~1GB |

Switch models from the system tray icon.

## Files

- `ocr.ahk` / `hotkey.exe` — hotkey listener + tray menu
- `ocr.py` / `ocr.exe` — the actual OCR engine
- `settings.json` — config (hotkey, models, etc.)

## Settings

Edit `settings.json` or right-click tray → Settings.

```json
{
  "hotkey": "ctrl+shift+p",
  "active_model": "glm-ocr",
  "ngl": 99,
  "preprocess": false,
  "debug": false
}
