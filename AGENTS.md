# AGENTS.md

This file provides guidance to Codex (Codex.ai/code) when working with code in this repository.

## Project Overview

EasyChat Momo is a PC WeChat automation assistant focused on keyword-triggered auto-replies. It uses `uiautomation` to watch an independent WeChat chat window and send an image or text reply after an optional delay.

Important notes:
- The maintained GUI entry is `wechat_gui_momo.py`.
- Configuration is stored in `wechat_config_momo.json`.
- The project targets WeChat 4.1 desktop UI behavior.

## Prerequisites

**CRITICAL**: Windows Narrator mode must be enabled for `uiautomation` to correctly detect WeChat controls. This is the most common cause of runtime failures.

## Key Dependencies

Core dependencies from `requirements.txt`:
- `PyQt5 5.15.7`
- `uiautomation 2.0.17`
- `pyperclip 1.8.2`
- `pyinstaller 5.12.0`
- `pywin32 304`

## Architecture

### Core Components

1. `wechat_gui_momo.py`
   - Main PyQt5 GUI and current application entry.
   - Manages rules, delay settings, timed monitoring windows, logging, and config persistence.
   - Cancels delayed sends when a later message clears the alert state.

2. `ui_auto_wechat.py`
   - Core WeChat automation layer.
   - Finds chat windows, monitors the latest message, sends text, and sends files.

3. `wechat_locale.py`
   - Localized control names for different WeChat language settings.

4. `clipboard.py`
   - File clipboard helper used by the automation layer.

5. `tools/automation.py`
   - Helper for inspecting the WeChat control tree when UI changes break automation.

## Development Commands

### Setup

```bash
pip install -r requirements.txt
```

### Run the application

```bash
py wechat_gui_momo.py
```

### Build the executable

```bash
py pack.py
```

This uses `wechat_gui_momo.spec` and outputs `dist/wechat_gui_momo.exe`.

### Build the portable zip package

```bash
py pack.py --portable
```

This creates `wechat_gui_momo_portable.zip`.

## Configuration Notes

`wechat_config_momo.json` stores:
- `settings`: WeChat path, language, target sender, active rule count, delay settings, and auto-timer options
- `rules`: up to 5 keyword rules, each with reply type, folder or text content, and match mode

Rule behavior:
- Match mode supports exact match and contains match.
- Delayed replies are invalidated if a later message no longer matches a keyword.
- Image replies remove the sent image from the source folder after a successful send.

## Common Pitfalls

1. The target chat must be an independent WeChat chat window whose title matches the configured trigger sender.
2. Narrator mode must be enabled or control detection will fail.
3. WeChat UI depth and control names are version-sensitive; use `tools/automation.py` when upgrading WeChat.
4. If the image folder is empty, image rules will log an error and skip sending.
5. Some legacy methods in `ui_auto_wechat.py` remain `NotImplementedError` for WeChat 4.1 and are not used by the momo GUI.
