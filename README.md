# Windows Hover Folder Size

A lightweight Windows 10/11 utility that shows the size of files and folders when you hover over them in **File Explorer**, without opening Properties.

## Features

- Hover over a file for its size.
- Hover over a folder for the recursive total size of its contents.
- Asynchronous folder scanning so File Explorer does not freeze.
- Progress updates for large folders.
- Two-minute cache for recent folder calculations.
- Skips directory junctions, reparse points, and symbolic links to avoid loops.
- No file deletion, modification, or upload.

## Quick start

1. Install **Python 3.10+** if you haven't already.
2. Download the repository using **Code → Download ZIP**, then extract it.
3. Double-click `Start_Hover_Size.bat`.
4. On the first run, the launcher creates a local `.venv` and installs dependencies using pip.
5. Open a regular folder in Windows File Explorer, then hold the pointer over a file or folder for roughly **0.5 seconds**.
6. Keep the command window open while using the tool. Close it or press **Ctrl+C** to stop.

Dependencies are listed in [requirements.txt](requirements.txt).

Alternatively:

```powershell
py -3 -m venv .venv
.\.venv\Scripts\python.exe -m pip install -r requirements.txt
.\.venv\Scripts\python.exe hover_size.py
```

## Supported behavior and limitations

- Supports real file-system items in ordinary File Explorer windows. Virtual locations such as **This PC**, taskbar/Start items, and desktop icons are not supported.
- Some Windows 11 Explorer tab layouts and configurations may not expose the hovered item through UI Automation.
- Calculates **logical file size** (similar to Properties → Size), not **Size on disk**.
- Skips symlinks and junctions, and may undercount unreadable folders. The tooltip reports skipped entries.
- Large folder calculations can take time; hovering elsewhere cancels an in-progress scan.
- The cache may be briefly out of date after file changes (up to 2 minutes).
- No administrator permission is normally required.
- Tested for Python syntax and basic code review, but Windows hover behavior has **not been verified on every Windows 11 build**.

## Privacy

The program scans local file metadata to calculate sizes. It does **not** delete or modify files and does **not** send file names, content, or sizes to any server. Initial setup requires internet access only to install the Python packages `uiautomation` and `pywin32`.

## Project files

- `hover_size.py`: File Explorer hover detection, tooltip UI, and background folder scanner.
- `Start_Hover_Size.bat`: Windows quick-start launcher with virtual environment setup.
- `requirements.txt`: Python dependencies.

## Troubleshooting

If hovering does not show a tooltip, try hovering over an item inside a normal folder instead of the desktop or “This PC.” Leave the command window visible and inspect it for setup errors. Some Explorer configurations may not work with UI Automation.

