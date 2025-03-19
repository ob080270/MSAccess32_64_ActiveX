# MSAccess32_64_ActiveX

## Overview
This project provides a solution for handling **ActiveX control (`mscomctl.ocx`) compatibility issues** in Microsoft Access when switching between 32-bit and 64-bit versions. The issue primarily affects ActiveX controls such as:

- **TreeView**
- **ListView**
- **ImageList**
- **ProgressBar**
- **Slider (TrackBar)**
- **StatusBar**

Since messages in **Microsoft Access** do not allow users to copy text, a clipboard function has been implemented to facilitate copying and pasting the warning message into a text editor.

## Features
- **Detects the Access bitness (32-bit or 64-bit)** and warns the user, if 32-bit.
- **Provides step-by-step instructions** to resolve ActiveX control issues in 32-bit Access.
- **Allows copying warning messages to the clipboard** for easy pasting into a text editor (Word, Notepad, etc.).

## Screenshots
### 1. Warning Message in 32-bit Access
![32-bit Access Detected](1_Access32_Detected.jpg)

### 2. Clipboard Confirmation Message
![Clipboard Message](2_InsToClipBrdMsg.jpg)

### 3. Pasting Instructions from Clipboard to Word
![Clipboard to Word](3_ClipBrdToEditor.jpg)

## How It Works
### 1. Check Access Bitness
The **`CheckAccess.bas`** module determines whether the user is running **32-bit or 64-bit Access**. If 32-bit is detected, a warning message is displayed.

### 2. Copy Message to Clipboard
The **`CClipboard.cls`** module enables copying the warning message to the clipboard. This allows users to paste it into any text editor for easy reference.

## Installation & Usage
1. **Import the provided VBA modules** (`CheckAccess.bas`, `CClipboard.cls`) into your Access application.
2. **Call the `CheckAccessBitness()` function** at startup to detect the Access version and display warnings if necessary.
3. **Copy the warning message** to the clipboard by clicking the provided button.
4. **Paste the message into a text editor** for easier access to instructions and file paths.

## Troubleshooting
### If `mscomctl.ocx` is missing or not registered
1. Ensure the file **`mscomctl.ocx`** is located at:
   - `C:\WINDOWS\SysWow64\mscomctl.ocx` (for 32-bit systems)
   - `C:\WINDOWS\System32\mscomctl.ocx` (for 64-bit systems)
2. Open **Command Prompt (CMD) as Administrator**.
3. Run the following command:
   ```cmd
   regsvr32 "C:\WINDOWS\SysWow64\mscomctl.ocx"
   ```
4. Restart **Microsoft Access**.

## License
This project is licensed under the **MIT License**.

## Author
[ob080270](https://github.com/ob080270)


