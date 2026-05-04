# 🖱️ Mouse Position Macro

A lightweight macro tool that automates mouse clicks and movement with customizable keybinds, hardware-level input support, and full configuration support.

## 🔑 Key Features

- **Move Mouse to Custom Positions:** Instantly move your mouse to predefined screen coordinates.
- **Simulate Mouse Movement:** Optionally send movement through `SendInput` at the hardware level and bypasses game restrictions that ignore standard cursor repositioning.
- **Hardware Input for Clicks:** All clicks (left, right, middle) are always sent via `SendInput`, making them compatible with games and applications that block simulated input.
- **Click at Cursor:** Click at your current cursor position without moving to a predefined coordinate.
- **Bindable Keys:** Assign hotkeys to set positions and start/stop the macro.
- **Save & Load Configurations:** Store your macros in configuration files for easy reuse.
- **Auto-Load on Start:** Automatically load your preferred config when the program starts.

Perfect for repetitive tasks, testing, or gaming automation.

---

## ⚠️ Why my executable might get flagged by antivirus software

The executable is built using PyInstaller, which bundles the Python script into a single file. Many antivirus programs flag PyInstaller-packed executables as suspicious because similar techniques are commonly used by malware authors to obfuscate code.

This is a **false positive** and does not mean the program is malicious. The source code is open, clean, and safe to review.

If you have concerns, you can review the source code or run the script directly using Python instead of the compiled executable.

---

📄 This project is licensed under the [Creative Commons Attribution 4.0 International License](https://creativecommons.org/licenses/by/4.0/) — attribution required to **RegaMega**.
