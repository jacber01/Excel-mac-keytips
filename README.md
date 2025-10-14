# Excel Mac KeyTips (Hammerspoon)

This project brings Windows-style **KeyTips** to Microsoft Excel on macOS using [Hammerspoon](https://www.hammerspoon.org/).  
It overlays clean visual labels for the **Borders**, **Format**, **Freeze**, **Insert** and **Delete** dropdown menus and allows activating options directly from the keyboard.

## Features
- Custom visual KeyTips for Excel menus (`Borders`, `Format`, `Freeze`,`Insert`, `Delete` & `More...`).
- Matches native KeyTips appearance (dark gray pill, white text).
- Automatically appears when menus are opened (`Option→H→B`, `Option→H→O`, `Option→W→F`, `Option→H→I`, `Option→H→D`).
- Efficient and lightweight, safe to keep running in the background.
- Emergency self-disabling if Excel quits unexpectedly.
- NEW: Added support for the insert and delete menus.
- NEW: Added support for Merge, Paste, Copy, Group, Ungroup, and More... menus.

## Usage
1. Install [Hammerspoon](https://www.hammerspoon.org/).
2. Copy the provided `init.lua` into your `~/.hammerspoon/` directory. (Open Finder. Press ⌘ + Shift + G. In the dialog, type ~/.hammerspoon)
3. Reload Hammerspoon.
4. Open Excel and use the standard macOS equivalents:
   - `Option→H→B` → Borders (with custom KeyTips)
   - `Option→H→O` → Format (with custom KeyTips)
   - `Option→W→F` → Freeze (with custom KeyTips)
   - `Option→H→I` → Insert (with custom KeyTips)
   - `Option→H→D` → Delete (with custom KeyTips)

## Notes
- Designed for Microsoft Excel 365/2024 on macOS in Dark Mode.
- Minor Excel updates should not affect functionality.
- You have to enable Keytips in Preferences→Accessibility
- You have to set your Excel to English

## License
MIT License

