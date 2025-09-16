# Excel Mac KeyTips (Hammerspoon)

This project brings Windows-style **KeyTips** to Microsoft Excel on macOS using [Hammerspoon](https://www.hammerspoon.org/).  
It overlays clean visual labels for the **Borders**, **Format**, and **Freeze** dropdown menus and allows activating options directly from the keyboard.

## Features
- Custom visual KeyTips for Excel menus (`Borders`, `Format`, `Freeze`).
- Matches native KeyTips appearance (dark gray pill, white text).
- Automatically appears when menus are opened (`Option→H→B`, `Option→H→O`, `Option→W→F`).
- Always-on with automatic re-arming.
- Efficient and lightweight, safe to keep running in the background.
- Emergency self-disabling if Excel quits unexpectedly.
- NEW: Added support for the insert and delete menus. 

## Usage
1. Install [Hammerspoon](https://www.hammerspoon.org/).
2. Copy the provided `init.lua` into your `~/.hammerspoon/` directory.
3. Reload Hammerspoon.
4. Open Excel and use the standard macOS equivalents:
   - `Option→H→B` → Borders (with custom KeyTips)
   - `Option→H→O` → Format (with custom KeyTips)
   - `Option→W→F` → Freeze (with custom KeyTips)

## Notes
- Designed for Microsoft Excel 365/2024 on macOS.
- Minor Excel updates should not affect functionality.

## License
MIT License

