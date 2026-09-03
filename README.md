# czUI
**Modern UI Control for Visual Basic 6 — Single-file, GDI+ Powered**

> Drop-in UserControl that transforms VB6 forms into modern, dark-themed applications with anti-aliased rendering, smooth animations, and Windows 11-style chrome.

## Features

- **Single file** — just one .ctl file, no dependencies, no OCX registration
- **GDI+ rendering** — anti-aliased text, rounded corners, smooth gradients
- **7 control types** in one UserControl:

| ControlType | Description |
|---|---|
| `czTitleBar` | Custom window chrome with min/max/close/fullscreen buttons, drag-to-move, DWM rounded corners |
| `czButton` | Rounded button with hover/press states, primary/secondary/danger styles |
| `czLabel` | Anti-aliased label with alignment options |
| `czPanel` | Card container with optional header, border, expandable side panel |
| `czTextBox` | Styled text input with placeholder text |
| `czToggle` | iOS-style toggle switch with 60fps ease-in-out animation |
| `czProgressBar` | Rounded progress bar with percentage text |

## Quick Start

1. Open your VB6 project
2. **Project** > **Add User Control** > **Existing** > select `czUI.ctl`
3. Drop `czControl` instances on your form
4. Set `ControlType` property in the Property Window
5. Configure colors, fonts, and behavior

## Screenshots

*Coming soon*

## Theme Colors (Dark)

| Element | Color | Hex |
|---|---|---|
| Form Background | Dark brown | `#1B2838` |
| Panel | Slightly lighter | `#243447` |
| Border | Subtle edge | `#2C4158` |
| Accent (Amber) | Primary action | `#F5A623` |
| Text | White | `#FFFFFF` |
| Secondary text | Gray | `#8899AA` |

## Custom TitleBar Features

- Windows 11-style glyphs (pixel-perfect)
- Fullscreen mode with auto-show on mouse hover
- Double-click to maximize/restore
- Drag to move, gear icon for settings
- Rounded corners via DWM API

## Toggle Switch Animation

- 60fps timer-based animation
- Ease-in-out cubic easing
- Color interpolation (track blends between on/off colors)
- Hover zoom effect on thumb (+2px)

## Requirements

- Visual Basic 6.0 SP6
- Windows XP or later (GDI+ is built-in since XP)
- No external dependencies

## Version

**v0.1** — Initial release

## License

MIT License — see [LICENSE](LICENSE)

## Author

**Abu Dzakiyyah aka Cyberzilla**