# hammerspoon-config

Personal [Hammerspoon](https://www.hammerspoon.org/) config (`~/.hammerspoon`).

| Hotkey | Action |
|---|---|
| `fn+⌘T` → `fn+⌘S` | Switch Microsoft Teams between the two orgs in `ORGS` |
| `fn+⌘R` → `fn+⌘R` | Move windows to their desktop and tile them (see `PLACE` in `init.lua`) |

## Setup

```sh
git clone --recurse-submodules git@github.com:jarzt-pham/hammerspoon-config.git ~/.hammerspoon
make -C ~/.hammerspoon/Spoons/SpaceMover.spoon   # native helper for moving windows between desktops
```

macOS settings the desktop layout relies on:

- Hammerspoon has Accessibility permission.
- 3 desktops; Keyboard Shortcuts → Mission Control → "Switch to Desktop 1–3" = `⌃⌥⌘1–3`.
- Mission Control → "Automatically rearrange Spaces based on most recent use" is off.
