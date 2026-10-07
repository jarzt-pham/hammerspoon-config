# hammerspoon-config

Personal [Hammerspoon](https://www.hammerspoon.org/) config (`~/.hammerspoon`).

| Hotkey | Action |
|---|---|
| `⌃⌘⇧T` → `⌃⌘⇧S` | Switch Microsoft Teams between the two orgs in `ORGS` |
| `⌃⌘⇧R` → `⌃⌘⇧R` | Move windows to their desktop and tile them (see `PLACE` in `init.lua`) |
| `⌃⌘⇧L` → `⌃⌘⇧L` | Open the apps in `PLACE` that have no window (e.g. after a restart), then arrange |

## Setup

```sh
git clone --recurse-submodules git@github.com:jarzt-pham/hammerspoon-config.git ~/.hammerspoon
make -C ~/.hammerspoon/Spoons/SpaceMover.spoon   # native helper for moving windows between desktops
```

macOS settings the desktop layout relies on:

- Hammerspoon has Accessibility permission.
- 3 desktops; Keyboard Shortcuts → Mission Control → "Switch to Desktop 1–3" = `⌃⌥⌘1–3`.
- Mission Control → "Automatically rearrange Spaces based on most recent use" is off.
