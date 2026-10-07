require("hs.ipc")

-- Timers nobody references can be garbage-collected before they fire, silently stopping a
-- multi-step sequence. Keep every pending timer alive until it is done.
local liveTimers = {}
local function doAfter(sec, fn)
  local t
  t = hs.timer.doAfter(sec, function() liveTimers[t] = nil; fn() end)
  liveTimers[t] = true
  return t
end
local function doUntil(pred, fn, interval)
  local t
  t = hs.timer.doUntil(function()
    local stop = pred()
    if stop then liveTimers[t] = nil end
    return stop
  end, fn, interval)
  liveTimers[t] = true
  return t
end

-- Teams: ctrl+cmd+shift+t, then ctrl+cmd+shift+s  ->  switch to the other org
-- Deep links (msteams:...?tenantId=) don't switch org in new Teams, so drive the UI:
-- avatar button -> click the other org's entry in the profile popup.
local ORGS = { "JTL-Software-GmbH", "Coduct Solutions GmbH" }
local TEAMS = "com.microsoft.teams2"
local ax = require("hs.axuielement")

local function text(e)
  return (e:attributeValue("AXTitle") or "") .. " " .. (e:attributeValue("AXDescription") or "")
end

-- Depth-first search; `inside(frame)` prunes subtrees outside the region of interest.
local function find(root, pred, inside, limit)
  local n = 0
  local function walk(e, d)
    n = n + 1
    if n > limit or d > 50 then return end
    local f = e:attributeValue("AXFrame")
    if f and f.h > 0 and not inside(f) then return end
    if pred(e) then return e end
    for _, c in ipairs(e:attributeValue("AXChildren") or {}) do
      local r = walk(c, d + 1)
      if r then return r end
    end
  end
  return walk(root, 0)
end

local function mainWindow(appEl)
  for _, w in ipairs(appEl:attributeValue("AXWindows") or {}) do
    if (w:attributeValue("AXTitle") or ""):find("Microsoft Teams") then return w end
  end
end

local function click(e)
  local f = e:attributeValue("AXFrame")
  hs.eventtap.leftClick({ x = f.x + f.w / 2, y = f.y + f.h / 2 })
end

local function switchTeamsOrg()
  local app = hs.application.get(TEAMS)
  if not app then return hs.alert.show("Teams is not running") end
  app:activate()
  local appEl = ax.applicationElement(app)
  appEl:setTimeout(2)
  local win = mainWindow(appEl)
  if not win then return hs.alert.show("Teams: main window not found") end

  local title = win:attributeValue("AXTitle")
  local target = title:find(ORGS[1], 1, true) and ORGS[2] or ORGS[1]
  local wf = win:attributeValue("AXFrame")

  local avatar = find(win,
    function(e) return (e:attributeValue("AXDescription") or ""):find("^Your profile") end,
    function(f) return f.y < wf.y + 110 end, 3000)
  if not avatar then return hs.alert.show("Teams: avatar button not found") end
  click(avatar)

  -- Profile popup sits under the avatar at the window's right edge; poll until the org entry shows.
  local tries = 0
  doUntil(function() return tries >= 10 end, function()
    tries = tries + 1
    local entry = find(win,
      function(e) return text(e):find(target, 1, true) and e ~= avatar end,
      function(f) return f.x + f.w > wf.x + wf.w - 600 end, 6000)
    if entry then
      tries = 10
      click(entry)
      hs.alert.show("Teams → " .. target)
    elseif tries >= 10 then
      hs.alert.show("Teams: '" .. target .. "' not found in profile menu")
    end
  end, 0.4)
end
TeamsSwitchOrg = switchTeamsOrg -- exposed for `hs -c "TeamsSwitchOrg()"`

-- Desktop layout. macOS 15 ignores hs.spaces.moveWindowToSpace, so windows are moved with
-- SpaceMover.spoon's native helper; desktops are switched with ⌃⌥⌘1-3 (Mission Control
-- shortcuts, "Automatically rearrange Spaces" off so numbering stays fixed).
-- App -> { desktop, slot }. Apps not listed go to desktop 3 untouched in size.
local SLOTS = {
  left        = { x = 0,   y = 0,   w = 0.5, h = 1 },
  topRight    = { x = 0.5, y = 0,   w = 0.5, h = 0.5 },
  bottomRight = { x = 0.5, y = 0.5, w = 0.5, h = 0.5 },
}
local PLACE = {
  ["company.thebrowser.Browser"] = { 1, "left" },        -- Arc
  ["com.tinyspeck.slackmacgap"]  = { 1, "topRight" },    -- Slack + Teams stacked
  ["com.microsoft.teams2"]       = { 1, "topRight" },
  ["com.openai.codex"]           = { 1, "bottomRight" }, -- ChatGPT + Obsidian + Claude stacked
  ["md.obsidian"]                = { 1, "bottomRight" },
  ["com.anthropic.claudefordesktop"] = { 1, "bottomRight" },
  ["com.microsoft.VSCode"]       = { 2, "left" },
  ["org.alacritty"]              = { 2, "alternate" },   -- top-right, bottom-right, top-right, …
}
local OTHER_DESKTOP = 3

hs.loadSpoon("SpaceMover")
local MOVER = hs.configdir .. "/Spoons/SpaceMover.spoon/native/space-mover"

local function gotoDesktop(n, desktops, then_)
  hs.eventtap.keyStroke({ "ctrl", "alt", "cmd" }, tostring(n), 0)
  local waited = 0
  doUntil(function() return waited < 0 end, function()
    waited = waited + 0.1
    if hs.spaces.focusedSpace() == desktops[n] or waited > 2 then
      waited = -1
      doAfter(0.4, then_) -- let the switch animation finish before querying windows
    end
  end, 0.1)
end

-- Visit each desktop, push misplaced windows to their desktop and resize the ones in a slot.
-- Moved windows are resized right away (AX frames work off-screen), so one visit per desktop is enough.
local function arrangeDesktops()
  local desktops = spoon.SpaceMover:desktopSpaces()
  if not desktops or #desktops < OTHER_DESKTOP then
    return hs.alert.show("Need " .. OTHER_DESKTOP .. " desktops, found " .. (desktops and #desktops or 0))
  end
  local startDesktop = hs.fnutils.indexOf(desktops, hs.spaces.focusedSpace()) or 1
  local startWin = hs.window.focusedWindow()
  local frame = hs.screen.mainScreen():frame()
  -- Rectangle's "Gaps between windows" setting, so tiles look the same as Rectangle's own.
  local gap = tonumber((hs.execute("defaults read com.knollsoft.Rectangle gapSize 2>/dev/null"))) or 0
  local alternate, failed, done = 0, 0, {}

  local function place(win, bundleID, d)
    if done[win:id()] then return end -- already moved here from an earlier desktop
    done[win:id()] = true
    local p = PLACE[bundleID] or { OTHER_DESKTOP }
    if p[1] ~= d then
      local _, ok = hs.execute(MOVER .. " " .. win:id() .. " " .. desktops[p[1]])
      if not ok then failed = failed + 1 return end
    end
    local slot = p[2]
    if slot == "alternate" then
      alternate = alternate + 1
      slot = alternate % 2 == 1 and "topRight" or "bottomRight"
    end
    local u = SLOTS[slot]
    if u then
      -- Like Rectangle: full gap at screen edges, half a gap on each side of a shared edge,
      -- so neighbouring tiles end up exactly one gap apart.
      local function inset(start, size)
        return (start == 0 and gap or gap / 2), (start + size >= 1 and gap or gap / 2)
      end
      local l, r = inset(u.x, u.w)
      local t, b = inset(u.y, u.h)
      win:setFrame({
        x = frame.x + u.x * frame.w + l,
        y = frame.y + u.y * frame.h + t,
        w = u.w * frame.w - l - r,
        h = u.h * frame.h - t - b,
      }, 0)
    end
  end

  local d = 0
  local function visitNext()
    d = d + 1
    if d > OTHER_DESKTOP then
      return gotoDesktop(startDesktop, desktops, function()
        if startWin then startWin:focus() end
        hs.alert.show(failed == 0 and "Desktops arranged" or ("Arranged, " .. failed .. " window(s) failed to move"))
      end)
    end
    gotoDesktop(d, desktops, function()
      local here = desktops[d]
      -- Resolve each window's app once: the pid -> app lookup fails intermittently (LuaSkin
      -- "Unable to fetch NSRunningApplication"), and a second lookup returning nil used to
      -- throw and abort the whole sequence. Unresolvable windows are skipped.
      local wins = {}
      for _, w in ipairs(hs.window.allWindows()) do
        local s, app = hs.spaces.windowSpaces(w), w:application()
        if app and w:isStandard() and not w:isFullScreen() and s and #s == 1 and s[1] == here then
          table.insert(wins, { win = w, bundleID = app:bundleID() })
        end
      end
      table.sort(wins, function(a, b) return a.win:id() < b.win:id() end) -- stable Alacritty top/bottom order
      for _, e in ipairs(wins) do place(e.win, e.bundleID, d) end
      visitNext()
    end)
  end
  visitNext()
end
ArrangeDesktops = arrangeDesktops -- exposed for `hs -c "ArrangeDesktops()"`

-- After a restart (or after closing windows): make sure every app in PLACE has a window,
-- then arrange. Closing an app's last window often leaves it running (Arc, Slack…), so
-- "has a window" is what counts, not "is running".
local ALACRITTY_WINDOWS = 2

-- Neither app:allWindows() nor the AX window list sees windows on other desktops, so count
-- WindowServer windows (CGWindowList, via JXA) that sit on some desktop. The desktop check
-- drops windows an app keeps hidden after you close them.
local CG_WINDOWS_JXA = [[osascript -l JavaScript -e 'ObjC.import("CoreGraphics");
JSON.stringify(ObjC.deepUnwrap(ObjC.castRefToObject($.CGWindowListCopyWindowInfo($.kCGWindowListOptionAll, 0)))
  .filter(w => w.kCGWindowLayer === 0 && w.kCGWindowBounds.Height > 100)
  .map(w => [w.kCGWindowNumber, w.kCGWindowOwnerPID]))']]

local function windowCountsByPid()
  local onDesktop = {}
  for _, space in ipairs(spoon.SpaceMover:desktopSpaces() or {}) do
    for _, id in ipairs(hs.spaces.windowsForSpace(space) or {}) do onDesktop[id] = true end
  end
  local counts = {}
  for _, w in ipairs(hs.json.decode((hs.execute(CG_WINDOWS_JXA))) or {}) do
    if onDesktop[w[1]] then counts[w[2]] = (counts[w[2]] or 0) + 1 end
  end
  return counts
end

local function windowCount(bundleID, counts)
  local app = hs.application.get(bundleID)
  return app and (counts or windowCountsByPid())[app:pid()] or 0
end

local function launchWorkspace()
  local opened = {}
  local counts = windowCountsByPid()
  for bundleID in pairs(PLACE) do
    if windowCount(bundleID, counts) == 0 then
      hs.execute("open -b " .. bundleID) -- launches, or "reopens" a running app into a new window
      opened[bundleID] = true
    end
  end
  local topUpAlacritty = windowCount("org.alacritty", counts) < ALACRITTY_WINDOWS
  if not next(opened) and not topUpAlacritty then return arrangeDesktops() end
  hs.alert.show("Opening apps…")
  local waited = 0
  doUntil(function() return waited < 0 end, function()
    waited = waited + 1
    local now, pending = windowCountsByPid(), {}
    for bundleID in pairs(opened) do
      if windowCount(bundleID, now) == 0 then table.insert(pending, bundleID) end
    end
    local have = windowCount("org.alacritty", now)
    if topUpAlacritty and have > 0 then
      topUpAlacritty = false
      for _ = have + 1, ALACRITTY_WINDOWS do
        hs.execute("/Applications/Alacritty.app/Contents/MacOS/alacritty msg create-window")
      end
    end
    if #pending == 0 or waited > 60 then
      waited = -1
      if #pending > 0 then hs.alert.show("No window yet: " .. table.concat(pending, ", ")) end
      doAfter(1, arrangeDesktops) -- let the last windows settle
    end
  end, 1)
end
LaunchWorkspace = launchWorkspace -- exposed for `hs -c "LaunchWorkspace()"`

-- Chords: ctrl+cmd+shift+<first>, then ctrl+cmd+shift+<second> within 2s. (Not fn: external
-- keyboards handle fn in firmware and never send it to macOS.) Other combos pass through untouched.
local KC = hs.keycodes.map
local CHORDS = {
  { first = KC.t, second = KC.s, action = switchTeamsOrg },
  { first = KC.r, second = KC.r, action = arrangeDesktops },
  { first = KC.l, second = KC.l, action = launchWorkspace },
}
local armed, armedAt = nil, 0
ChordTap = hs.eventtap.new({ hs.eventtap.event.types.keyDown }, function(ev)
  local fl = ev:getFlags()
  if not (fl.ctrl and fl.cmd and fl.shift and not fl.alt) then return false end
  local key = ev:getKeyCode()
  if armed and key == armed.second and hs.timer.secondsSinceEpoch() - armedAt < 2 then
    local action = armed.action
    armed = nil
    doAfter(0, action)
    return true
  end
  for _, c in ipairs(CHORDS) do
    if key == c.first then
      armed, armedAt = c, hs.timer.secondsSinceEpoch()
      return true
    end
  end
  return false
end):start()
