local app = require("hs.application")
local ax = require("hs.axuielement")
local eventtap = require("hs.eventtap")
local canvas = require("hs.canvas")
local screen = require("hs.screen")
local timer = require("hs.timer")
local keycodes = require("hs.keycodes")

local EXCEL_BUNDLE = "com.microsoft.Excel"

-- Borders
local LETTER_MAP_BORDERS = {
B = { "Bottom Border" },
T = { "Top Border" },
L = { "Left Border" },
R = { "Right Border" },
N = { "No Border" },
A = { "All Borders" },
O = { "Outside Borders" },
H = { "Thick Box Border" },
M = { "More Borders" },
}

-- Format
local LETTER_MAP_FORMAT = {
H = { "Row Height" },
A = { "AutoFit Row Height" },
W = { "Column Width" },
I = { "AutoFit Column Width" },
D = { "Default Width" },
R = { "Rename Sheet" },
M = { "Move or Copy Sheet" },
P = { "Protect Sheet" },
L = { "Lock Cell" },
F = { "Format Cells" },
}

-- Freeze
local LETTER_MAP_FREEZE = {
F = { "Freeze Panes" },
R = { "Freeze Top Row" },
C = { "Freeze First Column" },
}

-- ---------- helpers ----------
local function excelApp()
  return app.get(EXCEL_BUNDLE) or app.appFromBundleID(EXCEL_BUNDLE)
end

local function excelRunning()
  local e = excelApp()
  return e and e:isRunning()
end

local function isExcelFrontmost()
  local e = excelApp()
  return e and e:isFrontmost()
end

local function safeAttr(el, attr)
  local ok, val = pcall(function() return el:attributeValue(attr) end)
  if ok then return val end
end

local function getFrame(el)
  if not el then return nil end
  local f = safeAttr(el, "AXFrame")
  if f and f.w and f.h then return f end
  local p = safeAttr(el, "AXPosition")
  local s = safeAttr(el, "AXSize")
  if p and s then return { x = p.x, y = p.y, w = s.w, h = s.h } end
end

-- Pick screen with largest intersection
local function screenForRectCompat(rect)
  local bestScr, bestArea = nil, -1
  for _, scr in ipairs(screen.allScreens()) do
    local sf = scr:frame()
    local ix1 = math.max(sf.x, rect.x)
    local iy1 = math.max(sf.y, rect.y)
    local ix2 = math.min(sf.x + sf.w, rect.x + rect.w)
    local iy2 = math.min(sf.y + sf.h, rect.y + rect.h)
    local w = ix2 - ix1
    local h = iy2 - iy1
    local area = (w > 0 and h > 0) and (w * h) or 0
    if area > bestArea then bestArea, bestScr = area, scr end
  end
  return bestScr or screen.mainScreen()
end

local function visibleOnScreen(frame)
  if not frame or frame.w <= 1 or frame.h <= 1 then return false end
  local scr = screenForRectCompat(frame):frame()
  local s = scr
  return frame.x + frame.w > s.x and frame.x < s.x + s.w
     and frame.y + frame.h > s.y and frame.y < s.y + s.h
end

local function labelOf(el)
  if not el then return "" end
  local t = safeAttr(el, "AXTitle")
  if t and t ~= "" then return t end
  local d = safeAttr(el, "AXDescription")
  if d and d ~= "" then return d end
  local h = safeAttr(el, "AXHelp")
  if h and h ~= "" then return h end
  local v = safeAttr(el, "AXValue")
  if type(v) == "string" and v ~= "" then return v end
  return ""
end

local function labelMatches(letterMap, letter, lbl)
  local pats = letterMap[letter]
  if not pats or not lbl or lbl == "" then return false end
  local lower = lbl:lower()
  for _, pat in ipairs(pats) do
    if lower:find(pat:lower(), 1, true) then return true end
  end
  return false
end

local function anyLabelMatches(letterMap, lbl)
  if not lbl or lbl == "" then return false end
  local lower = lbl:lower()
  for _, pats in pairs(letterMap) do
    for _, pat in ipairs(pats) do
      if lower:find(pat:lower(), 1, true) then return true end
    end
  end
  return false
end

-- ---------- container detection (generic) ----------
local function findMenuContainerForMap(letterMap)
  if not excelRunning() then return nil end
  local e = excelApp()
  if not e then return nil end
  local ok, root = pcall(function() return ax.applicationElement(e) end)
  if not ok or not root then return nil end

  local q, steps = { root }, 0
  while #q > 0 and steps < 1500 do
    steps = steps + 1
    local el = table.remove(q, 1)
    if safeAttr(el, "AXRole") == "AXGroup" then
      local kids = safeAttr(el, "AXChildren")
      local btnCount, matchCount = 0, 0
      if kids then
        for _, ch in ipairs(kids) do
          if safeAttr(ch, "AXRole") == "AXMenuButton" then
            local f = getFrame(ch)
            if visibleOnScreen(f) then
              btnCount = btnCount + 1
              local lbl = labelOf(ch)
              if anyLabelMatches(letterMap, lbl) then
                matchCount = matchCount + 1
              end
            end
          end
        end
        if btnCount >= 5 and matchCount >= 1 then
          return el
        end
      end
      if kids then for _, ch in ipairs(kids) do table.insert(q, ch) end end
    else
      local kids = safeAttr(el, "AXChildren")
      if kids then for _, ch in ipairs(kids) do table.insert(q, ch) end end
    end
  end
  return nil
end

-- Freeze-specific container detection (three options; do not touch the generic one)
local function findFreezeMenuContainer()
  if not excelRunning() then return nil end
  local e = excelApp()
  if not e then return nil end
  local ok, root = pcall(function() return ax.applicationElement(e) end)
  if not ok or not root then return nil end

  local q, steps = { root }, 0
  while #q > 0 and steps < 1500 do
    steps = steps + 1
    local el = table.remove(q, 1)
    if safeAttr(el, "AXRole") == "AXGroup" then
      local kids = safeAttr(el, "AXChildren")
      local btnCount, matchCount = 0, 0
      if kids then
        for _, ch in ipairs(kids) do
          if safeAttr(ch, "AXRole") == "AXMenuButton" then
            local f = getFrame(ch)
            if visibleOnScreen(f) then
              btnCount = btnCount + 1
              local lbl = labelOf(ch)
              if anyLabelMatches(LETTER_MAP_FREEZE, lbl) then
                matchCount = matchCount + 1
              end
            end
          end
        end
        if btnCount >= 3 and matchCount >= 1 then
          return el
        end
      end
      if kids then for _, ch in ipairs(kids) do table.insert(q, ch) end end
    else
      local kids = safeAttr(el, "AXChildren")
      if kids then for _, ch in ipairs(kids) do table.insert(q, ch) end end
    end
  end
  return nil
end

local function collectItems(letterMap)
  local container = findMenuContainerForMap(letterMap)
  if not container then return {} end
  local found = {}

  local function traverse(el, depth)
    depth = depth or 0
    if depth > 6 then return end
    local role = safeAttr(el, "AXRole")
    if role == "AXMenuButton" then
      local lbl = labelOf(el)
      local f = getFrame(el)
      if visibleOnScreen(f) and lbl and lbl ~= "" then
        for letter, _ in pairs(letterMap) do
          if not found[letter] and labelMatches(letterMap, letter, lbl) then
            found[letter] = { el = el, frame = f, label = lbl }
            break
          end
        end
      end
    end
    local kids = safeAttr(el, "AXChildren")
    if kids then for _, ch in ipairs(kids) do traverse(ch, depth + 1) end end
  end

  traverse(container, 0)
  return found
end

-- Freeze-specific collector (uses the freeze container finder)
local function collectFreezeItems()
  local container = findFreezeMenuContainer()
  if not container then return {} end
  local found = {}

  local function traverse(el, depth)
    depth = depth or 0
    if depth > 6 then return end
    local role = safeAttr(el, "AXRole")
    if role == "AXMenuButton" then
      local lbl = labelOf(el)
      local f = getFrame(el)
      if visibleOnScreen(f) and lbl and lbl ~= "" then
        for letter, _ in pairs(LETTER_MAP_FREEZE) do
          if not found[letter] and labelMatches(LETTER_MAP_FREEZE, letter, lbl) then
            found[letter] = { el = el, frame = f, label = lbl }
            break
          end
        end
      end
    end
    local kids = safeAttr(el, "AXChildren")
    if kids then for _, ch in ipairs(kids) do traverse(ch, depth + 1) end end
  end

  traverse(container, 0)
  return found
end

-- ---------- state ----------
local state = {
  active = false,
  tips = {},
  items = {},
  tap = nil,
  context = "none",
  checkTimer = nil,
}
local masterEnabled = true -- always on unless hardStopped

local excelMonitor = nil
local appWatcher = nil
local excelCheckTimer = nil
local bStartTap = nil
local bStartTimer = nil
local oStartTap = nil
local oStartTimer = nil
local fStartTap = nil
local fStartTimer = nil
local toggleTap = nil
local hardStopped = false -- requires HS reload once set

local optionTap = nil
local optionIsDown = false

-- ---------- visuals ----------
local function clearTips()
  for _, c in pairs(state.tips) do
    pcall(function() c:hide(); c:delete() end)
  end
  state.tips = {}
  state.items = {}
  if state.tap then state.tap:stop(); state.tap = nil end
  if state.checkTimer then state.checkTimer:stop(); state.checkTimer = nil end
  state.active = false
  state.context = "none"
end

-- Center the pill within the menu item's row (both horizontally and vertically).
local function showTip(letter, frame)
  local w, h = 24, 20
  local x = frame.x + math.floor((frame.w - w) / 2)
  local y = frame.y + math.max(0, math.floor((frame.h - h) / 2))

  -- Clamp to screen to avoid edge cases
  local scr = screenForRectCompat(frame):frame()
  if x < scr.x then x = scr.x + 2 end
  if x + w > scr.x + scr.w then x = scr.x + scr.w - w - 2 end
  if y < scr.y then y = scr.y + 2 end

  local c = canvas.new({ x = x, y = y, w = w, h = h })
  c:appendElements(
    { type = "rectangle", action = "fill",
      roundedRectRadii = { xRadius = 6, yRadius = 6 },
      fillColor = { red = 0.28, green = 0.28, blue = 0.30, alpha = 1 },
      strokeColor = { white = 0, alpha = 0 }, strokeWidth = 0 },
    { type = "text", text = letter, textSize = 13,
      textColor = { red = 1, green = 1, blue = 1, alpha = 1 },
      frame = { x = 0, y = 0, w = w, h = h },
      textAlignment = "center" }
  )
  c:level(canvas.windowLevels.overlay)
  c:show()
  state.tips[letter] = c
end

-- ---------- activation ----------
local function activateKeytips()
  if hardStopped or not masterEnabled then return end
  clearTips()
  if not isExcelFrontmost() or not excelRunning() then return end

  -- Prefer Borders; if not found, fall back to Format; then Freeze.
  local items = collectItems(LETTER_MAP_BORDERS)
  local which = "borders"
  if not items or next(items) == nil then
    items = collectItems(LETTER_MAP_FORMAT)
    which = (items and next(items) ~= nil) and "format" or nil
  end
  if not which then
    items = collectFreezeItems()
    which = (items and next(items) ~= nil) and "freeze" or "none"
  end
  if which == "none" then return end

  state.items = items
  state.context = which
  for letter, info in pairs(items) do
    showTip(letter, info.frame)
  end

  state.tap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
    if hardStopped then return false end
    if not state.active then return false end
    local char = ev:getCharacters()
    if not char or char == "" then return true end
    local letter = string.upper(char)
    local target = state.items[letter]
    if target and target.el then
      pcall(function() target.el:performAction("AXPress") end)
      clearTips()
      timer.doAfter(0.12, function()
        if not hardStopped and masterEnabled then activateKeytips() end
      end)
      return true
    end
    -- Any other key hides tips briefly; only reactivates if still enabled.
    clearTips()
    timer.doAfter(0.12, function()
      if not hardStopped and masterEnabled then activateKeytips() end
    end)
    return true
  end)
  state.active = true
  state.tap:start()

  state.checkTimer = timer.doEvery(0.05, function()
    local container
    if state.context == "borders" then
      container = findMenuContainerForMap(LETTER_MAP_BORDERS)
    elseif state.context == "format" then
      container = findMenuContainerForMap(LETTER_MAP_FORMAT)
    elseif state.context == "freeze" then
      container = findFreezeMenuContainer()
    end
    if not container then
      clearTips()
    end
  end)
  state.checkTimer:start()
end

-- Non-consuming triggers: 'B' (Borders), 'O' (Format), 'F' (Freeze)
bStartTap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
  if hardStopped or not masterEnabled then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  if ev:getKeyCode() ~= keycodes.map["b"] then return false end
  if bStartTimer then bStartTimer:stop(); bStartTimer = nil end
  bStartTimer = timer.doAfter(0.08, function()
    bStartTimer = nil
    if not hardStopped and masterEnabled then activateKeytips() end
  end)
  return false
end)
bStartTap:start()

oStartTap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
  if hardStopped or not masterEnabled then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  if ev:getKeyCode() ~= keycodes.map["o"] then return false end
  if oStartTimer then oStartTimer:stop(); oStartTimer = nil end
  oStartTimer = timer.doAfter(0.08, function()
    oStartTimer = nil
    if not hardStopped and masterEnabled then activateKeytips() end
  end)
  return false
end)
oStartTap:start()

fStartTap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
  if hardStopped or not masterEnabled then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  if ev:getKeyCode() ~= keycodes.map["f"] then return false end
  if fStartTimer then fStartTimer:stop(); fStartTimer = nil end
  fStartTimer = timer.doAfter(0.08, function()
    fStartTimer = nil
    if not hardStopped and masterEnabled then activateKeytips() end
  end)
  return false
end)
fStartTap:start()

-- Excel activity monitor with debouncing
excelMonitor = eventtap.new(
  { eventtap.event.types.keyDown, eventtap.event.types.leftMouseUp },
  function(ev)
    if hardStopped or not masterEnabled then return false end
    if isExcelFrontmost() and excelRunning() then
      if excelCheckTimer then excelCheckTimer:stop() end
      excelCheckTimer = timer.doAfter(0.35, function()
        excelCheckTimer = nil
        if not hardStopped and masterEnabled then activateKeytips() end
      end)
    end
    return false
  end
)

-- Helper: full teardown and DISABLE (no app quit; requires HS reload)
local function teardownAndDisable()
  if hardStopped then return end
  hardStopped = true
  masterEnabled = false
  clearTips()
  if excelMonitor then excelMonitor:stop(); excelMonitor = nil end
  if bStartTap then bStartTap:stop(); bStartTap = nil end
  if oStartTap then oStartTap:stop(); oStartTap = nil end
  if fStartTap then fStartTap:stop(); fStartTap = nil end
  if toggleTap then toggleTap:stop(); toggleTap = nil end
  if optionTap then optionTap:stop(); optionTap = nil end
  if excelCheckTimer then excelCheckTimer:stop(); excelCheckTimer = nil end
  if bStartTimer then bStartTimer:stop(); bStartTimer = nil end
  if oStartTimer then oStartTimer:stop(); oStartTimer = nil end
  if fStartTimer then fStartTimer:stop(); fStartTimer = nil end
  if appWatcher then appWatcher:stop(); appWatcher = nil end
  hs.alert.show("Excel KeyTips: disabled (reload Hammerspoon to re-enable)")
end

-- App watcher
appWatcher = app.watcher.new(function(appName, eventType, appObj)
  if appName == "Microsoft Excel" then
    if eventType == app.watcher.activated then
      if hardStopped then return end
      if excelMonitor then excelMonitor:start() end
      timer.doAfter(0.15, function() if not hardStopped then activateKeytips() end end)
    elseif eventType == app.watcher.deactivated then
      if hardStopped then return end
      if excelMonitor then excelMonitor:stop() end
      if excelCheckTimer then excelCheckTimer:stop(); excelCheckTimer = nil end
      clearTips()
    elseif eventType == app.watcher.terminated then
      teardownAndDisable()
    end
  end
end)
appWatcher:start()

toggleTap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
  if hardStopped then return false end
  local kc = ev:getKeyCode()
  if kc ~= 53 then return false end -- ESC only
  if not isExcelFrontmost() or not excelRunning() then return false end
  if not state.active then return false end -- Only consume ESC when tips are shown
  -- One-shot OFF: hide and keep disabled until Option pressed again
  masterEnabled = false
  clearTips()
  return true
end)
toggleTap:start()

optionTap = eventtap.new({ eventtap.event.types.flagsChanged }, function(ev)
  if hardStopped then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end

  local flags = ev:getFlags()
  local nowDown = flags.alt or flags.altgr

  if nowDown ~= optionIsDown then
    optionIsDown = nowDown
    if nowDown then
      -- Only process Option toggle when tips are shown OR when disabled by ESC
      if state.active or not masterEnabled then
        if masterEnabled then
          masterEnabled = false
          clearTips()
        else
          masterEnabled = true
          activateKeytips()
        end
        return true
      end
    end
  end
  return false
end)
optionTap:start()
