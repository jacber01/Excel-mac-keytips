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
D = { "Bottom Double Border" },
K = { "Thick Bottom Border" },
P = { "Top and Bottom Border" },
U = { "Top and Thick Bottom Border" },
E = { "Top and Double Bottom Border" },
G = { "Draw Border Grid" },
S = { "Erase Border" },
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

-- Insert
local LETTER_MAP_INSERT = {
I = { "Insert Cells" },
R = { "Insert Sheet Rows" },
C = { "Insert Sheet Columns" },
S = { "Insert Sheet" },
}

-- Delete
local LETTER_MAP_DELETE = {
D = { "Delete Cells" },
R = { "Delete Sheet Rows" },
C = { "Delete Sheet Columns" },
T = { "Delete Table Rows" },
L = { "Delete Table Columns" },
S = { "Delete Sheet" },
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

  -- Disambiguation for Insert/Delete 'S' (Sheet) vs Rows/Columns
  if (letterMap == LETTER_MAP_INSERT or letterMap == LETTER_MAP_DELETE) and letter == "S" then
    local lower2 = lbl:lower()
    if lower2:find("columns", 1, true) or lower2:find("rows", 1, true) then
      return false
    end
  end

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

-- Freeze-specific container detection (three options)
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

-- Insert-specific container detection (4 visible items)
local function findInsertMenuContainer()
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
              if anyLabelMatches(LETTER_MAP_INSERT, lbl) then
                matchCount = matchCount + 1
              end
            end
          end
        end
        if btnCount >= 4 and matchCount >= 1 then
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

-- Freeze-specific collector
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

-- Insert-specific collector
local function collectInsertItems()
  local container = findInsertMenuContainer()
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
        for letter, _ in pairs(LETTER_MAP_INSERT) do
          if not found[letter] and labelMatches(LETTER_MAP_INSERT, letter, lbl) then
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
local iStartTap = nil
local iStartTimer = nil
local dStartTap = nil
local dStartTimer = nil
local clickHideTap = nil
local mouseOpenTap = nil
local mouseDetectTimer = nil
local suppressAutoFromMouse = false
local suppressAutoFromOption = false
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
      fillColor = { red = 0.478, green = 0.478, blue = 0.478, alpha = 1 },
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

  -- Prefer Borders; then Format; then Freeze; then Insert; then Delete.
  local items = collectItems(LETTER_MAP_BORDERS)
  local which = "borders"
  if not items or next(items) == nil then
    items = collectItems(LETTER_MAP_FORMAT)
    which = (items and next(items) ~= nil) and "format" or nil
  end
  if not which then
    items = collectFreezeItems()
    which = (items and next(items) ~= nil) and "freeze" or nil
  end
  if not which then
    items = collectInsertItems()
    which = (items and next(items) ~= nil) and "insert" or nil
  end
  if not which then
    items = collectItems(LETTER_MAP_DELETE)
    which = (items and next(items) ~= nil) and "delete" or "none"
  end
  if which == "none" then return end

  state.items = items
  state.context = which
  for letter, info in pairs(items) do
    showTip(letter, info.frame)
  end

  -- Immediate hide on any left click while a custom menu is active (non-consuming)
  if clickHideTap then clickHideTap:stop(); clickHideTap = nil end
  clickHideTap = eventtap.new({ eventtap.event.types.leftMouseDown }, function()
    if hardStopped or not masterEnabled then return false end
    if not state.active then return false end
    if state.context == "borders" or state.context == "format"
       or state.context == "freeze" or state.context == "insert"
       or state.context == "delete" then
      clearTips()
      timer.doAfter(0.12, function()
        if not hardStopped and masterEnabled then activateKeytips() end
      end)
    end
    return false
  end)
  clickHideTap:start()

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
    elseif state.context == "insert" then
      container = findInsertMenuContainer()
    elseif state.context == "delete" then
      container = findMenuContainerForMap(LETTER_MAP_DELETE)
    end
    if not container then
      clearTips()
    end
  end)
  state.checkTimer:start()
end

-- Non-consuming triggers: 'B', 'O', 'F', 'I', 'D'
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

iStartTap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
  if hardStopped or not masterEnabled then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  if ev:getKeyCode() ~= keycodes.map["i"] then return false end
  if iStartTimer then iStartTimer:stop(); iStartTimer = nil end
  iStartTimer = timer.doAfter(0.08, function()
    iStartTimer = nil
    if not hardStopped and masterEnabled then activateKeytips() end
  end)
  return false
end)
iStartTap:start()

dStartTap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
  if hardStopped or not masterEnabled then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  if ev:getKeyCode() ~= keycodes.map["d"] then return false end
  if dStartTimer then dStartTimer:stop(); dStartTimer = nil end
  dStartTimer = timer.doAfter(0.08, function()
    dStartTimer = nil
    if not hardStopped and masterEnabled then activateKeytips() end
  end)
  return false
end)
dStartTap:start()

-- Global mouse listener to detect mouse-opened dropdowns and suppress auto custom tips
mouseOpenTap = eventtap.new({ eventtap.event.types.leftMouseDown }, function()
  if hardStopped then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  if mouseDetectTimer then mouseDetectTimer:stop(); mouseDetectTimer = nil end
  mouseDetectTimer = timer.doAfter(0.12, function()
    local opened = findMenuContainerForMap(LETTER_MAP_BORDERS)
                or findMenuContainerForMap(LETTER_MAP_FORMAT)
                or findFreezeMenuContainer()
                or findInsertMenuContainer()
                or findMenuContainerForMap(LETTER_MAP_DELETE)
    if opened then suppressAutoFromMouse = true end
  end)
  return false
end)
mouseOpenTap:start()

-- Excel activity monitor with debouncing
excelMonitor = eventtap.new(
  { eventtap.event.types.keyDown, eventtap.event.types.leftMouseUp },
  function(ev)
    if hardStopped or not masterEnabled then return false end
    if isExcelFrontmost() and excelRunning() then
      -- If we're in "restart native" mode, clear it once user starts keying the sequence.
      if ev:getType() == eventtap.event.types.keyDown and suppressAutoFromOption then
        suppressAutoFromOption = false
      end
      if excelCheckTimer then excelCheckTimer:stop() end
      excelCheckTimer = timer.doAfter(0.35, function()
        excelCheckTimer = nil
        -- Suppress our auto-appearance after mouse-open or while restarting native.
        if suppressAutoFromMouse or suppressAutoFromOption then return end
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
  if iStartTap then iStartTap:stop(); iStartTap = nil end
  if dStartTap then dStartTap:stop(); dStartTap = nil end
  if clickHideTap then clickHideTap:stop(); clickHideTap = nil end
  if mouseOpenTap then mouseOpenTap:stop(); mouseOpenTap = nil end
  if mouseDetectTimer then mouseDetectTimer:stop(); mouseDetectTimer = nil end
  if toggleTap then toggleTap:stop(); toggleTap = nil end
  if optionTap then optionTap:stop(); optionTap = nil end
  if excelCheckTimer then excelCheckTimer:stop(); excelCheckTimer = nil end
  if bStartTimer then bStartTimer:stop(); bStartTimer = nil end
  if oStartTimer then oStartTimer:stop(); oStartTimer = nil end
  if fStartTimer then fStartTimer:stop(); fStartTimer = nil end
  if iStartTimer then iStartTimer:stop(); iStartTimer = nil end
  if dStartTimer then dStartTimer:stop(); dStartTimer = nil end
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
      if state.active or not masterEnabled then
        if masterEnabled then
          -- Custom tips visible: hide them and CONSUME Option
          masterEnabled = false
          clearTips()
          suppressAutoFromOption = false
          return true
        else
          -- Re-enable: pass Option through so native KeyTips appear; close any open dropdown first
          masterEnabled = true
          suppressAutoFromMouse = false
          suppressAutoFromOption = true
          local opened = findMenuContainerForMap(LETTER_MAP_BORDERS)
                      or findMenuContainerForMap(LETTER_MAP_FORMAT)
                      or findFreezeMenuContainer()
                      or findInsertMenuContainer()
                      or findMenuContainerForMap(LETTER_MAP_DELETE)
          if opened then
            eventtap.keyStroke({}, "escape", 0)
          end
          return false 
        end
      end
    end
  end
  return false
end)
optionTap:start()
