local app = require("hs.application")
local ax = require("hs.axuielement")
local eventtap = require("hs.eventtap")
local canvas = require("hs.canvas")
local screen = require("hs.screen")
local timer = require("hs.timer")
local keycodes = require("hs.keycodes")

local EXCEL_BUNDLE = "com.microsoft.Excel"

-- Letter maps
local LETTER_MAP_BORDERS = {
B = { "Bottom Border" }, T = { "Top Border" }, L = { "Left Border" },
R = { "Right Border" }, N = { "No Border" }, A = { "All Borders" },
O = { "Outside Borders" }, H = { "Thick Box Border" }, M = { "More Borders" },
D = { "Bottom Double Border" }, K = { "Thick Bottom Border" },
P = { "Top and Bottom Border" }, U = { "Top and Thick Bottom Border" },
E = { "Top and Double Bottom Border" }, G = { "Draw Border Grid" },
S = { "Erase Border" },
}

local LETTER_MAP_FORMAT = {
H = { "Row Height" }, A = { "AutoFit Row Height" }, W = { "Column Width" },
I = { "AutoFit Column Width" }, D = { "Default Width" }, R = { "Rename Sheet" },
M = { "Move or Copy Sheet" }, P = { "Protect Sheet" }, L = { "Lock Cell" },
F = { "Format Cells" },
}

local LETTER_MAP_FREEZE = {
F = { "Freeze Panes" }, R = { "Freeze Top Row" }, C = { "Freeze First Column" },
}

local LETTER_MAP_INSERT = {
I = { "Insert Cells" }, R = { "Insert Sheet Rows" }, C = { "Insert Sheet Columns" },
S = { "Insert Sheet" },
}

local LETTER_MAP_DELETE = {
D = { "Delete Cells" }, R = { "Delete Sheet Rows" }, C = { "Delete Sheet Columns" },
T = { "Delete Table Rows" }, L = { "Delete Table Columns" }, S = { "Delete Sheet" },
}

-- Cache and state management
local cache = { root = nil, lastUpdate = 0, containers = {} }
local state = {
  active = false, tips = {}, items = {}, tap = nil,
  context = "none", checkTimer = nil
}
local masterEnabled = true
local hardStopped = false

-- Consolidated event taps and timers
local mainTap = nil
local appWatcher = nil
local activityTimer = nil
local suppressAutoFromMouse = false
local suppressAutoFromOption = false
local optionIsDown = false

-- ---------- Optimized helpers ----------
local excelApp = function()
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
  if not el then return nil end
  local ok, val = pcall(function() return el:attributeValue(attr) end)
  return ok and val or nil
end

local function getFrame(el)
  if not el then return nil end
  local f = safeAttr(el, "AXFrame")
  if f and f.w and f.h then return f end
  local p, s = safeAttr(el, "AXPosition"), safeAttr(el, "AXSize")
  return p and s and { x = p.x, y = p.y, w = s.w, h = s.h } or nil
end

local function screenForRectCompat(rect)
  local bestScr, bestArea = screen.mainScreen(), -1
  for _, scr in ipairs(screen.allScreens()) do
    local sf = scr:frame()
    local ix1, iy1 = math.max(sf.x, rect.x), math.max(sf.y, rect.y)
    local ix2, iy2 = math.min(sf.x + sf.w, rect.x + rect.w), math.min(sf.y + sf.h, rect.y + rect.h)
    local area = math.max(0, ix2 - ix1) * math.max(0, iy2 - iy1)
    if area > bestArea then bestArea, bestScr = area, scr end
  end
  return bestScr
end

local function visibleOnScreen(frame)
  if not frame or frame.w <= 1 or frame.h <= 1 then return false end
  local s = screenForRectCompat(frame):frame()
  return frame.x + frame.w > s.x and frame.x < s.x + s.w
     and frame.y + frame.h > s.y and frame.y < s.y + s.h
end

local function labelOf(el)
  if not el then return "" end
  return safeAttr(el, "AXTitle") or safeAttr(el, "AXDescription") or 
         safeAttr(el, "AXHelp") or 
         (type(safeAttr(el, "AXValue")) == "string" and safeAttr(el, "AXValue")) or ""
end

local function labelMatches(letterMap, letter, lbl)
  local pats = letterMap[letter]
  if not pats or not lbl or lbl == "" then return false end

  -- Disambiguation for Insert/Delete 'S'
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

-- ---------- Cached UI traversal ----------
local function updateRootCache()
  local now = os.time()
  if cache.root and (now - cache.lastUpdate) < 2 then return cache.root end
  
  if not excelRunning() then
    cache.root = nil
    return nil
  end
  
  local e = excelApp()
  if not e then
    cache.root = nil
    return nil
  end
  
  local ok, root = pcall(function() return ax.applicationElement(e) end)
  cache.root = (ok and root) and root or nil
  cache.lastUpdate = now
  cache.containers = {} -- Clear container cache
  return cache.root
end

local function findContainerOptimized(letterMap, minButtons)
  local root = updateRootCache()
  if not root then return nil end
  
  -- Use breadth-first with early termination
  local queue = { root }
  local processed = 0
  local maxProcessed = 800 -- Reduced from 1500
  
  while #queue > 0 and processed < maxProcessed do
    processed = processed + 1
    local el = table.remove(queue, 1)
    
    if safeAttr(el, "AXRole") == "AXGroup" then
      local kids = safeAttr(el, "AXChildren")
      if kids then
        local btnCount, matchCount = 0, 0
        -- Process children in batch
        for _, ch in ipairs(kids) do
          if safeAttr(ch, "AXRole") == "AXMenuButton" then
            local f = getFrame(ch)
            if f and visibleOnScreen(f) then
              btnCount = btnCount + 1
              if anyLabelMatches(letterMap, labelOf(ch)) then
                matchCount = matchCount + 1
              end
            end
          end
        end
        
        if btnCount >= minButtons and matchCount >= 1 then
          return el
        end
        
        -- Add children to queue (but limit depth)
        if processed < maxProcessed - #kids then
          for _, ch in ipairs(kids) do
            table.insert(queue, ch)
          end
        end
      end
    else
      local kids = safeAttr(el, "AXChildren")
      if kids and processed < maxProcessed - #kids then
        for _, ch in ipairs(kids) do
          table.insert(queue, ch)
        end
      end
    end
  end
  return nil
end

-- Optimized container finders
local function findMenuContainerForMap(letterMap)
  return findContainerOptimized(letterMap, 5)
end

local function findFreezeMenuContainer()
  return findContainerOptimized(LETTER_MAP_FREEZE, 3)
end

local function findInsertMenuContainer()
  return findContainerOptimized(LETTER_MAP_INSERT, 4)
end

-- ---------- Optimized item collection ----------
local function collectItemsOptimized(container, letterMap)
  if not container then return {} end
  
  local found = {}
  local function traverse(el, depth)
    if depth > 4 then return end -- Reduced max depth
    
    if safeAttr(el, "AXRole") == "AXMenuButton" then
      local lbl = labelOf(el)
      local f = getFrame(el)
      if f and visibleOnScreen(f) and lbl ~= "" then
        for letter, _ in pairs(letterMap) do
          if not found[letter] and labelMatches(letterMap, letter, lbl) then
            found[letter] = { el = el, frame = f, label = lbl }
            break
          end
        end
      end
    end
    
    local kids = safeAttr(el, "AXChildren")
    if kids then
      for _, ch in ipairs(kids) do
        traverse(ch, depth + 1)
      end
    end
  end
  
  traverse(container, 0)
  return found
end

-- ---------- Visual tips (unchanged) ----------
local function clearTips()
  for _, c in pairs(state.tips) do
    pcall(function() c:hide(); c:delete() end)
  end
  state.tips, state.items = {}, {}
  if state.tap then state.tap:stop(); state.tap = nil end
  if state.checkTimer then state.checkTimer:stop(); state.checkTimer = nil end
  state.active = false
  state.context = "none"
end

local function showTip(letter, frame)
  local w, h = 24, 20
  local x = frame.x + math.floor((frame.w - w) / 2)
  local y = frame.y + math.max(0, math.floor((frame.h - h) / 2))

  local scr = screenForRectCompat(frame):frame()
  x = math.max(scr.x + 2, math.min(x, scr.x + scr.w - w - 2))
  y = math.max(scr.y + 2, y)

  local c = canvas.new({ x = x, y = y, w = w, h = h })
  c:appendElements(
    { type = "rectangle", action = "fill",
      roundedRectRadii = { xRadius = 6, yRadius = 6 },
      fillColor = { red = 0.4, green = 0.4, blue = 0.4, alpha = 1 },
      strokeColor = { white = 0, alpha = 0 }, strokeWidth = 0 },
    { type = "text", text = letter, textSize = 13,
      textColor = { red = 1, green = 1, blue = 1, alpha = 1 },
      frame = { x = 0, y = 2, w = w, h = h },
      textAlignment = "center" }
  )
  c:level(canvas.windowLevels.overlay)
  c:show()
  state.tips[letter] = c
end

-- ---------- Main activation logic ----------
local function activateKeytips()
  if hardStopped or not masterEnabled or not isExcelFrontmost() or not excelRunning() then 
    return 
  end
  
  clearTips()
  
  -- Priority order with optimized detection
  local contexts = {
    { name = "borders", map = LETTER_MAP_BORDERS, finder = findMenuContainerForMap },
    { name = "format", map = LETTER_MAP_FORMAT, finder = findMenuContainerForMap },
    { name = "freeze", map = LETTER_MAP_FREEZE, finder = findFreezeMenuContainer },
    { name = "insert", map = LETTER_MAP_INSERT, finder = findInsertMenuContainer },
    { name = "delete", map = LETTER_MAP_DELETE, finder = findMenuContainerForMap }
  }
  
  local container, items, which
  for _, ctx in ipairs(contexts) do
    container = ctx.finder(ctx.map)
    if container then
      items = collectItemsOptimized(container, ctx.map)
      if next(items) ~= nil then
        which = ctx.name
        break
      end
    end
  end
  
  if not which then return end
  
  state.items, state.context = items, which
  for letter, info in pairs(items) do
    showTip(letter, info.frame)
  end

  -- Set up interaction handlers
  state.tap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
    if hardStopped or not state.active then return false end
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

  -- Validation timer with longer interval
  state.checkTimer = timer.doEvery(0.1, function() -- Increased from 0.05
    local currentContainer
    if state.context == "borders" then
      currentContainer = findMenuContainerForMap(LETTER_MAP_BORDERS)
    elseif state.context == "format" then
      currentContainer = findMenuContainerForMap(LETTER_MAP_FORMAT)
    elseif state.context == "freeze" then
      currentContainer = findFreezeMenuContainer()
    elseif state.context == "insert" then
      currentContainer = findInsertMenuContainer()
    elseif state.context == "delete" then
      currentContainer = findMenuContainerForMap(LETTER_MAP_DELETE)
    end
    if not currentContainer then clearTips() end
  end)
  state.checkTimer:start()
end

-- ---------- Consolidated event handling ----------
local function handleKeyDown(ev)
  if hardStopped or not masterEnabled or not isExcelFrontmost() or not excelRunning() then 
    return false 
  end
  
  local keyCode = ev:getKeyCode()
  local triggerKeys = {
    [keycodes.map["b"]] = true, [keycodes.map["o"]] = true, 
    [keycodes.map["f"]] = true, [keycodes.map["i"]] = true, 
    [keycodes.map["d"]] = true
  }
  
  -- ESC handling
  if keyCode == 53 and state.active then -- ESC
    masterEnabled = false
    clearTips()
    return true
  end
  
  -- Trigger key handling
  if triggerKeys[keyCode] then
    if activityTimer then activityTimer:stop() end
    activityTimer = timer.doAfter(0.08, function()
      activityTimer = nil
      if not hardStopped and masterEnabled then activateKeytips() end
    end)
    return false
  end
  
  -- Clear auto-suppress on normal typing
  if suppressAutoFromOption then
    suppressAutoFromOption = false
  end
  
  return false
end

local function handleMouseDown(ev)
  if hardStopped then return false end
  if not isExcelFrontmost() or not excelRunning() then return false end
  
  -- Hide tips immediately on click
  if state.active and (state.context == "borders" or state.context == "format" or 
                      state.context == "freeze" or state.context == "insert" or 
                      state.context == "delete") then
    clearTips()
    timer.doAfter(0.12, function()
      if not hardStopped and masterEnabled then activateKeytips() end
    end)
  end
  
  -- Detect mouse-opened dropdowns
  timer.doAfter(0.12, function()
    local opened = findMenuContainerForMap(LETTER_MAP_BORDERS) or
                  findMenuContainerForMap(LETTER_MAP_FORMAT) or
                  findFreezeMenuContainer() or
                  findInsertMenuContainer() or
                  findMenuContainerForMap(LETTER_MAP_DELETE)
    if opened then suppressAutoFromMouse = true end
  end)
  
  return false
end

local function handleMouseUp(ev)
  if hardStopped or not masterEnabled or not isExcelFrontmost() or not excelRunning() then 
    return false 
  end
  
  -- Activity-based activation with debouncing
  if activityTimer then activityTimer:stop() end
  activityTimer = timer.doAfter(0.35, function()
    activityTimer = nil
    if not suppressAutoFromMouse and not suppressAutoFromOption and 
       not hardStopped and masterEnabled then
      activateKeytips()
    end
  end)
  
  return false
end

local function handleFlagsChanged(ev)
  if hardStopped or not isExcelFrontmost() or not excelRunning() then return false end

  local flags = ev:getFlags()
  local nowDown = flags.alt or flags.altgr

  if nowDown ~= optionIsDown then
    optionIsDown = nowDown
    if nowDown then
      if state.active or not masterEnabled then
        if masterEnabled then
          masterEnabled = false
          clearTips()
          suppressAutoFromOption = false
          return true
        else
          masterEnabled = true
          suppressAutoFromMouse = false
          suppressAutoFromOption = true
          
          -- Close any open dropdown
          local opened = findMenuContainerForMap(LETTER_MAP_BORDERS) or
                        findMenuContainerForMap(LETTER_MAP_FORMAT) or
                        findFreezeMenuContainer() or
                        findInsertMenuContainer() or
                        findMenuContainerForMap(LETTER_MAP_DELETE)
          if opened then
            eventtap.keyStroke({}, "escape", 0)
          end
          return false 
        end
      end
    end
  end
  return false
end

-- ---------- System integration ----------
local function teardownAndDisable()
  if hardStopped then return end
  hardStopped, masterEnabled = true, false
  clearTips()
  
  if mainTap then mainTap:stop(); mainTap = nil end
  if appWatcher then appWatcher:stop(); appWatcher = nil end
  if activityTimer then activityTimer:stop(); activityTimer = nil end
  
  -- Clear cache
  cache = { root = nil, lastUpdate = 0, containers = {} }
  
  hs.alert.show("Excel KeyTips: disabled (reload Hammerspoon to re-enable)")
end

-- Consolidated main event tap
mainTap = eventtap.new({
  eventtap.event.types.keyDown,
  eventtap.event.types.leftMouseDown,
  eventtap.event.types.leftMouseUp,
  eventtap.event.types.flagsChanged
}, function(ev)
  local evType = ev:getType()
  if evType == eventtap.event.types.keyDown then
    return handleKeyDown(ev)
  elseif evType == eventtap.event.types.leftMouseDown then
    return handleMouseDown(ev)
  elseif evType == eventtap.event.types.leftMouseUp then
    return handleMouseUp(ev)
  elseif evType == eventtap.event.types.flagsChanged then
    return handleFlagsChanged(ev)
  end
  return false
end)

-- App lifecycle management
appWatcher = app.watcher.new(function(appName, eventType, appObj)
  if appName == "Microsoft Excel" then
    if eventType == app.watcher.activated then
      if hardStopped then return end
      if mainTap then mainTap:start() end
      timer.doAfter(0.15, function() 
        if not hardStopped then activateKeytips() end 
      end)
    elseif eventType == app.watcher.deactivated then
      if hardStopped then return end
      if mainTap then mainTap:stop() end
      if activityTimer then activityTimer:stop(); activityTimer = nil end
      clearTips()
    elseif eventType == app.watcher.terminated then
      teardownAndDisable()
    end
  end
end)

-- Initialize
appWatcher:start()
if isExcelFrontmost() and excelRunning() and mainTap then
  mainTap:start()
end
