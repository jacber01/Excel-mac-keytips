local app = require("hs.application")
local ax = require("hs.axuielement")
local eventtap = require("hs.eventtap")
local canvas = require("hs.canvas")
local screen = require("hs.screen")
local timer = require("hs.timer")
local keycodes = require("hs.keycodes")

local CONFIG = {
  excelBundleID      = "com.microsoft.Excel",
  rootCacheDuration  = 2,
  searchNodeLimit    = 800,
  searchDepthLimit   = 4,
  validationInterval = 0.2,
  activationDelay    = 0.01,
  reactivateDelay    = 0.12,
  mouseActivityDelay = 0.35,
  tipWidth           = 24,
  tipHeight          = 20,
}

local LETTER_MAPS = {
  BORDERS = { B = {"Bottom Border"}, T = {"Top Border"}, L = {"Left Border"}, R = {"Right Border"}, N = {"No Border"}, A = {"All Borders"}, O = {"Outside Borders"}, H = {"Thick Box Border"}, M = {"More Borders"}, D = {"Bottom Double Border"}, K = {"Thick Bottom Border"}, P = {"Top and Bottom Border"}, U = {"Top and Thick Bottom Border"}, E = {"Top and Double Bottom Border"}, G = {"Draw Border Grid"}, S = {"Erase Border"} },
  MERGE = { C = {"Center"}, A = {"Merge Across"}, M = {"Merge Cells"}, U = {"Unmerge Cells"}},
  FORMAT  = { H = {"Row Height"}, A = {"AutoFit Row Height"}, W = {"Column Width"}, I = {"AutoFit Column Width"}, D = {"Default Width"}, R = {"Rename Sheet"}, M = {"Move or Copy Sheet"}, P = {"Protect Sheet"}, L = {"Lock Cell"}, F = {"Format Cells"} },
  FREEZE  = { F = {"Freeze Panes"}, R = {"Freeze Top Row"}, C = {"Freeze First Column"} },
  INSERT  = { I = {"Insert Cells"}, R = {"Insert Sheet Rows"}, C = {"Insert Sheet Columns"}, S = {"Insert Sheet"} },
  DELETE  = { D = {"Delete Cells"}, R = {"Delete Sheet Rows"}, C = {"Delete Sheet Columns"}, T = {"Delete Table Rows"}, L = {"Delete Table Columns"}, S = {"Delete Sheet"} },
  DELETE  = { D = {"Delete Cells"}, R = {"Delete Sheet Rows"}, C = {"Delete Sheet Columns"}, T = {"Delete Table Rows"}, L = {"Delete Table Columns"}, S = {"Delete Sheet"} },
  PASTE = { A = {"Paste"}, F = {"Formulas"}, N = {"Formulas & Number Formatting"}, S = {"Keep Source Formatting"}, B = {"No Borders"}, W = {"Keep Source Column Widths"}, T = {"Transpose"}, V = {"Paste Values"}, U = {"Values & Number Formatting"}, O = {"Values & Source Formatting"}, R = {"Formatting"}, L = {"Paste Link"}, P = {"Paste as Picture"}, I = {"Linked Picture"}, M = {"Match Destination Formatting"}, E = {"Paste Special"}},
  COPY  = { C = {"Copy"}, P = {"Copy as Picture"}},
  GROUP  = { G = {"Group"}},
  UNGROUP  = { U = {"Ungroup"}},
  AUTOSUM  = { S = {"Sum"}, A = {"Average"}, C = {"Count Numbers"}, X = {"Max"}, I = {"Min"}, M = {"More Functions"}},
  FILL  = { W = {"Up"}, S = {"Down"}, D = {"Right"}, A = {"Left"}, E = {"Series"}, R = {"Across Workbooks"}, J = {"Justify"}, F = {"Flash Fill"}},
  CLEAR = { A = {"Clear All"}, F = {"Clear Formats"}, C = {"Clear Contents"}, N = {"Clear Comments and Notes"}, H = {"Clear Hyperlinks"}, R = {"Remove Hyperlinks"}},
  SORT  = { A = {"Sort A to Z"}, Z = {"Sort Z to A"}, C = {"Custom Sort"}, F = {"Filter"}, L = {"Clear"}, R = {"Reapply"}},
  FIND  = { F = {"Find"}, R = {"Replace"}, G = {"Go To"}, S = {"Go To Special"}, O = {"Formulas"}, N = {"Notes"}, C = {"Conditional Formatting"}, S = {"Constants"}, D = {"Data Validation"}, B = {"Select Objects"}, P = {"Selection Pane"}},
}

local CONTEXTS = {
  { name = "borders", map = LETTER_MAPS.BORDERS, minButtons = 5, trigger = "b" },
  { name = "merge", map = LETTER_MAPS.MERGE, minButtons = 4, trigger = "m" },
  { name = "format", map = LETTER_MAPS.FORMAT, minButtons = 5, trigger = "o" },
  { name = "freeze", map = LETTER_MAPS.FREEZE, minButtons = 3, trigger = "f" },
  { name = "insert", map = LETTER_MAPS.INSERT, minButtons = 4, trigger = "i" },
  { name = "delete", map = LETTER_MAPS.DELETE, minButtons = 5, trigger = "d" },
  { name = "paste", map = LETTER_MAPS.PASTE, minButtons = 1, trigger = "v" },
  { name = "copy", map = LETTER_MAPS.COPY, minButtons = 1, trigger = "c" },
  { name = "rotate", map = LETTER_MAPS.ROTATE, minButtons = 4, trigger = "q" },
  { name = "group", map = LETTER_MAPS.GROUP, minButtons = 1, trigger = "g" },
  { name = "ungroup", map = LETTER_MAPS.UNGROUP, minButtons = 1, trigger = "u" },
  { name = "autosum", map = LETTER_MAPS.AUTOSUM, minButtons = 5, trigger = "u" },
  { name = "fill", map = LETTER_MAPS.FILL, minButtons = 6, trigger = "i" },
  { name = "clear", map = LETTER_MAPS.CLEAR, minButtons = 5, trigger = "e" },
  { name = "sort", map = LETTER_MAPS.SORT, minButtons = 5, trigger = "s" },
  { name = "find", map = LETTER_MAPS.FIND, minButtons = 6, trigger = "d" },
}

-- Build trigger key to context mapping (support multiple contexts per key)
local TRIGGER_CONTEXTS = {}
for _, ctx in ipairs(CONTEXTS) do
  if ctx.trigger then
    local code = keycodes.map[ctx.trigger]
    if code then
      TRIGGER_CONTEXTS[code] = TRIGGER_CONTEXTS[code] or {}
      table.insert(TRIGGER_CONTEXTS[code], ctx)
    end
  end
end

local cache = { root = nil, lastUpdate = 0 }
local state = {
  active          = false,
  activeContainer = nil,
  tips            = {},
  items           = {},
  tap             = nil,
  context         = "none",
  checkTimer      = nil,
}
local masterEnabled = true
local hardStopped = false

local mainTap = nil
local appWatcher = nil
local activityTimer = nil
local suppressAutoFromMouse = false
local suppressAutoFromOption = false
local optionIsDown = false

-- ---------- Optimized helpers ----------
local function excelApp()
  return hs.application.find(CONFIG.excelBundleID)
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

  if (letterMap == LETTER_MAPS.INSERT or letterMap == LETTER_MAPS.DELETE) and letter == "S" then
    local lowerLbl = lbl:lower()
    if lowerLbl:find("columns", 1, true) or lowerLbl:find("rows", 1, true) then
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

local function anyContextMenuIsOpen()
  for _, ctx in ipairs(CONTEXTS) do
    if findContainerOptimized(ctx.map, ctx.minButtons) then
      return true
    end
  end
  return false
end

-- ---------- Cached UI traversal ----------
local function updateRootCache()
  local now = os.time()
  if cache.root and (now - cache.lastUpdate) < CONFIG.rootCacheDuration then
    return cache.root
  end

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
  return cache.root
end

function findContainerOptimized(letterMap, minButtons)
  local root = updateRootCache()
  if not root then return nil end

  local queue = { root }
  local processed = 0

  while #queue > 0 and processed < CONFIG.searchNodeLimit do
    processed = processed + 1
    local el = table.remove(queue, 1)

    if safeAttr(el, "AXRole") == "AXGroup" then
      local kids = safeAttr(el, "AXChildren")
      if kids then
        local btnCount, matchCount = 0, 0
        for _, ch in ipairs(kids) do
          if safeAttr(ch, "AXRole") == "AXMenuButton" then
            if visibleOnScreen(getFrame(ch)) then
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

        if processed < CONFIG.searchNodeLimit - #kids then
          for _, ch in ipairs(kids) do table.insert(queue, ch) end
        end
      end
    else
      local kids = safeAttr(el, "AXChildren")
      if kids and processed < CONFIG.searchNodeLimit - #kids then
        for _, ch in ipairs(kids) do table.insert(queue, ch) end
      end
    end
  end
  return nil
end

local function collectItemsOptimized(container, letterMap)
  if not container then return {} end

  local found = {}
  local function traverse(el, depth)
    if depth > CONFIG.searchDepthLimit then return end

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

-- ---------- Visual tips ----------
local function clearTips()
  for _, c in pairs(state.tips) do
    pcall(function() c:hide(); c:delete() end)
  end
  state.tips, state.items = {}, {}
  if state.tap then state.tap:stop(); state.tap = nil end
  if state.checkTimer then state.checkTimer:stop(); state.checkTimer = nil end
  state.active = false
  state.context = "none"
  state.activeContainer = nil
end

local function showTip(letter, frame)
  local w, h = CONFIG.tipWidth, CONFIG.tipHeight
  local x = frame.x + math.floor((frame.w - w) / 2)
  local y = frame.y + math.max(0, math.floor((frame.h - h) / 2))

  local scr = screenForRectCompat(frame):frame()
  x = math.max(scr.x + 2, math.min(x, scr.x + scr.w - w - 2))
  y = math.max(scr.y + 2, y)

  local c = canvas.new({ x = x, y = y, w = w, h = h })
  c:appendElements({
    { type = "rectangle", action = "fill",
      roundedRectRadii = { xRadius = 6, yRadius = 6 },
      fillColor = { red = 0.4, green = 0.4, blue = 0.4, alpha = 1 },
      strokeColor = { white = 0, alpha = 0 }, strokeWidth = 0 },
    { type = "text", text = letter, textSize = 13,
      textColor = { red = 1, green = 1, blue = 1, alpha = 1 },
      frame = { x = 0, y = 2, w = w, h = h },
      textAlignment = "center" }
  })
  c:level(canvas.windowLevels.overlay)
  c:show()
  state.tips[letter] = c
end

-- ---------- Main activation logic ----------
local function activateKeytips(request)
  if hardStopped or not masterEnabled or not isExcelFrontmost() then
    return
  end
  clearTips()

  local function tryContexts(list)
    for _, ctx in ipairs(list) do
      local container = findContainerOptimized(ctx.map, ctx.minButtons)
      if container then
        local items = collectItemsOptimized(container, ctx.map)
        if next(items) ~= nil then
          state.items = items
          state.context = ctx.name
          state.activeContainer = container
          return true
        end
      end
    end
    return false
  end

  if type(request) == "table" then
    tryContexts(request)
  elseif type(request) == "string" then
    for _, ctx in ipairs(CONTEXTS) do
      if ctx.name == request then
        tryContexts({ ctx })
        break
      end
    end
  else
    tryContexts(CONTEXTS)
  end

  if not state.activeContainer then return end

  for letter, info in pairs(state.items) do
    showTip(letter, info.frame)
  end

  state.tap = eventtap.new({ eventtap.event.types.keyDown }, function(ev)
    if hardStopped or not state.active then return false end
    local char = ev:getCharacters()
    if not char or char == "" then return true end
    local letter = string.upper(char)
    local target = state.items[letter]

    if target and target.el then
      pcall(function() target.el:performAction("AXPress") end)
    end
    clearTips()
    -- Removed auto-reactivation after key press
    return true
  end)

  state.active = true
  state.tap:start()

  state.checkTimer = timer.doEvery(CONFIG.validationInterval, function()
    if not safeAttr(state.activeContainer, "AXRole") then
      clearTips()
    end
  end)
  state.checkTimer:start()
end

-- ---------- Consolidated event handling ----------
local function handleKeyDown(ev)
  if hardStopped or not masterEnabled or not isExcelFrontmost() then return false end

  local keyCode = ev:getKeyCode()

  if keyCode == keycodes.map.escape and state.active then
    -- Only handle ESC when keytips overlay is active
    eventtap.keyStroke({}, "escape", 0)

    local frame = state.activeContainer and getFrame(state.activeContainer) or nil
    if frame then
      local rawX = frame.x + frame.w / 2
      local rawY = frame.y - 20
      local scr = screenForRectCompat(frame):frame()
      local cx = math.max(scr.x, math.min(rawX, scr.x + scr.w - 1))
      local cy = math.max(scr.y, math.min(rawY, scr.y + scr.h - 1))

      hs.eventtap.leftClick({ x = cx, y = cy })
    end

    masterEnabled = false
    clearTips()
    return true
  end

  local ctxList = TRIGGER_CONTEXTS[keyCode]
  if ctxList then
    if activityTimer then activityTimer:stop() end
    activityTimer = timer.doAfter(CONFIG.activationDelay, function()
      activityTimer = nil
      if not hardStopped and masterEnabled then activateKeytips(ctxList) end
    end)
    return false
  end

  if suppressAutoFromOption then
    suppressAutoFromOption = false
  end

  return false
end

local function handleMouseDown()
  if hardStopped or not isExcelFrontmost() then return false end

  if state.active then
    clearTips()
  end

  timer.doAfter(CONFIG.reactivateDelay, function()
    if anyContextMenuIsOpen() then
      suppressAutoFromMouse = true
    end
  end)

  return false
end

local function handleMouseUp()
  if hardStopped or not masterEnabled or not isExcelFrontmost() then return false end

  -- Removed auto-activation on mouse activity
  return false
end

local function handleFlagsChanged(ev)
  if hardStopped or not isExcelFrontmost() then return false end

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
          if anyContextMenuIsOpen() then
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

  cache = { root = nil, lastUpdate = 0 }
  hs.alert.show("Excel KeyTips: Disabled\n(Reload config to re-enable)")
end

mainTap = eventtap.new({
  eventtap.event.types.keyDown,
  eventtap.event.types.leftMouseDown,
  eventtap.event.types.leftMouseUp,
  eventtap.event.types.flagsChanged,
}, function(ev)
  local evType = ev:getType()
  if evType == eventtap.event.types.keyDown then
    return handleKeyDown(ev)
  elseif evType == eventtap.event.types.leftMouseDown then
    return handleMouseDown()
  elseif evType == eventtap.event.types.leftMouseUp then
    return handleMouseUp()
  elseif evType == eventtap.event.types.flagsChanged then
    return handleFlagsChanged(ev)
  end
  return false
end)

appWatcher = app.watcher.new(function(appName, eventType, appObj)
  if appObj:bundleID() == CONFIG.excelBundleID then
    if eventType == app.watcher.activated then
      if hardStopped then return end
      mainTap:start()
    elseif eventType == app.watcher.deactivated then
      if hardStopped then return end
      mainTap:stop()
      if activityTimer then activityTimer:stop(); activityTimer = nil end
      clearTips()
    elseif eventType == app.watcher.terminated then
      teardownAndDisable()
    end
  end
end)

appWatcher:start()
if isExcelFrontmost() and mainTap then
  mainTap:start()
end
