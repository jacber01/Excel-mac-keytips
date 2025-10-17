local app       = require("hs.application")
local ax        = require("hs.axuielement")
local eventtap  = require("hs.eventtap")
local canvas    = require("hs.canvas")
local screen    = require("hs.screen")
local timer     = require("hs.timer")
local keycodes  = require("hs.keycodes")

-- Keytip Maps and Contexts
local LETTER_MAPS = {
    BORDERS = {O={"Bottom Border"},P={"Top Border"},L={"Left Border"},R={"Right Border"},N={"No Border"},A={"All Borders"},S={"Outside Borders"},T={"Thick Box Border"},M={"More Borders"},B={"Bottom Double Border"},H={"Thick Bottom Border"},C={"Top and Bottom Border"},U={"Top and Thick Bottom Border"},E={"Top and Double Bottom Border"},G={"Draw Border Grid"},X={"Erase Border"}},
    FORMAT  = {H={"Row Height"},A={"AutoFit Row Height"},W={"Column Width"},I={"AutoFit Column Width"},D={"Default Width"},R={"Rename Sheet"},M={"Move or Copy Sheet"},P={"Protect Sheet"},L={"Lock Cell"},F={"Format Cells"}},
    FREEZE  = {F={"Freeze Panes"},R={"Freeze Top Row"},C={"Freeze First Column"}},
    INSERT  = {I={"Insert Cells"},R={"Insert Sheet Rows"},C={"Insert Sheet Columns"},S={"Insert Sheet"}},
    DELETE  = {D={"Delete Cells"},R={"Delete Sheet Rows"},C={"Delete Sheet Columns"},T={"Delete Table Rows"},L={"Delete Table Columns"},S={"Delete Sheet"}},
    MERGE   = {C={"Center"},A={"Merge Across"},M={"Merge Cells"},U={"Unmerge Cells"}},
    PASTE   = {K={"Keep Source Formatting"},M={"Match Destination Formatting"}, T={"Keep Text Only"}, S={"Paste Special"}},
    COPY    = {C={"Copy"},P={"Copy as Picture"}},
    GROUP   = {G={"Group"}},
    UNGROUP = {U={"Ungroup"}},
    AUTOSUM = {S={"Sum"},A={"Average"},C={"Count Numbers"},X={"Max"},I={"Min"},M={"More Functions"}},
    ROTATE  = {A={"Angle Counterclockwise"},C={"Angle Clockwise"},V={"Vertical Text"},U={"Rotate Text Up"},D={"Rotate Text Down"},M={"Format Cell Alignment"}},
    FILL    = {U={"Up"},D={"Down"},R={"Right"},L={"Left"},S={"Series"},A={"Across Workbooks"},J={"Justify"},F={"Flash Fill"}},
    CLEAR   = {A={"Clear All"},F={"Clear Formats"},C={"Clear Contents"},N={"Clear Comments and Notes"},H={"Clear Hyperlinks"},R={"Remove Hyperlinks"}},
    SORT    = {A={"Sort A to Z"},Z={"Sort Z to A"},C={"Custom Sort"},F={"Filter"},L={"Clear"},R={"Reapply"}},
    FIND    = {F={"Find"},R={"Replace"},G={"Go To"},S={"Go To Special"},K={"Constants"},O={"Formulas"},N={"Notes"},C={"Conditional Formatting"},D={"Data Validation"},B={"Select Objects"},P={"Selection Pane"}},
}

local CONTEXTS = {
    {name="borders", map=LETTER_MAPS.BORDERS, minButtons=5},
    {name="format",  map=LETTER_MAPS.FORMAT,  minButtons=5},
    {name="freeze",  map=LETTER_MAPS.FREEZE,  minButtons=3},
    {name="insert",  map=LETTER_MAPS.INSERT,  minButtons=4},
    {name="delete",  map=LETTER_MAPS.DELETE,  minButtons=5},
    {name="merge",   map=LETTER_MAPS.MERGE,   minButtons=4},
    {name="paste",   map=LETTER_MAPS.PASTE,   minButtons=1},
    {name="copy",    map=LETTER_MAPS.COPY,    minButtons=1},
    {name="group",   map=LETTER_MAPS.GROUP,   minButtons=1},
    {name="ungroup", map=LETTER_MAPS.UNGROUP, minButtons=1},
    {name="autosum", map=LETTER_MAPS.AUTOSUM, minButtons=5},
    {name="rotate",  map=LETTER_MAPS.ROTATE,  minButtons=4},
    {name="fill",    map=LETTER_MAPS.FILL,    minButtons=6},
    {name="clear",   map=LETTER_MAPS.CLEAR,   minButtons=5},
    {name="sort",    map=LETTER_MAPS.SORT,    minButtons=5},
    {name="find",    map=LETTER_MAPS.FIND,    minButtons=6},
}

-- Trigger metadata (used only during the scan window)
local TRIGGER_CONTEXTS = {
    b={{name="borders", parent_menu_trigger="h"}},
    o={{name="format",  parent_menu_trigger="h"}},
    f={{name="freeze",  parent_menu_trigger="w"}},
    i={{name="insert",  parent_menu_trigger="h"}},
    d={{name="delete",  parent_menu_trigger="h"}},
    m={{name="merge",   parent_menu_trigger="h"}},
    v={{name="paste",   parent_menu_trigger="h"}},
    c={{name="copy",    parent_menu_trigger="h"}},
    g={{name="group",   parent_menu_trigger="a"}},
    u={{name="ungroup", parent_menu_trigger="a"},{name="autosum", parent_menu_trigger="h"}},
    fq={{name="rotate", parent_menu_trigger="h"}},
    fi={{name="fill",   parent_menu_trigger="h"}},
    e={{name="clear",   parent_menu_trigger="h"}},
    s={{name="sort",    parent_menu_trigger="h"}},
    fd={{name="find",   parent_menu_trigger="h"}},
}
local PARENT_KEYS = {h=true, w=true, a=true}

-- State and Utility
local cache = {root=nil, lastUpdate=0}
local state = {active=false, activeContainer=nil, tips={}, items={}, tap=nil, checkTimer=nil}
local masterEnabled, hardStopped, scanWindowActive = true, false, false
local mainTap, appWatcher, activityTimer, scanWindowTimer = nil, nil, nil, nil

-- Sequence timing (active only during scan window)
local keySeqBuffer, lastKeyTs = "", 0
local lastParentKey, lastParentTs, lastAltDownTs = nil, 0, 0
local SEQ_TIMEOUT, CHAIN_TIMEOUT = 1.5, 3.0
local altIsDown = false
local containerGoneSince = nil

local function nowSeconds() if timer and timer.secondsSinceEpoch then return timer.secondsSinceEpoch() end return os.time() end
local function excelApp() return hs.application.find("com.microsoft.Excel") end
local function excelRunning() local e=excelApp() return e and e:isRunning() end
local function isExcelFrontmost() local e=excelApp() return e and e:isFrontmost() end

-- Accessibility Helpers
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
    return p and s and {x=p.x, y=p.y, w=s.w, h=s.h} or nil
end

local function screenForRectCompat(rect)
    local bestScr, bestArea = screen.mainScreen(), -1
    for _, scr in ipairs(screen.allScreens()) do
        local sf = scr:frame()
        local ix1, iy1 = math.max(sf.x, rect.x), math.max(sf.y, rect.y)
        local ix2, iy2 = math.min(sf.x+sf.w, rect.x+rect.w), math.min(sf.y+sf.h, rect.y+rect.h)
        local area = math.max(0, ix2-ix1) * math.max(0, iy2-iy1)
        if area > bestArea then bestArea, bestScr = area, scr end
    end
    return bestScr
end

local function visibleOnScreen(frame)
    if not frame or frame.w <= 1 or frame.h <= 1 then return false end
    local s = screenForRectCompat(frame):frame()
    return frame.x+frame.w > s.x and frame.x < s.x+s.w and frame.y+frame.h > s.y and frame.y < s.y+s.h
end

local function labelOf(el)
    if not el then return "" end
    return safeAttr(el, "AXTitle") or safeAttr(el, "AXDescription") or safeAttr(el, "AXHelp")
        or (type(safeAttr(el, "AXValue"))=="string" and safeAttr(el, "AXValue")) or ""
end

local function labelMatches(letterMap, letter, lbl)
    local pats = letterMap[letter]
    if not pats or not lbl or lbl=="" then return false end
    if (letterMap==LETTER_MAPS.INSERT or letterMap==LETTER_MAPS.DELETE) and letter=="S" then
        local lowerLbl = lbl:lower()
        if lowerLbl:find("columns",1,true) or lowerLbl:find("rows",1,true) then return false end
    end
    local lower = lbl:lower()
    for _, pat in ipairs(pats) do if lower:find(pat:lower(),1,true) then return true end end
    return false
end

local function anyLabelMatches(letterMap, lbl)
    if not lbl or lbl=="" then return false end
    local lower = lbl:lower()
    for _, pats in pairs(letterMap) do
        for _, pat in ipairs(pats) do if lower:find(pat:lower(),1,true) then return true end end
    end
    return false
end

-- UI Traversal (Cached)
local function updateRootCache()
    local now = os.time()
    if cache.root and (now - cache.lastUpdate) < 2 then return cache.root end
    if not excelRunning() then cache.root=nil return nil end
    local e = excelApp(); if not e then cache.root=nil return nil end
    local ok, root = pcall(function() return ax.applicationElement(e) end)
    cache.root = (ok and root) and root or nil
    cache.lastUpdate = now
    return cache.root
end

function findContainerOptimized(letterMap, minButtons)
    local root = updateRootCache(); if not root then return nil end
    local queue, processed = {root}, 0
    while #queue>0 and processed<800 do
        processed = processed + 1
        local el = table.remove(queue,1)

        local kids = safeAttr(el,"AXChildren")
        if safeAttr(el,"AXRole")=="AXGroup" then
            if kids then
                local btnCount, matchCount = 0, 0
                for _, ch in ipairs(kids) do
                    if safeAttr(ch,"AXRole")=="AXMenuButton" then
                        if visibleOnScreen(getFrame(ch)) then
                            btnCount = btnCount + 1
                            if anyLabelMatches(letterMap, labelOf(ch)) then matchCount = matchCount + 1 end
                        end
                    end
                end
                if btnCount>=minButtons and matchCount>=1 then return el end
            end
        end
        if kids and processed < 800 - #kids then
            for _, ch in ipairs(kids) do table.insert(queue,ch) end
        end
    end
    return nil
end

local function collectItemsOptimized(container, letterMap)
    if not container then return {} end
    local found = {}
    local function traverse(el, depth)
        if depth>4 then return end
        if safeAttr(el,"AXRole")=="AXMenuButton" then
            local lbl = labelOf(el); local f = getFrame(el)
            if f and visibleOnScreen(f) and lbl~="" then
                for letter,_ in pairs(letterMap) do
                    if not found[letter] and labelMatches(letterMap, letter, lbl) then
                        found[letter] = {el=el, frame=f, label=lbl}; break
                    end
                end
            end
        end
        local kids = safeAttr(el,"AXChildren")
        if kids then for _, ch in ipairs(kids) do traverse(ch, depth+1) end end
    end
    traverse(container, 0)
    return found
end

-- Overlay Rendering
local function clearTips()
    for _, c in pairs(state.tips) do pcall(function() c:hide(); c:delete() end) end
    state.tips, state.items = {}, {}
    if state.tap then state.tap:stop(); state.tap=nil end
    if state.checkTimer then state.checkTimer:stop(); state.checkTimer=nil end
    state.active, state.activeContainer = false, nil
    keySeqBuffer = ""; lastParentKey = nil
end

local function showTip(letter, frame)
    local x = frame.x + math.floor((frame.w - 18)/2)
    local y = frame.y + math.max(0, math.floor((frame.h - 20)/2))
    local scr = screenForRectCompat(frame):frame()
    x = math.max(scr.x+2, math.min(x, scr.x+scr.w-18-2))
    y = math.max(scr.y+2, y)
    local c = canvas.new({x=x, y=y, w=18, h=20})
    c:appendElements({
        {type="rectangle", action="fill", roundedRectRadii={xRadius=6, yRadius=6}, fillColor={red=0.4,green=0.4,blue=0.4,alpha=1}, strokeColor={white=0,alpha=0}, strokeWidth=0},
        {type="text", text=letter, textSize=13, textColor={red=1,green=1,blue=1,alpha=1}, frame={x=0,y=2,w=18,h=20}, textAlignment="center"},
    })
    c:level(canvas.windowLevels.overlay); c:show()
    state.tips[letter] = c
end

-- Scan Window Control
local function startScanWindow()
    scanWindowActive = true
    if scanWindowTimer then scanWindowTimer:stop(); scanWindowTimer=nil end
    scanWindowTimer = timer.doAfter(8, function()
        scanWindowActive = false
        if activityTimer then activityTimer:stop(); activityTimer=nil end
        if not state.active then clearTips() end
        keySeqBuffer = ""; lastParentKey = nil
    end)
end

local function cancelScanWindow()
    scanWindowActive = false
    if scanWindowTimer then scanWindowTimer:stop(); scanWindowTimer=nil end
    if activityTimer then activityTimer:stop(); activityTimer=nil end
    clearTips()
end

-- Activation Logic
local function activateKeytips(request)
    if hardStopped or not masterEnabled or not isExcelFrontmost() or not scanWindowActive then return end
    clearTips()

    local function tryContexts(list)
        for _, req in ipairs(list) do
            for _, ctx in ipairs(CONTEXTS) do
                if ctx.name == req.name then
                    local container = findContainerOptimized(ctx.map, ctx.minButtons)
                    if container then
                        local items = collectItemsOptimized(container, ctx.map)
                        if next(items) ~= nil then
                            state.items = items; state.activeContainer = container
                            return true
                        end
                    end
                end
            end
        end
        return false
    end

    if type(request)=="table" and #request>0 then
        if not tryContexts(request) then return end
    else
        for _, ctx in ipairs(CONTEXTS) do
            local container = findContainerOptimized(ctx.map, ctx.minButtons)
            if container then
                local items = collectItemsOptimized(container, ctx.map)
                if next(items)~=nil then state.items=items; state.activeContainer=container; break end
            end
        end
        if not state.activeContainer then return end
    end

    for letter, info in pairs(state.items) do showTip(letter, info.frame) end
    state.tap = eventtap.new({eventtap.event.types.keyDown}, function(ev)
        if hardStopped or not state.active then return false end
        if ev:getKeyCode()==keycodes.map.escape then
            clearTips(); cancelScanWindow(); return false
        end

        local kc = ev:getKeyCode()
        local allowNav = (
            kc == keycodes.map["return"] or
            kc == keycodes.map.enter or
            kc == keycodes.map.tab or
            kc == keycodes.map.left or kc == keycodes.map.right or
            kc == keycodes.map.up   or kc == keycodes.map.down or
            kc == keycodes.map.delete or kc == keycodes.map.forwarddelete
        )
        if allowNav then
            clearTips(); cancelScanWindow(); return false
        end

        local char = ev:getCharacters()
        if char and char~="" then
            local letter = string.upper(char); local target = state.items[letter]
            if target and target.el then
                pcall(function() target.el:performAction("AXPress") end)
                clearTips()
                cancelScanWindow()
                return true
            end
        end

        clearTips(); cancelScanWindow(); return false
    end)
    state.active = true; state.tap:start()
    state.checkTimer = timer.doEvery(0.2, function()
        if not state.activeContainer then return end
        if safeAttr(state.activeContainer,"AXRole") then
            containerGoneSince = nil
        else
            containerGoneSince = containerGoneSince or nowSeconds()
            if nowSeconds() - containerGoneSince > 0.6 then
                clearTips()
            end
        end
    end)
end

-- Event Handlers
local function handleKeyDown(ev)
    if hardStopped or not masterEnabled or not isExcelFrontmost() then return false end
    local keyCode = ev:getKeyCode()
    if keyCode==keycodes.map.escape then cancelScanWindow(); return false end

    local flags = ev:getFlags()
    if flags and (flags.alt or flags.altgr) then
        altIsDown = true
        lastAltDownTs = nowSeconds()
    end

    if scanWindowActive then
        local ts = nowSeconds()
        local char = ev:getCharacters()

        if char and #char==1 then
            local lower = char:lower()
            if PARENT_KEYS[lower] then lastParentKey=lower; lastParentTs=ts end
        end

        if char and #char==1 then
            local lower = char:lower()
            if (ts - lastKeyTs) > SEQ_TIMEOUT then keySeqBuffer = "" end
            keySeqBuffer = (keySeqBuffer .. lower):sub(-2); lastKeyTs = ts

            local candidates = TRIGGER_CONTEXTS[keySeqBuffer] or TRIGGER_CONTEXTS[keySeqBuffer:sub(-1)]
            if candidates then
                local function chainValid(req)
                    local p = req.parent_menu_trigger
                    if not p then return true end
                    if not lastParentKey or lastParentKey ~= p then return false end
                    if (ts - lastParentTs)  > CHAIN_TIMEOUT then return false end
                    if not altIsDown and ((ts - lastAltDownTs) > CHAIN_TIMEOUT) then return false end
                    return true
                end
                local valid = {}
                for _, c in ipairs(candidates) do if chainValid(c) then table.insert(valid, c) end end
                if #valid>0 then
                    if activityTimer then activityTimer:stop() end
                    activityTimer = timer.doAfter(0.03, function()
                        activityTimer=nil
                        if not hardStopped and masterEnabled and scanWindowActive then activateKeytips(valid) end
                    end)
                    keySeqBuffer = ""
                    return false
                end
            end
        end
    end

    if state.active then
        local char = ev:getCharacters()
        if char and char~="" then
            local letter = string.upper(char)
            if not state.items[letter] then
                clearTips()
                startScanWindow()
                return false
            end
        end
    end
    return false
end

local function handleMouseDown()
    if hardStopped or not isExcelFrontmost() then return false end
    cancelScanWindow()
    return false
end

local function handleFlagsChanged(ev)
    if hardStopped or not isExcelFrontmost() then return false end
    local flags = ev:getFlags()
    local nowDown = flags.alt or flags.altgr
    altIsDown = nowDown
    if nowDown then
        lastAltDownTs = nowSeconds()
        if state.active then 
            clearTips() 
        else 
            startScanWindow()
        end
    end
    return false
end

-- Integration
local function teardownAndDisable()
    if hardStopped then return end
    hardStopped, masterEnabled = true, false
    cancelScanWindow()
    if mainTap then mainTap:stop(); mainTap=nil end
    if appWatcher then appWatcher:stop(); appWatcher=nil end
    cache = {root=nil, lastUpdate=0}
    hs.alert.show("Excel KeyTips: Disabled\n(Reload config to re-enable)")
end

mainTap = eventtap.new(
    {eventtap.event.types.keyDown, eventtap.event.types.leftMouseDown, eventtap.event.types.flagsChanged},
    function(ev)
        local t = ev:getType()
        if t==eventtap.event.types.keyDown then return handleKeyDown(ev)
        elseif t==eventtap.event.types.leftMouseDown then return handleMouseDown()
        elseif t==eventtap.event.types.flagsChanged then return handleFlagsChanged(ev)
        end
        return false
    end
)

appWatcher = app.watcher.new(function(appName, eventType, appObj)
    if appObj:bundleID()=="com.microsoft.Excel" then
        if eventType==app.watcher.activated then
            if hardStopped then return end
            mainTap:start()
            timer.doAfter(0.15, function() if not hardStopped and scanWindowActive then activateKeytips() end end)
        elseif eventType==app.watcher.deactivated then
            if hardStopped then return end
            mainTap:stop()
            cancelScanWindow()
        elseif eventType==app.watcher.terminated then
            teardownAndDisable()
        end
    end
end)

appWatcher:start()
if isExcelFrontmost() and mainTap then mainTap:start() end