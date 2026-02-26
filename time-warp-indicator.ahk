#Requires AutoHotkey v2.0
#SingleInstance Force
Persistent

global SETTINGS_PATH := A_ScriptDir "\\settings.ini"
global SUMMARY_DIR := A_ScriptDir "\\summaries"
global SESSION_DIR := A_ScriptDir "\\sessions"
global LOG_DIR := A_ScriptDir "\\logs"
global BOOKMARKS_PATH := A_ScriptDir "\\bookmarks.csv"

global cfg := Map()
global ignoredProcesses := Map()
global focusProcesses := Map()
global distractionProcesses := Map()

global isPaused := false
global overlayVisible := false
global debugEnabled := false
global isWorkstationLocked := false
global isSuspended := false

global lastHwnd := 0
global lastSwitchProc := ""
global lastSwitchTitle := ""
global lastMode := "Normal"
global modeStartNow := ""
global modeTransitionCount := 0
global timeWarpAlertSent := false

global switchTimestamps := []
global wheelTimestamps := []
global keyTimestamps := []
global tabHopTimestamps := []
global switchesTrend := []
global appStats := Map()

global currentFocusProc := ""
global currentFocusTitle := ""
global currentFocusStartTick := A_TickCount
global currentFocusStartNow := A_Now

global overlayGui := ""
global overlayText := ""
global overlayCompact := true

global pausedSnapshot := Map()
global snoozeUntilTick := 0
global focusBlockActive := false
global focusBlockEndTick := 0
global focusBlockMinutes := 25

global timeWarpStreakActive := false
global timeWarpStreakStartNow := ""
global timeWarpStreakSamples := 0
global timeWarpStreakSwitchesSum := 0
global timeWarpStreakWheelSum := 0
global timeWarpStreakKeysSum := 0

global trayStatsLabel := "Now: Initializing"
global trayFocusLabel := "FocusBlock: Off"
global traySnoozeLabel := "Snooze: Off"

global iconMode := ""

Init()
return

Init() {
    global cfg, overlayVisible, debugEnabled, pausedSnapshot, overlayCompact
    global lastMode, modeStartNow, SUMMARY_DIR, SESSION_DIR, LOG_DIR

    EnsureSettingsFile()
    LoadSettings()

    overlayVisible := cfg["overlay_enabled"]
    overlayCompact := cfg["overlay_compact"]
    debugEnabled := cfg["debug_enabled"]

    DirCreate(SUMMARY_DIR)
    DirCreate(SESSION_DIR)
    DirCreate(LOG_DIR)

    SetupOverlay()
    SetupTray()
    SetupInputTracking()
    RegisterSystemNotifications()

    lastMode := "Normal"
    modeStartNow := A_Now

    pausedSnapshot := CurrentMetrics()

    SetTimer(PollForeground, cfg["poll_interval_ms"])
    SetTimer(UpdateUiTick, 1000)
    SetTimer(WriteDailySummary, 60000)
    SetTimer(SampleTrendTick, 60000)
    SetTimer(FocusBlockTick, 1000)
    SetTimer(SnoozeTick, 1000)

    PollForeground()
    UpdateUiTick()
    AppendLog("info", "startup", Map("settings_version", cfg["settings_version"]))
}

SetupTray() {
    global isPaused, trayStatsLabel, trayFocusLabel, traySnoozeLabel

    try TraySetIcon("shell32.dll", 44)

    A_TrayMenu.Delete()
    A_TrayMenu.Add(trayStatsLabel, Noop)
    A_TrayMenu.Disable(trayStatsLabel)

    A_TrayMenu.Add(trayFocusLabel, Noop)
    A_TrayMenu.Disable(trayFocusLabel)

    A_TrayMenu.Add(traySnoozeLabel, Noop)
    A_TrayMenu.Disable(traySnoozeLabel)

    A_TrayMenu.Add()
    A_TrayMenu.Add("Toggle Overlay", ToggleOverlay)
    A_TrayMenu.Add("Toggle Overlay Detail", ToggleOverlayDetail)
    A_TrayMenu.Add("Pause Tracking", TogglePauseTracking)
    if isPaused {
        A_TrayMenu.Check("Pause Tracking")
    }

    A_TrayMenu.Add("Snooze 15 min", Snooze15)
    A_TrayMenu.Add("Snooze 30 min", Snooze30)
    A_TrayMenu.Add("Snooze 60 min", Snooze60)
    A_TrayMenu.Add("Cancel Snooze", CancelSnooze)

    A_TrayMenu.Add("Start Focus Block (25m)", StartFocusBlock25)
    A_TrayMenu.Add("Cancel Focus Block", CancelFocusBlock)

    A_TrayMenu.Add("Reset Counters", ResetCounters)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Show Top Apps", ShowTopApps)
    A_TrayMenu.Add("Generate Weekly Dashboard", GenerateWeeklyDashboard)
    A_TrayMenu.Add("Mark Distraction Moment", MarkDistractionMoment)
    A_TrayMenu.Add("Export Aggregates (JSON/CSV)", ExportAggregatesNow)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Toggle Run At Startup", ToggleRunAtStartup)
    A_TrayMenu.Add("Open Session Folder", OpenSessionFolder)
    A_TrayMenu.Add("Open Settings", OpenSettings)
    A_TrayMenu.Add("Reload Settings", ReloadSettings)
    A_TrayMenu.Add()
    A_TrayMenu.Add("Exit", ExitAppNow)
}

SetupOverlay() {
    global overlayGui, overlayText, cfg, overlayVisible

    overlayGui := Gui("+AlwaysOnTop -Caption +ToolWindow")
    overlayGui.BackColor := cfg["overlay_bg_color"]
    overlayGui.MarginX := 10
    overlayGui.MarginY := 8
    overlayGui.SetFont("s" cfg["font_size"], cfg["font_name"])
    overlayText := overlayGui.AddText("c" cfg["text_color"] " w520", "Initializing...")

    try WinSetTransparent(cfg["overlay_opacity"], overlayGui.Hwnd)
    ApplyOverlayClickThrough()

    PositionOverlay()
    if overlayVisible {
        overlayGui.Show("NoActivate")
    } else {
        overlayGui.Hide()
    }
}

SetupInputTracking() {
    static alreadySetup := false
    if alreadySetup {
        return
    }
    alreadySetup := true

    Hotkey("~*WheelUp", CountWheelEvent)
    Hotkey("~*WheelDown", CountWheelEvent)
    Hotkey("^!o", ToggleOverlay)
    Hotkey("^!m", MarkDistractionMoment)

    keys := [
        "a","b","c","d","e","f","g","h","i","j","k","l","m","n","o","p","q","r","s","t","u","v","w","x","y","z",
        "0","1","2","3","4","5","6","7","8","9",
        "Space","Enter","Tab","Backspace","Delete","NumpadEnter",".",",","/",";","'","[","]","-","=","\\"
    ]

    for key in keys {
        try Hotkey("~*" key, CountKeyEvent)
    }
}

RegisterSystemNotifications() {
    OnMessage(0x02B1, OnSessionChange)
    OnMessage(0x0218, OnPowerBroadcast)
    try DllCall("Wtsapi32\\WTSRegisterSessionNotification", "ptr", A_ScriptHwnd, "uint", 0)
}

OnSessionChange(wParam, lParam, msg, hwnd) {
    global isWorkstationLocked, lastHwnd

    if (wParam = 0x7) {
        isWorkstationLocked := true
        FlushCurrentFocusBlock("session_lock")
        lastHwnd := 0
        AppendLog("info", "session_lock", Map())
    } else if (wParam = 0x8) {
        isWorkstationLocked := false
        lastHwnd := 0
        AppendLog("info", "session_unlock", Map())
    }
}

OnPowerBroadcast(wParam, lParam, msg, hwnd) {
    global isSuspended, lastHwnd

    if (wParam = 4) {
        isSuspended := true
        FlushCurrentFocusBlock("sleep")
        lastHwnd := 0
        AppendLog("info", "sleep", Map())
    } else if (wParam = 7 || wParam = 18) {
        isSuspended := false
        lastHwnd := 0
        AppendLog("info", "resume", Map("code", wParam))
    }
}

Noop(*) {
}

PollForeground() {
    global lastHwnd, switchTimestamps, lastSwitchProc, lastSwitchTitle
    global currentFocusProc, currentFocusStartTick, currentFocusTitle, currentFocusStartNow
    global tabHopTimestamps, appStats

    if IsTrackingSuppressed() {
        return
    }

    hwnd := WinExist("A")
    if !hwnd {
        return
    }

    if (hwnd = lastHwnd) {
        return
    }

    nowTick := A_TickCount
    nowNow := A_Now
    proc := SafeProcessName(hwnd)
    title := SafeWindowTitle(hwnd)

    if IsIgnoredProcess(proc) {
        lastHwnd := hwnd
        return
    }

    if (currentFocusProc = proc && currentFocusTitle != "" && title != "" && currentFocusTitle != title) {
        tabHopTimestamps.Push(nowTick)
    }

    RecordFocusBlock(nowNow, nowTick, "switch")

    switchTimestamps.Push(nowTick)
    lastHwnd := hwnd
    lastSwitchProc := proc
    lastSwitchTitle := title

    currentFocusProc := proc
    currentFocusTitle := title
    currentFocusStartTick := nowTick
    currentFocusStartNow := nowNow

    EnsureAppStat(proc)
    appStats[proc]["switches"] += 1

    TrimOldTimestamps()
}

UpdateUiTick() {
    global isPaused, pausedSnapshot

    if isPaused {
        metrics := CloneMetrics(pausedSnapshot)
        metrics["mode"] := "Paused"
    } else {
        metrics := CurrentMetrics()
    }

    HandleModeTransition(metrics)
    UpdateTrayTooltip(metrics)
    UpdateTrayQuickStats(metrics)
    UpdateOverlay(metrics)
    SetTrayModeIcon(metrics["mode"])
}

CurrentMetrics() {
    global switchTimestamps, wheelTimestamps, keyTimestamps, tabHopTimestamps
    global currentFocusProc, cfg, isPaused

    TrimOldTimestamps()

    switches5 := CountRecent(switchTimestamps, 5 * 60 * 1000)
    switches60 := CountRecent(switchTimestamps, 60 * 60 * 1000)
    wheel5 := CountRecent(wheelTimestamps, 5 * 60 * 1000)
    keys5 := CountRecent(keyTimestamps, 5 * 60 * 1000)
    tabHops5 := CountRecent(tabHopTimestamps, 5 * 60 * 1000)
    idleS := Floor(A_TimeIdlePhysical / 1000)

    profile := TimeOfDayProfile()
    ctx := BuildModeContext(currentFocusProc, profile)

    mode := DetermineModeCore(Map(
        "switches5", switches5,
        "wheel5", wheel5,
        "keys5", keys5,
        "idleS", idleS
    ), cfg, ctx)

    if IsTrackingSuppressed() {
        mode := isPaused ? "Paused" : "Suppressed"
    }

    return Map(
        "switches5", switches5,
        "switches60", switches60,
        "wheel5", wheel5,
        "keys5", keys5,
        "tabHops5", tabHops5,
        "idleS", idleS,
        "idleText", FormatDuration(idleS),
        "mode", mode,
        "profile", profile,
        "currentProc", currentFocusProc
    )
}

BuildModeContext(proc, profile) {
    global focusProcesses, distractionProcesses, isWorkstationLocked

    ctx := Map(
        "locked", isWorkstationLocked,
        "profile", profile,
        "tw_app_bias", 1.0,
        "focus_bias", 1.0,
        "idle_bias", 1.0,
        "wheel_bias", 1.0,
        "keys_bias", 1.0
    )

    p := StrLower(proc)
    if (p != "" && distractionProcesses.Has(p)) {
        ctx["tw_app_bias"] := 0.8
        ctx["keys_bias"] := 1.2
    }

    if (p != "" && focusProcesses.Has(p)) {
        ctx["tw_app_bias"] := 1.25
        ctx["focus_bias"] := 1.25
    }

    return ctx
}

HandleModeTransition(metrics) {
    global lastMode, modeStartNow, modeTransitionCount, timeWarpAlertSent
    global timeWarpStreakActive, timeWarpStreakStartNow
    global timeWarpStreakSamples, timeWarpStreakSwitchesSum, timeWarpStreakWheelSum, timeWarpStreakKeysSum
    global cfg

    mode := metrics["mode"]

    if (mode = "TimeWarp") {
        if !timeWarpStreakActive {
            timeWarpStreakActive := true
            timeWarpStreakStartNow := A_Now
            timeWarpStreakSamples := 0
            timeWarpStreakSwitchesSum := 0
            timeWarpStreakWheelSum := 0
            timeWarpStreakKeysSum := 0
            timeWarpAlertSent := false
        }

        timeWarpStreakSamples += 1
        timeWarpStreakSwitchesSum += metrics["switches5"]
        timeWarpStreakWheelSum += metrics["wheel5"]
        timeWarpStreakKeysSum += metrics["keys5"]

        durationS := DateDiff(A_Now, timeWarpStreakStartNow, "Seconds")
        if (!timeWarpAlertSent && durationS >= cfg["tw_alert_seconds"]) {
            Notify("TimeWarp alert", "You have been in TimeWarp for " durationS " seconds.")
            timeWarpAlertSent := true
        }
    } else if timeWarpStreakActive {
        durationS := Max(0, DateDiff(A_Now, timeWarpStreakStartNow, "Seconds"))
        if (durationS >= cfg["tw_streak_min_seconds"] && timeWarpStreakSamples > 0) {
            avgSwitch := Round(timeWarpStreakSwitchesSum / timeWarpStreakSamples, 2)
            avgWheel := Round(timeWarpStreakWheelSum / timeWarpStreakSamples, 2)
            avgKeys := Round(timeWarpStreakKeysSum / timeWarpStreakSamples, 2)
            severity := ComputeTimeWarpSeverity(avgSwitch, avgWheel, avgKeys, durationS)
            AppendDistractionStreak(timeWarpStreakStartNow, A_Now, durationS, avgSwitch, avgWheel, avgKeys, severity)
        }

        timeWarpStreakActive := false
        timeWarpAlertSent := false
    }

    if (mode != lastMode) {
        durationS := Max(0, DateDiff(A_Now, modeStartNow, "Seconds"))
        AppendModeTimeline(modeStartNow, A_Now, lastMode, durationS)
        modeTransitionCount += 1

        lastMode := mode
        modeStartNow := A_Now
    }
}

CountRecent(arr, lastMs) {
    cutoff := A_TickCount - lastMs
    count := 0
    for ts in arr {
        if (ts >= cutoff) {
            count += 1
        }
    }
    return count
}

TrimOldTimestamps() {
    global switchTimestamps, wheelTimestamps, keyTimestamps, tabHopTimestamps, switchesTrend

    maxAgeMs := 24 * 60 * 60 * 1000
    TrimArrayOlderThan(switchTimestamps, maxAgeMs)
    TrimArrayOlderThan(wheelTimestamps, maxAgeMs)
    TrimArrayOlderThan(keyTimestamps, maxAgeMs)
    TrimArrayOlderThan(tabHopTimestamps, maxAgeMs)

    while (switchesTrend.Length > 0 && switchesTrend[1]["ts"] < (A_TickCount - 30 * 60 * 1000)) {
        switchesTrend.RemoveAt(1)
    }
}

TrimArrayOlderThan(arr, maxAgeMs) {
    cutoff := A_TickCount - maxAgeMs
    while (arr.Length > 0 && arr[1] < cutoff) {
        arr.RemoveAt(1)
    }
}

SampleTrendTick() {
    global switchesTrend, switchTimestamps
    switchesTrend.Push(Map("ts", A_TickCount, "v", CountRecent(switchTimestamps, 5 * 60 * 1000)))
    TrimOldTimestamps()
}

ToggleOverlay(*) {
    global overlayVisible, overlayGui

    overlayVisible := !overlayVisible
    if overlayVisible {
        PositionOverlay()
        overlayGui.Show("NoActivate")
    } else {
        overlayGui.Hide()
    }
}

ToggleOverlayDetail(*) {
    global overlayCompact
    overlayCompact := !overlayCompact
    UpdateUiTick()
}

TogglePauseTracking(*) {
    global isPaused, pausedSnapshot, lastHwnd, currentFocusProc, currentFocusStartTick, currentFocusStartNow, currentFocusTitle

    isPaused := !isPaused

    if isPaused {
        FlushCurrentFocusBlock("pause")
        pausedSnapshot := CurrentMetrics()
        A_TrayMenu.Check("Pause Tracking")
    } else {
        A_TrayMenu.Uncheck("Pause Tracking")
        lastHwnd := 0
        current := WinExist("A")
        if current {
            currentFocusProc := SafeProcessName(current)
            currentFocusStartTick := A_TickCount
            currentFocusStartNow := A_Now
            currentFocusTitle := SafeWindowTitle(current)
        }
    }

    UpdateUiTick()
}

ResetCounters(*) {
    global switchTimestamps, wheelTimestamps, keyTimestamps, tabHopTimestamps
    global appStats, currentFocusProc, currentFocusStartTick, currentFocusStartNow, currentFocusTitle
    global switchesTrend

    switchTimestamps := []
    wheelTimestamps := []
    keyTimestamps := []
    tabHopTimestamps := []
    switchesTrend := []
    appStats := Map()

    currentFocusProc := ""
    currentFocusTitle := ""
    currentFocusStartTick := A_TickCount
    currentFocusStartNow := A_Now

    AppendLog("info", "counters_reset", Map())
    UpdateUiTick()
}

Snooze15(*) {
    SetSnoozeMinutes(15)
}

Snooze30(*) {
    SetSnoozeMinutes(30)
}

Snooze60(*) {
    SetSnoozeMinutes(60)
}

SetSnoozeMinutes(minutes) {
    global snoozeUntilTick

    snoozeUntilTick := A_TickCount + (minutes * 60 * 1000)
    FlushCurrentFocusBlock("snooze")
    AppendLog("info", "snooze_start", Map("minutes", minutes))
    UpdateUiTick()
}

CancelSnooze(*) {
    global snoozeUntilTick
    snoozeUntilTick := 0
    AppendLog("info", "snooze_cancel", Map())
    UpdateUiTick()
}

SnoozeTick() {
    global snoozeUntilTick

    if (snoozeUntilTick <= 0) {
        return
    }

    if (A_TickCount >= snoozeUntilTick) {
        snoozeUntilTick := 0
        Notify("Time Warp Indicator", "Snooze ended; tracking resumed.")
        AppendLog("info", "snooze_end", Map())
    }
}

StartFocusBlock25(*) {
    StartFocusBlock(25)
}

StartFocusBlock(minutes) {
    global focusBlockActive, focusBlockEndTick, focusBlockMinutes

    focusBlockMinutes := minutes
    focusBlockEndTick := A_TickCount + (minutes * 60 * 1000)
    focusBlockActive := true

    ResetCounters()
    Notify("Focus Block", "Started " minutes " minute focus block.")
    AppendLog("info", "focus_block_start", Map("minutes", minutes))
    UpdateUiTick()
}

CancelFocusBlock(*) {
    global focusBlockActive, focusBlockEndTick

    if !focusBlockActive {
        return
    }

    focusBlockActive := false
    focusBlockEndTick := 0
    Notify("Focus Block", "Cancelled.")
    AppendLog("info", "focus_block_cancel", Map())
    UpdateUiTick()
}

FocusBlockTick() {
    global focusBlockActive, focusBlockEndTick

    if !focusBlockActive {
        return
    }

    if (A_TickCount >= focusBlockEndTick) {
        focusBlockActive := false
        focusBlockEndTick := 0
        Notify("Focus Block", "Completed. Counters auto-reset.")
        ResetCounters()
        AppendLog("info", "focus_block_complete", Map())
    }
}

OpenSettings(*) {
    global SETTINGS_PATH
    Run(SETTINGS_PATH)
}

OpenSessionFolder(*) {
    global SESSION_DIR
    Run(SESSION_DIR)
}

ReloadSettings(*) {
    global cfg, overlayGui, overlayText, overlayCompact, overlayVisible

    oldPoll := cfg.Has("poll_interval_ms") ? cfg["poll_interval_ms"] : 1200
    LoadSettings()

    overlayCompact := cfg["overlay_compact"]
    overlayVisible := cfg["overlay_enabled"]

    if (cfg["poll_interval_ms"] != oldPoll) {
        SetTimer(PollForeground, 0)
        SetTimer(PollForeground, cfg["poll_interval_ms"])
    }

    overlayGui.SetFont("s" cfg["font_size"], cfg["font_name"])
    overlayText.Opt("c" cfg["text_color"])
    overlayGui.BackColor := cfg["overlay_bg_color"]
    try WinSetTransparent(cfg["overlay_opacity"], overlayGui.Hwnd)
    ApplyOverlayClickThrough()

    if overlayVisible {
        overlayGui.Show("NoActivate")
    } else {
        overlayGui.Hide()
    }

    AppendLog("info", "settings_reload", Map())
    UpdateUiTick()
}

ShowTopApps(*) {
    rows := BuildTopAppsRows()
    if (rows.Length = 0) {
        MsgBox("No tracked app activity yet.")
        return
    }

    out := "Top Apps`n`n"
    for idx, row in rows {
        out .= idx ". " row["proc"] " | focus " FormatDuration(Floor(row["focusMs"] / 1000)) " | switches " row["switches"] "`n"
        if (idx >= 10) {
            break
        }
    }

    MsgBox(out)
}

GenerateWeeklyDashboard(*) {
    global SUMMARY_DIR

    DirCreate(SUMMARY_DIR)

    currentWeekRows := ReadSummaryRowsLastDays(7)
    previousWeekRows := ReadSummaryRowsForRange(8, 14)

    daysCurrent := AggregateRowsByDay(currentWeekRows)
    daysPrevious := AggregateRowsByDay(previousWeekRows)

    summary := BuildWeeklySummary(daysCurrent, daysPrevious)
    recommendations := BuildWeeklyRecommendations(summary)

    html := "<html><head><meta charset='utf-8'><title>Time Warp Weekly Dashboard</title>"
    html .= "<style>body{font-family:Segoe UI,Arial,sans-serif;background:#eef3fb;color:#1f2430;margin:0;padding:18px}"
    html .= ".wrap{max-width:980px;margin:0 auto}h1{font-size:38px;margin:0 0 10px}h2{margin:22px 0 10px}.meta{color:#4f5f78;margin-bottom:16px}"
    html .= ".cards{display:flex;gap:12px;flex-wrap:wrap}.card{background:#fff;border:1px solid #d7dfec;border-radius:10px;padding:10px 12px;min-width:145px}"
    html .= "table{border-collapse:collapse;width:100%;background:#fff}.section{margin-top:18px}"
    html .= "th,td{border:1px solid #d9deea;padding:7px;text-align:left}th{background:#eef2fb}"
    html .= ".chart{background:#fff;border:1px solid #d9deea;border-radius:10px;padding:10px;display:inline-block}.chart svg{display:block}"
    html .= ".recs{background:#fff;border:1px solid #d9deea;border-radius:10px;padding:10px 12px}</style></head><body><div class='wrap'>"
    html .= "<h1>Time Warp Weekly Dashboard</h1>"
    html .= "<div class='meta'>Generated: " FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "</div>"

    html .= "<div class='cards'>"
    html .= "<div class='card'><b>Rows</b><br>" summary["rows"] "</div>"
    html .= "<div class='card'><b>Switches</b><br>" summary["switchesTotal"] "</div>"
    html .= "<div class='card'><b>TimeWarp Minutes</b><br>" summary["timeWarpMinutes"] "</div>"
    html .= "<div class='card'><b>Avg Idle</b><br>" FormatDuration(summary["avgIdleS"]) "</div>"
    html .= "<div class='card'><b>vs Prev Switches</b><br>" DeltaString(summary["switchesTotal"], summary["previousSwitchesTotal"]) "</div>"
    html .= "</div>"

    html .= "<div class='section'><h2>Switches by Day</h2><div class='chart'>" BuildDailyBarChartSvg(daysCurrent) "</div></div>"
    html .= "<div class='section'><h2>Mode Mix (minutes)</h2><div class='chart'>" BuildModeMixChartSvg(summary) "</div></div>"

    html .= "<div class='section'><h2>Daily Aggregates</h2><table><thead><tr><th>Date</th><th>Rows</th><th>Switches</th><th>TimeWarp Min</th><th>Focus Min</th><th>Idle Avg</th></tr></thead><tbody>"
    for day in daysCurrent {
        html .= "<tr><td>" day["date"] "</td><td>" day["rows"] "</td><td>" day["switches"] "</td><td>" day["timeWarpMinutes"] "</td><td>" day["focusMinutes"] "</td><td>" FormatDuration(day["avgIdleS"]) "</td></tr>"
    }
    html .= "</tbody></table></div>"

    html .= "<div class='section'><h2>Recommendations</h2><div class='recs'><ul>"
    for rec in recommendations {
        html .= "<li>" HtmlEscape(rec) "</li>"
    }
    html .= "</ul></div></div>"

    html .= "</div></body></html>"

    filePath := SUMMARY_DIR "\\weekly-dashboard.html"
    SafeDelete(filePath)
    AppendLineSafe(filePath, html)

    ExportWeeklyAggregates(daysCurrent, daysPrevious, summary)
    Run(filePath)
}

ExportAggregatesNow(*) {
    rowsToday := ReadSummaryRowsForRange(1, 1)
    ExportDailyAggregate(FormatTime(A_Now, "yyyy-MM-dd"), rowsToday)
    GenerateWeeklyDashboard()
}

MarkDistractionMoment(*) {
    global BOOKMARKS_PATH

    if !FileExist(BOOKMARKS_PATH) {
        AppendLineSafe(BOOKMARKS_PATH, "timestamp,mode,switches_5m,wheel_5m,keys_5m,idle,proc")
    }

    m := CurrentMetrics()
    row := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "," m["mode"] "," m["switches5"] "," m["wheel5"] "," m["keys5"] "," m["idleText"] "," CsvField(m["currentProc"])
    AppendLineSafe(BOOKMARKS_PATH, row)
}

ToggleRunAtStartup(*) {
    scriptPath := A_ScriptFullPath
    cmd := Format('"{1}"', scriptPath)
    key := "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"

    if IsStartupEnabled() {
        try RegDelete(key, "TimeWarpIndicator")
        Notify("Startup", "Disabled run at startup.")
        AppendLog("info", "startup_disabled", Map())
    } else {
        try RegWrite(cmd, "REG_SZ", key, "TimeWarpIndicator")
        Notify("Startup", "Enabled run at startup.")
        AppendLog("info", "startup_enabled", Map())
    }
}

IsStartupEnabled() {
    key := "HKCU\\Software\\Microsoft\\Windows\\CurrentVersion\\Run"
    try {
        v := RegRead(key, "TimeWarpIndicator")
        return v != ""
    } catch {
        return false
    }
}

ExitAppNow(*) {
    global modeStartNow, lastMode, modeTransitionCount
    FlushCurrentFocusBlock("exit")
    AppendModeTimeline(modeStartNow, A_Now, lastMode, Max(0, DateDiff(A_Now, modeStartNow, "Seconds")))
    AppendLog("info", "exit", Map("transitions", modeTransitionCount))
    try DllCall("Wtsapi32\\WTSUnRegisterSessionNotification", "ptr", A_ScriptHwnd)
    ExitApp()
}

UpdateTrayTooltip(metrics) {
    global debugEnabled, lastSwitchProc, lastSwitchTitle

    text := "Time Warp Indicator`n"
    text .= "switches_5m: " metrics["switches5"] "`n"
    text .= "switches_60m: " metrics["switches60"] "`n"
    text .= "wheel_5m: " metrics["wheel5"] "`n"
    text .= "keys_5m: " metrics["keys5"] "`n"
    text .= "tab_hops_5m: " metrics["tabHops5"] "`n"
    text .= "idle: " metrics["idleText"] "`n"
    text .= "profile: " metrics["profile"] "`n"
    text .= "mode: " metrics["mode"]

    if debugEnabled {
        text .= "`nproc: " lastSwitchProc
        text .= "`ntitle: " lastSwitchTitle
    }

    A_IconTip := text
}

Notify(title, message) {
    global cfg
    if (cfg.Has("notifications_enabled") && !cfg["notifications_enabled"]) {
        return
    }
    TrayTip(title, message)
}

UpdateTrayQuickStats(metrics) {
    global trayStatsLabel, trayFocusLabel, traySnoozeLabel

    nextStats := "Now: " ModeShort(metrics["mode"]) " | SW5 " metrics["switches5"] " | Idle " metrics["idleText"]
    if (nextStats != trayStatsLabel) {
        A_TrayMenu.Rename(trayStatsLabel, nextStats)
        trayStatsLabel := nextStats
        A_TrayMenu.Disable(trayStatsLabel)
    }

    nextFocus := "FocusBlock: " FocusBlockLabel()
    if (nextFocus != trayFocusLabel) {
        A_TrayMenu.Rename(trayFocusLabel, nextFocus)
        trayFocusLabel := nextFocus
        A_TrayMenu.Disable(trayFocusLabel)
    }

    nextSnooze := "Snooze: " SnoozeLabel()
    if (nextSnooze != traySnoozeLabel) {
        A_TrayMenu.Rename(traySnoozeLabel, nextSnooze)
        traySnoozeLabel := nextSnooze
        A_TrayMenu.Disable(traySnoozeLabel)
    }
}

SetTrayModeIcon(mode) {
    global iconMode
    if (mode = iconMode) {
        return
    }

    try {
        switch mode {
            case "Focus":
                TraySetIcon("imageres.dll", 102)
            case "TimeWarp":
                TraySetIcon("imageres.dll", 109)
            case "Paused", "Suppressed", "Away":
                TraySetIcon("imageres.dll", 101)
            default:
                TraySetIcon("shell32.dll", 44)
        }
    }
    iconMode := mode
}

UpdateOverlay(metrics) {
    global overlayText, overlayVisible, overlayCompact

    if !overlayVisible {
        return
    }

    if overlayCompact {
        text := "Mode " ModeShort(metrics["mode"]) " | SW5 " metrics["switches5"] " | Idle " metrics["idleText"]
        text .= " | FB " FocusBlockLabel()
    } else {
        text := "SW5 " metrics["switches5"] " | SW60 " metrics["switches60"] " | TabHops5 " metrics["tabHops5"] "`n"
        text .= "Wheel5 " metrics["wheel5"] " | Keys5 " metrics["keys5"] " | Idle " metrics["idleText"] "`n"
        text .= "Mode " metrics["mode"] " | Profile " metrics["profile"] " | Snooze " SnoozeLabel() "`n"
        text .= "Trend " BuildSparkline()
    }

    overlayText.Value := text
    PositionOverlay()
}

ApplyOverlayClickThrough() {
    global cfg, overlayGui
    try {
        if cfg["overlay_click_through"] {
            WinSetExStyle("+0x20", "ahk_id " overlayGui.Hwnd)
        } else {
            WinSetExStyle("-0x20", "ahk_id " overlayGui.Hwnd)
        }
    }
}

PositionOverlay() {
    global overlayGui

    MouseGetPos(&mx, &my)
    mon := MonitorFromPoint(mx, my)
    MonitorGetWorkArea(mon, &l, &t, &r, &b)

    overlayGui.Show("AutoSize NoActivate Hide")
    overlayGui.GetPos(, , &w, &h)

    margin := 12
    x := r - w - margin
    y := t + margin

    overlayGui.Show("NoActivate x" x " y" y)
}

MonitorFromPoint(x, y) {
    monCount := MonitorGetCount()
    Loop monCount {
        MonitorGet(A_Index, &l, &t, &r, &b)
        if (x >= l && x <= r && y >= t && y <= b) {
            return A_Index
        }
    }
    return MonitorGetPrimary()
}

LoadSettings() {
    global cfg, ignoredProcesses, focusProcesses, distractionProcesses, SETTINGS_PATH

    MigrateSettings()

    cfg := Map()
    cfg["settings_version"] := IniRead(SETTINGS_PATH, "general", "settings_version", 2)
    cfg["poll_interval_ms"] := IniRead(SETTINGS_PATH, "general", "poll_interval_ms", 1200)
    cfg["notifications_enabled"] := ToBool(IniRead(SETTINGS_PATH, "general", "notifications_enabled", 1))

    cfg["overlay_enabled"] := ToBool(IniRead(SETTINGS_PATH, "overlay", "enabled", 0))
    cfg["overlay_opacity"] := IniRead(SETTINGS_PATH, "overlay", "opacity", 220)
    cfg["overlay_compact"] := ToBool(IniRead(SETTINGS_PATH, "overlay", "compact", 1))
    cfg["overlay_click_through"] := ToBool(IniRead(SETTINGS_PATH, "overlay", "click_through", 0))
    cfg["font_size"] := IniRead(SETTINGS_PATH, "overlay", "font_size", 11)
    cfg["font_name"] := IniRead(SETTINGS_PATH, "overlay", "font_name", "Segoe UI")
    cfg["overlay_bg_color"] := IniRead(SETTINGS_PATH, "overlay", "bg_color", "101820")
    cfg["text_color"] := IniRead(SETTINGS_PATH, "overlay", "text_color", "DCEBFF")

    cfg["focus_switches_5m_max"] := IniRead(SETTINGS_PATH, "thresholds", "focus_switches_5m_max", 3)
    cfg["focus_idle_max_s"] := IniRead(SETTINGS_PATH, "thresholds", "focus_idle_max_s", 20)

    cfg["tw_switches_5m"] := IniRead(SETTINGS_PATH, "thresholds", "tw_switches_5m", 10)
    cfg["tw_idle_max_s"] := IniRead(SETTINGS_PATH, "thresholds", "tw_idle_max_s", 45)
    cfg["tw_wheel_5m_min"] := IniRead(SETTINGS_PATH, "thresholds", "tw_wheel_5m_min", 25)
    cfg["tw_keys_5m_max"] := IniRead(SETTINGS_PATH, "thresholds", "tw_keys_5m_max", 70)
    cfg["tw_alert_seconds"] := IniRead(SETTINGS_PATH, "thresholds", "tw_alert_seconds", 20)
    cfg["tw_streak_min_seconds"] := IniRead(SETTINGS_PATH, "thresholds", "tw_streak_min_seconds", 30)

    cfg["tw_profile_mult_morning"] := IniRead(SETTINGS_PATH, "thresholds", "tw_profile_mult_morning", 1.1)
    cfg["tw_profile_mult_day"] := IniRead(SETTINGS_PATH, "thresholds", "tw_profile_mult_day", 1.0)
    cfg["tw_profile_mult_evening"] := IniRead(SETTINGS_PATH, "thresholds", "tw_profile_mult_evening", 0.9)
    cfg["tw_profile_mult_night"] := IniRead(SETTINGS_PATH, "thresholds", "tw_profile_mult_night", 0.85)

    cfg["debug_enabled"] := ToBool(IniRead(SETTINGS_PATH, "debug", "enabled", 0))

    ignoredRaw := IniRead(SETTINGS_PATH, "tracking", "ignored_processes", "ApplicationFrameHost.exe,SearchHost.exe,StartMenuExperienceHost.exe,ShellExperienceHost.exe")
    focusRaw := IniRead(SETTINGS_PATH, "tracking", "focus_processes", "code.exe,notepad.exe,pycharm64.exe")
    distractionRaw := IniRead(SETTINGS_PATH, "tracking", "distraction_processes", "chrome.exe,msedge.exe,firefox.exe")

    ignoredProcesses := ParseProcessList(ignoredRaw)
    focusProcesses := ParseProcessList(focusRaw)
    distractionProcesses := ParseProcessList(distractionRaw)
}

MigrateSettings() {
    global SETTINGS_PATH

    version := IniRead(SETTINGS_PATH, "general", "settings_version", 1)
    if (version >= 2) {
        return
    }

    IniWrite(2, SETTINGS_PATH, "general", "settings_version")
    IniWrite(1, SETTINGS_PATH, "general", "notifications_enabled")
    IniWrite(1, SETTINGS_PATH, "overlay", "compact")
    IniWrite(0, SETTINGS_PATH, "overlay", "click_through")
    IniWrite(20, SETTINGS_PATH, "thresholds", "tw_alert_seconds")
    IniWrite(30, SETTINGS_PATH, "thresholds", "tw_streak_min_seconds")
    IniWrite(1.1, SETTINGS_PATH, "thresholds", "tw_profile_mult_morning")
    IniWrite(1.0, SETTINGS_PATH, "thresholds", "tw_profile_mult_day")
    IniWrite(0.9, SETTINGS_PATH, "thresholds", "tw_profile_mult_evening")
    IniWrite(0.85, SETTINGS_PATH, "thresholds", "tw_profile_mult_night")
    IniWrite("code.exe,notepad.exe,pycharm64.exe", SETTINGS_PATH, "tracking", "focus_processes")
    IniWrite("chrome.exe,msedge.exe,firefox.exe", SETTINGS_PATH, "tracking", "distraction_processes")
}

ParseProcessList(raw) {
    out := Map()
    for item in StrSplit(raw, ",") {
        name := Trim(StrLower(item))
        if (name != "") {
            out[name] := true
        }
    }
    return out
}

EnsureSettingsFile() {
    global SETTINGS_PATH

    if FileExist(SETTINGS_PATH) {
        return
    }

    text := "[general]`nsettings_version=2`npoll_interval_ms=1200`nnotifications_enabled=1`n`n"
    text .= "[overlay]`nenabled=0`ncompact=1`nclick_through=0`nopacity=220`nfont_size=11`nfont_name=Segoe UI`nbg_color=101820`ntext_color=DCEBFF`n`n"
    text .= "[thresholds]`nfocus_switches_5m_max=3`nfocus_idle_max_s=20`n"
    text .= "tw_switches_5m=10`ntw_idle_max_s=45`ntw_wheel_5m_min=25`ntw_keys_5m_max=70`n"
    text .= "tw_alert_seconds=20`ntw_streak_min_seconds=30`n"
    text .= "tw_profile_mult_morning=1.1`ntw_profile_mult_day=1.0`ntw_profile_mult_evening=0.9`ntw_profile_mult_night=0.85`n`n"
    text .= "[tracking]`nignored_processes=ApplicationFrameHost.exe,SearchHost.exe,StartMenuExperienceHost.exe,ShellExperienceHost.exe`n"
    text .= "focus_processes=code.exe,notepad.exe,pycharm64.exe`n"
    text .= "distraction_processes=chrome.exe,msedge.exe,firefox.exe`n`n"
    text .= "[debug]`nenabled=0`n"

    AppendLineSafe(SETTINGS_PATH, text)
}

WriteDailySummary() {
    global SUMMARY_DIR
    if IsTrackingSuppressed() {
        return
    }

    metrics := CurrentMetrics()

    dateKey := FormatTime(A_Now, "yyyy-MM-dd")
    filePath := SUMMARY_DIR "\\" dateKey ".csv"

    if !FileExist(filePath) {
        AppendLineSafe(filePath, "timestamp,mode,switches_5m,switches_60m,wheel_5m,keys_5m,tab_hops_5m,idle,proc")
    }

    row := FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss") "," metrics["mode"] "," metrics["switches5"] "," metrics["switches60"]
    row .= "," metrics["wheel5"] "," metrics["keys5"] "," metrics["tabHops5"] "," metrics["idleText"] "," CsvField(metrics["currentProc"])
    AppendLineSafe(filePath, row)

    ExportDailyAggregate(dateKey, ReadSummaryRowsForRange(1, 1))
}

ExportDailyAggregate(dateKey, rows) {
    global SUMMARY_DIR
    outCsv := SUMMARY_DIR "\\daily-aggregate-" dateKey ".csv"
    outJson := SUMMARY_DIR "\\daily-aggregate-" dateKey ".json"

    agg := AggregateRows(rows)

    SafeDelete(outCsv)
    AppendLineSafe(outCsv, "date,rows,switches_total,timewarp_minutes,focus_minutes,normal_minutes,avg_idle_s")
    AppendLineSafe(outCsv, dateKey "," agg["rows"] "," agg["switches"] "," agg["timeWarpMinutes"] "," agg["focusMinutes"] "," agg["normalMinutes"] "," agg["avgIdleS"])

    json := Format(
        '{{"date":"{1}","rows":{2},"switches_total":{3},"timewarp_minutes":{4},"focus_minutes":{5},"normal_minutes":{6},"avg_idle_s":{7}}',
        dateKey,
        agg["rows"],
        agg["switches"],
        agg["timeWarpMinutes"],
        agg["focusMinutes"],
        agg["normalMinutes"],
        agg["avgIdleS"]
    )
    SafeDelete(outJson)
    AppendLineSafe(outJson, json)
}

ReadSummaryRowsLastDays(days) {
    return ReadSummaryRowsForRange(1, days)
}

ReadSummaryRowsForRange(startDay, endDay) {
    global SUMMARY_DIR
    rows := []

    Loop endDay - startDay + 1 {
        offset := startDay + A_Index - 1
        stamp := DateAdd(A_Now, -offset + 1, "Days")
        dateKey := FormatTime(stamp, "yyyy-MM-dd")
        filePath := SUMMARY_DIR "\\" dateKey ".csv"

        if !FileExist(filePath) {
            continue
        }

        text := FileRead(filePath, "UTF-8")
        lines := StrSplit(text, "`n", "`r")
        for line in lines {
            line := Trim(line)
            if (line = "" || InStr(line, "timestamp,mode,") = 1) {
                continue
            }

            parts := StrSplit(line, ",")
            if (parts.Length < 9) {
                continue
            }

            rows.Push(Map(
                "timestamp", parts[1],
                "mode", parts[2],
                "switches5", parts[3] + 0,
                "switches60", parts[4] + 0,
                "wheel5", parts[5] + 0,
                "keys5", parts[6] + 0,
                "tabHops5", parts[7] + 0,
                "idle", parts[8],
                "proc", parts[9]
            ))
        }
    }

    return rows
}

BuildTopAppsRows() {
    global appStats, currentFocusProc, currentFocusStartTick, isPaused

    rows := []
    now := A_TickCount

    for proc, stat in appStats {
        focus := stat["focusMs"]

        if (!isPaused && proc = currentFocusProc) {
            focus += Max(0, now - currentFocusStartTick)
        }

        rows.Push(Map(
            "proc", proc,
            "focusMs", focus,
            "switches", stat["switches"]
        ))
    }

    if (rows.Length < 2) {
        return rows
    }

    loop rows.Length {
        i := A_Index
        best := i
        loop rows.Length - i {
            j := i + A_Index
            if (rows[j]["focusMs"] > rows[best]["focusMs"]) {
                best := j
            }
        }
        if (best != i) {
            tmp := rows[i]
            rows[i] := rows[best]
            rows[best] := tmp
        }
    }

    return rows
}

RecordFocusBlock(nowNow, nowTick, reason) {
    global currentFocusProc, currentFocusTitle, currentFocusStartTick, currentFocusStartNow, SESSION_DIR, appStats

    if (currentFocusProc = "") {
        return
    }

    delta := Max(0, nowTick - currentFocusStartTick)
    EnsureAppStat(currentFocusProc)
    appStats[currentFocusProc]["focusMs"] += delta

    if (delta < 1000) {
        return
    }

    dateKey := FormatTime(currentFocusStartNow, "yyyy-MM-dd")
    filePath := SESSION_DIR "\\focus-timeline-" dateKey ".csv"
    if !FileExist(filePath) {
        AppendLineSafe(filePath, "start_ts,end_ts,duration_s,proc,title,reason")
    }

    durationS := Floor(delta / 1000)
    row := FormatTime(currentFocusStartNow, "yyyy-MM-dd HH:mm:ss") "," FormatTime(nowNow, "yyyy-MM-dd HH:mm:ss") "," durationS
    row .= "," CsvField(currentFocusProc) "," CsvField(currentFocusTitle) "," reason
    AppendLineSafe(filePath, row)
}

FlushCurrentFocusBlock(reason) {
    global currentFocusStartNow, currentFocusStartTick, currentFocusProc, currentFocusTitle

    RecordFocusBlock(A_Now, A_TickCount, reason)
    currentFocusStartNow := A_Now
    currentFocusStartTick := A_TickCount
    currentFocusProc := ""
    currentFocusTitle := ""
}

AppendModeTimeline(startNow, endNow, mode, durationS) {
    global SESSION_DIR
    if (startNow = "") {
        return
    }

    dateKey := FormatTime(startNow, "yyyy-MM-dd")
    filePath := SESSION_DIR "\\mode-timeline-" dateKey ".csv"
    if !FileExist(filePath) {
        AppendLineSafe(filePath, "start_ts,end_ts,duration_s,mode")
    }

    row := FormatTime(startNow, "yyyy-MM-dd HH:mm:ss") "," FormatTime(endNow, "yyyy-MM-dd HH:mm:ss") "," durationS "," mode
    AppendLineSafe(filePath, row)
}

AppendDistractionStreak(startNow, endNow, durationS, avgSwitch, avgWheel, avgKeys, severity) {
    global SESSION_DIR
    filePath := SESSION_DIR "\\distraction-streaks.csv"
    if !FileExist(filePath) {
        AppendLineSafe(filePath, "start_ts,end_ts,duration_s,avg_switches5,avg_wheel5,avg_keys5,severity")
    }

    row := FormatTime(startNow, "yyyy-MM-dd HH:mm:ss") "," FormatTime(endNow, "yyyy-MM-dd HH:mm:ss") "," durationS
    row .= "," avgSwitch "," avgWheel "," avgKeys "," severity
    AppendLineSafe(filePath, row)
}

AppendLog(level, event, fields) {
    global LOG_DIR
    dateKey := FormatTime(A_Now, "yyyy-MM-dd")
    filePath := LOG_DIR "\\" dateKey ".jsonl"

    dq := Chr(34)
    line := "{" dq "ts" dq ":" dq JsonEscape(FormatTime(A_Now, "yyyy-MM-dd HH:mm:ss")) dq
    line .= "," dq "level" dq ":" dq JsonEscape(level) dq
    line .= "," dq "event" dq ":" dq JsonEscape(event) dq
    for k, v in fields {
        line .= "," dq JsonEscape(k) dq ":" dq JsonEscape(v) dq
    }
    line .= "}"

    AppendLineSafe(filePath, line)
}

AppendLineSafe(filePath, text) {
    f := FileOpen(filePath, "a", "UTF-8")
    if !IsObject(f) {
        return
    }

    if (SubStr(text, -1) != "`n") {
        f.Write(text "`n")
    } else {
        f.Write(text)
    }
    f.Close()
}

SafeDelete(path) {
    if FileExist(path) {
        try FileDelete(path)
    }
}

AggregateRows(rows) {
    agg := Map("rows", 0, "switches", 0, "timeWarpMinutes", 0, "focusMinutes", 0, "normalMinutes", 0, "idleSum", 0, "avgIdleS", 0)

    for row in rows {
        agg["rows"] += 1
        agg["switches"] += row["switches5"]

        if (row["mode"] = "TimeWarp") {
            agg["timeWarpMinutes"] += 1
        } else if (row["mode"] = "Focus") {
            agg["focusMinutes"] += 1
        } else {
            agg["normalMinutes"] += 1
        }

        agg["idleSum"] += IdleTextToSeconds(row["idle"])
    }

    if (agg["rows"] > 0) {
        agg["avgIdleS"] := Round(agg["idleSum"] / agg["rows"])
    }

    return agg
}

AggregateRowsByDay(rows) {
    byDay := Map()

    for row in rows {
        dateKey := SubStr(row["timestamp"], 1, 10)
        if !byDay.Has(dateKey) {
            byDay[dateKey] := Map("date", dateKey, "rows", 0, "switches", 0, "timeWarpMinutes", 0, "focusMinutes", 0, "idleSum", 0, "avgIdleS", 0)
        }

        day := byDay[dateKey]
        day["rows"] += 1
        day["switches"] += row["switches5"]
        day["idleSum"] += IdleTextToSeconds(row["idle"])

        if (row["mode"] = "TimeWarp") {
            day["timeWarpMinutes"] += 1
        } else if (row["mode"] = "Focus") {
            day["focusMinutes"] += 1
        }
    }

    out := []
    for _, day in byDay {
        if (day["rows"] > 0) {
            day["avgIdleS"] := Round(day["idleSum"] / day["rows"])
        }
        out.Push(day)
    }

    SortDayArray(out)
    return out
}

BuildWeeklySummary(daysCurrent, daysPrevious) {
    summary := Map(
        "rows", 0,
        "switchesTotal", 0,
        "previousSwitchesTotal", 0,
        "timeWarpMinutes", 0,
        "focusMinutes", 0,
        "normalMinutes", 0,
        "avgIdleS", 0,
        "idleSum", 0,
        "topApp", ""
    )

    for day in daysCurrent {
        summary["rows"] += day["rows"]
        summary["switchesTotal"] += day["switches"]
        summary["timeWarpMinutes"] += day["timeWarpMinutes"]
        summary["focusMinutes"] += day["focusMinutes"]
        summary["normalMinutes"] += Max(0, day["rows"] - day["timeWarpMinutes"] - day["focusMinutes"])
        summary["idleSum"] += day["avgIdleS"] * day["rows"]
    }

    for day in daysPrevious {
        summary["previousSwitchesTotal"] += day["switches"]
    }

    if (summary["rows"] > 0) {
        summary["avgIdleS"] := Round(summary["idleSum"] / summary["rows"])
    }

    top := BuildTopAppsRows()
    if (top.Length > 0) {
        summary["topApp"] := top[1]["proc"]
    }

    return summary
}

ExportWeeklyAggregates(daysCurrent, daysPrevious, summary) {
    global SUMMARY_DIR
    csvPath := SUMMARY_DIR "\\weekly-aggregate.csv"
    jsonPath := SUMMARY_DIR "\\weekly-aggregate.json"

    SafeDelete(csvPath)
    AppendLineSafe(csvPath, "metric,value")
    AppendLineSafe(csvPath, "rows," summary["rows"])
    AppendLineSafe(csvPath, "switches_total," summary["switchesTotal"])
    AppendLineSafe(csvPath, "previous_switches_total," summary["previousSwitchesTotal"])
    AppendLineSafe(csvPath, "timewarp_minutes," summary["timeWarpMinutes"])
    AppendLineSafe(csvPath, "focus_minutes," summary["focusMinutes"])
    AppendLineSafe(csvPath, "normal_minutes," summary["normalMinutes"])
    AppendLineSafe(csvPath, "avg_idle_s," summary["avgIdleS"])

    json := Format(
        '{{"rows":{1},"switches_total":{2},"previous_switches_total":{3},"timewarp_minutes":{4},"focus_minutes":{5},"normal_minutes":{6},"avg_idle_s":{7}}',
        summary["rows"],
        summary["switchesTotal"],
        summary["previousSwitchesTotal"],
        summary["timeWarpMinutes"],
        summary["focusMinutes"],
        summary["normalMinutes"],
        summary["avgIdleS"]
    )

    SafeDelete(jsonPath)
    AppendLineSafe(jsonPath, json)
}

BuildDailyBarChartSvg(days) {
    width := 860
    height := 280
    left := 56
    right := 24
    top := 20
    bottom := 42
    plotW := width - left - right
    plotH := height - top - bottom

    maxSwitch := 1
    for day in days {
        if (day["switches"] > maxSwitch) {
            maxSwitch := day["switches"]
        }
    }

    count := Max(1, days.Length)
    gap := (count > 12) ? 3 : 8
    barW := Floor((plotW - (count - 1) * gap) / count)
    if (barW < 6) {
        barW := 6
    }

    svg := "<svg width='" width "' height='" height "' viewBox='0 0 " width " " height "'>"

    Loop 5 {
        ratio := (A_Index - 1) / 4
        gy := top + Floor(plotH * ratio)
        gv := Round(maxSwitch * (1 - ratio))
        svg .= "<line x1='" left "' y1='" gy "' x2='" (left + plotW) "' y2='" gy "' stroke='#dfe6f4'/>"
        svg .= "<text x='8' y='" (gy + 4) "' font-size='11' fill='#6a7995'>" gv "</text>"
    }

    svg .= "<line x1='" left "' y1='" top "' x2='" left "' y2='" (top + plotH) "' stroke='#a8b8d8'/>"
    svg .= "<line x1='" left "' y1='" (top + plotH) "' x2='" (left + plotW) "' y2='" (top + plotH) "' stroke='#a8b8d8'/>"

    for idx, day in days {
        x := left + (idx - 1) * (barW + gap)
        h := (maxSwitch > 0) ? Floor((day["switches"] / maxSwitch) * plotH) : 0
        y := top + plotH - h

        svg .= "<rect x='" x "' y='" y "' width='" barW "' height='" h "' fill='#5b8ff9' rx='3'/>"

        label := SubStr(day["date"], 6)
        labelX := x + Floor(barW / 2)
        svg .= "<text x='" labelX "' y='" (top + plotH + 16) "' text-anchor='middle' font-size='11' fill='#31425d'>" label "</text>"
        svg .= "<text x='" labelX "' y='" (y - 5) "' text-anchor='middle' font-size='11' fill='#233049'>" day["switches"] "</text>"
    }

    svg .= "</svg>"
    return svg
}

BuildModeMixChartSvg(summary) {
    width := 860
    height := 180
    total := Max(1, summary["timeWarpMinutes"] + summary["focusMinutes"] + summary["normalMinutes"])

    barX := 24
    barY := 36
    barW := width - 48
    barH := 34

    wFo := Floor((summary["focusMinutes"] / total) * barW)
    wNo := Floor((summary["normalMinutes"] / total) * barW)
    wTw := barW - wFo - wNo
    if (wTw < 0) {
        wTw := 0
    }

    x := barX
    svg := "<svg width='" width "' height='" height "' viewBox='0 0 " width " " height "'>"
    svg .= "<rect x='" barX "' y='" barY "' width='" barW "' height='" barH "' fill='#e8eef8' rx='6'/>"

    svg .= "<rect x='" x "' y='" barY "' width='" wFo "' height='" barH "' fill='#3aa76d' rx='6'/>"
    if (wFo > 60) {
        svg .= "<text x='" (x + Floor(wFo/2)) "' y='" (barY + 22) "' text-anchor='middle' font-size='12' fill='#ffffff'>Focus " summary["focusMinutes"] "m</text>"
    }
    x += wFo

    svg .= "<rect x='" x "' y='" barY "' width='" wNo "' height='" barH "' fill='#5b8ff9'/>"
    if (wNo > 70) {
        svg .= "<text x='" (x + Floor(wNo/2)) "' y='" (barY + 22) "' text-anchor='middle' font-size='12' fill='#ffffff'>Normal " summary["normalMinutes"] "m</text>"
    }
    x += wNo

    svg .= "<rect x='" x "' y='" barY "' width='" wTw "' height='" barH "' fill='#dd5f4b' rx='6'/>"
    if (wTw > 80) {
        svg .= "<text x='" (x + Floor(wTw/2)) "' y='" (barY + 22) "' text-anchor='middle' font-size='12' fill='#ffffff'>TimeWarp " summary["timeWarpMinutes"] "m</text>"
    }

    ly := 110
    svg .= "<rect x='24' y='" (ly - 10) "' width='12' height='12' fill='#3aa76d'/><text x='42' y='" ly "' font-size='12' fill='#1f2430'>Focus " summary["focusMinutes"] "m</text>"
    svg .= "<rect x='220' y='" (ly - 10) "' width='12' height='12' fill='#5b8ff9'/><text x='238' y='" ly "' font-size='12' fill='#1f2430'>Normal " summary["normalMinutes"] "m</text>"
    svg .= "<rect x='430' y='" (ly - 10) "' width='12' height='12' fill='#dd5f4b'/><text x='448' y='" ly "' font-size='12' fill='#1f2430'>TimeWarp " summary["timeWarpMinutes"] "m</text>"
    svg .= "</svg>"
    return svg
}

BuildSparkline() {
    global switchesTrend

    if (switchesTrend.Length = 0) {
        return "-"
    }

    blocks := ["▁","▂","▃","▄","▅","▆","▇","█"]
    maxV := 1
    for item in switchesTrend {
        if (item["v"] > maxV) {
            maxV := item["v"]
        }
    }

    out := ""
    for item in switchesTrend {
        idx := Floor((item["v"] / maxV) * 7) + 1
        if (idx < 1) {
            idx := 1
        } else if (idx > 8) {
            idx := 8
        }
        out .= blocks[idx]
    }

    return out
}

EnsureAppStat(proc) {
    global appStats

    if !appStats.Has(proc) {
        appStats[proc] := Map("focusMs", 0, "switches", 0)
    }
}

CountWheelEvent(*) {
    global wheelTimestamps
    if IsTrackingSuppressed() {
        return
    }
    wheelTimestamps.Push(A_TickCount)
}

CountKeyEvent(*) {
    global keyTimestamps
    if IsTrackingSuppressed() {
        return
    }
    keyTimestamps.Push(A_TickCount)
}

IsTrackingSuppressed() {
    global isPaused, snoozeUntilTick, isWorkstationLocked, isSuspended

    return isPaused || isWorkstationLocked || isSuspended || (snoozeUntilTick > A_TickCount)
}

SafeProcessName(hwnd) {
    try {
        return WinGetProcessName("ahk_id " hwnd)
    } catch {
        return "unknown"
    }
}

SafeWindowTitle(hwnd) {
    try {
        return WinGetTitle("ahk_id " hwnd)
    } catch {
        return ""
    }
}

IsIgnoredProcess(procName) {
    global ignoredProcesses
    return ignoredProcesses.Has(StrLower(procName))
}

ToBool(value) {
    s := StrLower(Trim(value ""))
    return (s = "1" || s = "true" || s = "yes" || s = "on")
}

FormatDuration(totalSeconds) {
    totalSeconds := Max(0, totalSeconds)
    m := Floor(totalSeconds / 60)
    s := Mod(totalSeconds, 60)
    return Format("{:02}:{:02}", m, s)
}

IdleTextToSeconds(idleText) {
    parts := StrSplit(idleText, ":")
    if (parts.Length != 2) {
        return 0
    }
    return (parts[1] + 0) * 60 + (parts[2] + 0)
}

TimeOfDayProfile() {
    h := A_Hour + 0
    if (h >= 5 && h < 11) {
        return "morning"
    }
    if (h >= 11 && h < 18) {
        return "day"
    }
    if (h >= 18 && h < 23) {
        return "evening"
    }
    return "night"
}

SnoozeLabel() {
    global snoozeUntilTick
    if (snoozeUntilTick <= A_TickCount) {
        return "Off"
    }
    sec := Floor((snoozeUntilTick - A_TickCount) / 1000)
    return FormatDuration(sec)
}

FocusBlockLabel() {
    global focusBlockActive, focusBlockEndTick
    if !focusBlockActive {
        return "Off"
    }

    sec := Floor((focusBlockEndTick - A_TickCount) / 1000)
    if (sec < 0) {
        sec := 0
    }
    return FormatDuration(sec)
}

CsvField(text) {
    text := text ""
    dq := Chr(34)
    text := StrReplace(text, dq, dq dq)
    if InStr(text, ",") || InStr(text, dq) {
        return dq text dq
    }
    return text
}

JsonEscape(value) {
    text := value ""
    bs := Chr(92)
    dq := Chr(34)
    text := StrReplace(text, bs, bs bs)
    text := StrReplace(text, dq, bs dq)
    return text
}

HtmlEscape(text) {
    text := text ""
    text := StrReplace(text, "&", "&amp;")
    text := StrReplace(text, "<", "&lt;")
    text := StrReplace(text, ">", "&gt;")
    return text
}

DeltaString(current, previous) {
    d := current - previous
    if (d > 0) {
        return "+" d
    }
    return d
}

ModeShort(mode) {
    if (mode = "TimeWarp") {
        return "TW"
    }
    if (mode = "Focus") {
        return "Focus"
    }
    if (mode = "Normal") {
        return "Normal"
    }
    if (mode = "Paused") {
        return "Paused"
    }
    if (mode = "Suppressed") {
        return "Supp"
    }
    if (mode = "Away") {
        return "Away"
    }
    return mode
}

CloneMetrics(src) {
    dst := Map()
    for k, v in src {
        dst[k] := v
    }
    return dst
}

SortDayArray(arr) {
    if (arr.Length < 2) {
        return
    }

    Loop arr.Length {
        i := A_Index
        best := i
        Loop arr.Length - i {
            j := i + A_Index
            if (arr[j]["date"] < arr[best]["date"]) {
                best := j
            }
        }
        if (best != i) {
            tmp := arr[i]
            arr[i] := arr[best]
            arr[best] := tmp
        }
    }
}

DetermineModeCore(input, cfg, ctx) {
    if ctx["locked"] {
        return "Away"
    }

    profileMult := cfg.Has("tw_profile_mult_" ctx["profile"]) ? cfg["tw_profile_mult_" ctx["profile"]] : 1.0

    twSwitches := Max(1, Round(cfg["tw_switches_5m"] * profileMult * ctx["tw_app_bias"]))
    twIdle := Max(3, Round(cfg["tw_idle_max_s"] * ctx["idle_bias"]))
    twWheel := Max(0, Round(cfg["tw_wheel_5m_min"] * ctx["wheel_bias"]))
    twKeys := Max(1, Round(cfg["tw_keys_5m_max"] * ctx["keys_bias"]))

    focusSwitchMax := Max(0, Round(cfg["focus_switches_5m_max"] * ctx["focus_bias"]))
    focusIdleMax := Max(3, Round(cfg["focus_idle_max_s"] * ctx["idle_bias"]))

    if (input["switches5"] >= twSwitches
        && input["idleS"] <= twIdle
        && input["wheel5"] >= twWheel
        && input["keys5"] <= twKeys) {
        return "TimeWarp"
    }

    if (input["switches5"] <= focusSwitchMax && input["idleS"] <= focusIdleMax) {
        return "Focus"
    }

    return "Normal"
}

ComputeTimeWarpSeverity(avgSwitches5, avgWheel5, avgKeys5, durationS) {
    score := (avgSwitches5 * 1.7) + (avgWheel5 * 0.25) - (avgKeys5 * 0.08) + (durationS * 0.03)
    if (score < 0) {
        score := 0
    }
    return Round(score, 2)
}

BuildWeeklyRecommendations(summary) {
    recs := []

    if (summary["timeWarpMinutes"] >= 120) {
        recs.Push("High TimeWarp time this week. Try enabling focus blocks and lowering tw_switches_5m.")
    }

    if (summary["avgIdleS"] > 300) {
        recs.Push("Idle average is high. Consider pausing tracking during breaks or lock periods.")
    }

    if (summary["topApp"] != "") {
        recs.Push("Top app by focus was " summary["topApp"] ". Use per-app rules if this is distraction-heavy.")
    }

    if (summary["switchesTotal"] > summary["previousSwitchesTotal"]) {
        recs.Push("Window switching increased vs last week. Consider raising poll_interval_ms slightly.")
    }

    if (recs.Length = 0) {
        recs.Push("Patterns look stable. Keep current thresholds and monitor next week-over-week trend.")
    }

    return recs
}
