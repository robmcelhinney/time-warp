# Time Warp Indicator (AutoHotkey v2)

Time Warp Indicator is a Windows tray utility that tracks window switching, idle time, scroll activity, and typing to classify your current state as `Focus`, `Normal`, or `TimeWarp`. It includes a live overlay, focus blocks, snooze controls, per-app stats, and weekly reports.

## Features

- Live mode classification: `Focus`, `Normal`, `TimeWarp`
- Tray tooltip + optional overlay
- Focus block timer (`25m`) and tracking snooze (`15/30/60m`)
- Per-app switch/focus tracking
- Session timelines and distraction streak logging
- Weekly HTML dashboard with charts and recommendations
- Startup toggle from tray menu

## Requirements

- Windows
- AutoHotkey v2

## Quick Start

1. Install AutoHotkey v2.
2. Run `run-time-warp.cmd` (recommended) or `time-warp-indicator.ahk`.
3. Use the tray icon menu for controls.

Hotkeys:
- `Ctrl+Alt+O`: Toggle overlay
- `Ctrl+Alt+M`: Mark distraction moment

## Modes

- `Focus`: low switching + low idle
- `TimeWarp`: high switching + high scroll + low typing + low idle
- `Normal`: everything else

Thresholds are configurable in `settings.ini`.

## Tray Controls

- Overlay: toggle, compact/detailed
- Tracking: pause, snooze, reset counters
- Focus block: start/cancel
- Reports: top apps, weekly dashboard, aggregate export
- System: run at startup, open settings/session folder, reload settings

## Data Files

Raw history is persisted across restarts.

- `summaries/YYYY-MM-DD.csv` minute snapshots
- `summaries/daily-aggregate-YYYY-MM-DD.{csv,json}`
- `summaries/weekly-aggregate.{csv,json}`
- `summaries/weekly-dashboard.html`
- `sessions/focus-timeline-YYYY-MM-DD.csv`
- `sessions/mode-timeline-YYYY-MM-DD.csv`
- `sessions/distraction-streaks.csv`
- `logs/YYYY-MM-DD.jsonl`
- `bookmarks.csv`

Notes:
- `Reset Counters` resets in-memory counters for the current run only.
- Aggregate export files are regenerated when you run export/report actions.

## Configuration

Edit `settings.ini`, then use tray `Reload Settings`.

Important keys:
- `poll_interval_ms`
- `notifications_enabled` (1=yes, 0=no tray pop-ups)
- Overlay: `enabled`, `compact`, `click_through`, `opacity`, font/color keys
- Thresholds: focus/timewarp thresholds, alert and streak durations
- Time-of-day multipliers: `tw_profile_mult_morning/day/evening/night`
- Process lists: `ignored_processes`, `focus_processes`, `distraction_processes`

## Troubleshooting

- Error at line 1 with `#Requires AutoHotkey v2.0` means it was launched with AutoHotkey v1.
- Fix: run `run-time-warp.cmd`, or re-associate `.ahk` files with AutoHotkey v2.
