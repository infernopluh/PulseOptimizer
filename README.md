# Pulse — PC Optimizer (WPF, .NET 8)

A clean, safe Windows 10/11 PC optimization tool built with WPF.

## Features
- One‑Click Optimize (Temp + Browser Cache + Windows Update cache + Recycle Bin + RAM trim)
- Cleaner page with individual actions
- Startup Manager (disable/enable from registry & Startup folder safely)
- Performance page (Free RAM, kill common background bloat)
- Power Plan toggle (Balanced / High Performance)
- Admin manifest included (app runs elevated by default for full access)
- Release publish configured for single-file EXE (win-x64)


## Notes / Safety
- Uses safe cleanup paths. No registry cleaning or risky tweaks.
- Startup disable is reversible: entries are moved to `Run_DisabledByPulse` keys or `Startup_DisabledByPulse` folder.
- Power plan GUIDs: Balanced `381b4222-f694-41f0-9685-ff5bb260df2e`, High Performance `8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c`.
- RAM trim uses `EmptyWorkingSet` on non-critical processes; results vary by system.
- Logs are saved to `%LOCALAPPDATA%\PulseOptimizer\Logs`.


Enjoy!
