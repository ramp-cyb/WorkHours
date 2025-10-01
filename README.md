# Cybage MIS Automation (Windows & Linux)

This repository now drives both the original Windows automation experience and a new Avalonia desktop shell so we can bring MIS tooling to Linux and macOS.

```
WorkHours.sln
├─ CybageMISAutomation/            # Existing WPF app (Windows)
├─ CybageMISAutomation.Avalonia/   # Cross-platform UI (Windows/Linux/macOS)
└─ CybageMISAutomation.Core/       # Shared models + services
```

## Prerequisites

| Platform | Requirements |
|----------|--------------|
| Windows  | .NET 8 SDK, WebView2 Runtime, PowerShell 7+, optional WiX Toolset for MSI builds |
| Linux    | .NET 8 SDK, GTK 3 runtime (`libgtk-3-0`), WebKit runtime for future browser embedding |
| macOS    | .NET 8 SDK, Xcode Command Line Tools |

> Install Avalonia templates once:
> ```powershell
> dotnet new install Avalonia.Templates
> ```

## Quick Start

### Windows (WPF automation)

```powershell
git clone https://github.com/ramp-cyb/WorkHours.git
cd WorkHours

dotnet run --project CybageMISAutomation/CybageMISAutomation.csproj
```

### Linux or macOS (Avalonia preview)

```bash
git clone https://github.com/ramp-cyb/WorkHours.git
cd WorkHours

dotnet restore
# Ensure WebKit/GTK dependencies are installed before first run
# e.g. sudo apt install libgtk-3-0 libwebkit2gtk-4.1-0

dotnet run --project CybageMISAutomation.Avalonia/CybageMISAutomation.Avalonia.csproj
```

The Avalonia UI already shares configuration (`config.json`) with the Windows build. Buttons, status output, and log capture are wired up; browser automation will be connected next using `Avalonia.WebView` (WebKit) or, if required, a Cef-based host when MIS demands Chromium-only APIs.

## Shared Core Library

- `CybageMISAutomation.Core` now owns all first-class models (`AppConfig`, `SwipeLogEntry`, `FullReportViewModel`, etc.) and services (`ConfigurationService`, `FullReportBuilder`).
- Both UI heads reference the core project, preventing logic drift between Windows and Linux flows.
- WPF retains thin placeholder files under `Models/` and `Services/` to keep legacy includes compiling while downstream consumers are migrated.

## Build & Publish

```powershell
# Build everything
dotnet build WorkHours.sln

# Publish Windows automation (self-contained)
dotnet publish CybageMISAutomation/CybageMISAutomation.csproj `
    -c Release -r win-x64 --self-contained true

# Publish Avalonia preview for Linux (framework-dependent example)
dotnet publish CybageMISAutomation.Avalonia/CybageMISAutomation.Avalonia.csproj `
    -c Release -r linux-x64
```

MSI packaging (WiX) for Windows remains under `CybageMISAutomation/Installer`. Linux/macOS packaging (AppImage, Snap, `.dmg`) will land once the automation pipeline is finished on Avalonia.

## WebView Strategy

| OS      | Current Engine          | Plan |
|---------|-------------------------|------|
| Windows | WebView2 (Chromium)     | Keep evergreen runtime + Windows integrated auth |
| Linux   | Pending                 | Target `Avalonia.WebView` (WebKit); fall back to Cef if MIS strictly needs Chromium features |
| macOS   | Pending                 | Same approach as Linux |

Linux/macOS builds will initially rely on the MIS portal’s interactive login. If Kerberos/NTLM single sign-on becomes mandatory we will revisit the embedding layer to ensure appropriate protocol support.

## Roadmap

- Integrate the existing automation workflow into Avalonia once the WebView host is finalised.
- Remove the remaining placeholder files in the WPF project after all components consume the shared core directly.
- Unify logging so both heads can pipe into a single sink (desktop window + file).
- Document distribution-specific dependency setup (Debian/Ubuntu/RHEL/macOS) after WebView smoke tests complete.

For deeper architectural details, consult `DEVELOPER_GUIDE.md`. Windows installer instructions remain in `CybageMISAutomation/Installer/README.md`.

---

*Maintainers: keep both UI heads wired to the shared library to avoid divergence while the Linux experience matures.*
