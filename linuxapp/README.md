# Cybage MIS Automation - Linux Port

This directory contains the Linux port of the Windows WPF application using Avalonia UI framework.

## Overview

The original application is a Windows desktop app that automates web scraping of employee attendance data from the Cybage MIS system. This Linux port maintains the same functionality while replacing Windows-specific components with cross-platform alternatives.

## Key Changes Made

### UI Framework
- **Original**: Windows Presentation Foundation (WPF)
- **Linux Port**: Avalonia UI (cross-platform XAML-based framework)

### Web Browser Component
- **Original**: Microsoft WebView2 (Windows-only)
- **Linux Port**: Placeholder implementation (ready for CefNet integration)

### Project Structure
```
CybageMISAutomationLinux/
├── Models/           # Data structures (unchanged from original)
├── Services/         # Business logic (unchanged from original)  
├── Views/            # Avalonia UI windows (.axaml files)
├── App.axaml         # Application definition
├── Program.cs        # Entry point
└── config.json       # Configuration file
```

## Prerequisites

- .NET 8.0 or later
- Linux desktop environment (X11 or Wayland)
- For full web automation: CefNet or similar Chromium wrapper

## Installation & Setup

1. **Install .NET 8.0**:
   ```bash
   # Ubuntu/Debian
   sudo apt update
   sudo apt install dotnet-sdk-8.0
   
   # RHEL/CentOS/Fedora
   sudo dnf install dotnet-sdk-8.0
   ```

2. **Clone and build**:
   ```bash
   cd linuxapp/CybageMISAutomationLinux
   dotnet restore
   dotnet build
   ```

3. **Run the application**:
   ```bash
   dotnet run
   ```

## Configuration

Edit `config.json` to customize:
```json
{
  "employeeId": "1476",
  "misUrl": "https://cybagemis.cybage.com/Report%20Builder/RPTN/ReportPage.aspx",
  "showLogWindow": false,
  "showMonthly": false,
  "automationDelayMs": 2000,
  "windowTitle": "Cybage MIS Report Automation"
}
```

## Current Implementation Status

### ✅ Completed
- [x] Avalonia UI framework integration
- [x] Cross-platform project structure
- [x] Models and Services ported (unchanged)
- [x] Main window UI converted to Avalonia
- [x] Log window functionality
- [x] Configuration system
- [x] Basic application shell

### 🚧 In Progress / TODO
- [ ] Web browser integration (CefNet or alternative)
- [ ] Complete web automation logic
- [ ] Additional windows (ComparisonWindow, DataExtractionWindow, etc.)
- [ ] Linux-specific optimizations
- [ ] Package for distribution (.deb, .rpm, AppImage)

## Web Browser Integration

The current implementation includes a placeholder for web browser functionality. To enable full web automation, integrate a Chromium-based browser component:

### Option 1: CefNet (Recommended)
```xml
<PackageReference Include="CefNet" Version="105.3.22248.142" />
<PackageReference Include="CefNet.Avalonia" Version="105.3.22248.142" />
```

### Option 2: Alternative approaches
- WebView2 via Wine (limited compatibility)
- Custom Selenium WebDriver integration
- Playwright for .NET

## Architecture Notes

### Similarities to Original
- Same data models and business logic
- Identical configuration system
- Same automation workflow
- Compatible file formats (CSV export, etc.)

### Linux-specific Adaptations
- Avalonia XAML instead of WPF XAML
- Cross-platform file dialogs
- Linux-compatible web browser integration
- Proper namespace separation (`CybageMISAutomationLinux`)

## Deployment

### Development
```bash
dotnet run
```

### Production
```bash
dotnet publish -c Release -r linux-x64 --self-contained
```

### Creating Distribution Packages
- **AppImage**: Use `appimagetool`
- **Debian**: Use `dpkg-deb`
- **RPM**: Use `rpmbuild`
- **Snap**: Use `snapcraft`

## Contributing

When adding features:
1. Maintain compatibility with the original Windows version
2. Use cross-platform APIs only
3. Update both UI (.axaml) and code-behind (.axaml.cs)
4. Test on multiple Linux distributions
5. Update this README

## Known Limitations

1. **Web Browser**: Placeholder implementation - requires CefNet integration for full functionality
2. **Authentication**: May need additional configuration for Kerberos/NTLM on Linux
3. **File Paths**: Uses cross-platform path handling but may need Linux-specific adjustments

## License

Same as the original project.