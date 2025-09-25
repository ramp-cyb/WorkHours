# Cybage MIS Automation

A WPF application that automates interaction with the Cybage MIS system using WebView2.

## Features

- Windows Authentication integration with Cybage MIS
- Automated navigation to report pages
- Swipe log data extraction
- Real-time progress monitoring

## Requirements

- .NET 8.0 or higher
- Windows 10/11
- WebView2 Runtime (usually pre-installed)

## Getting Started

### Build and Run

```bash
dotnet restore
dotnet build
dotnet run
```

### Usage

1. **Initialize**: The application will automatically initialize WebView2 with Windows Authentication
2. **Navigate**: Click "Start Automation" to navigate to the Cybage MIS page
3. **Test**: Use "Test Page" to verify the page loaded correctly
4. **Extract**: Future versions will include automated data extraction

### Current Implementation

This initial version focuses on:
- ✅ WebView2 initialization with Windows Authentication  
- ✅ Navigation to Cybage MIS Report Page
- ✅ Page load detection and verification
- ✅ Basic logging and status updates
- 🔄 Tree expansion (next step)
- 🔄 Report generation (next step)  
- 🔄 Data extraction (next step)

### Configuration

The application uses these Windows Authentication parameters:
```
--enable-features=msIntegratedAuth 
--auth-server-allowlist=*.cybage.com 
--auth-negotiate-delegate-allowlist=*.cybage.com
```

### Logging

The application provides detailed logging in the bottom panel showing:
- WebView2 initialization status
- Navigation progress
- Page load confirmation
- Element detection results

## Next Steps

1. Implement tree node expansion for Leave Management System
2. Add "Today's and Yesterday's Swipe Log" clicking
3. Implement dropdown selection for Today/Yesterday
4. Add report generation and data extraction