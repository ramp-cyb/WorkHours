# Cybage MIS Report Automation

A streamlined WPF application for automating MIS report generation and calendar visualization at Cybage Technology Group.

![Cybage Logo](cybageLogo.png)

## 🚀 Quick Start

### Prerequisites
- **.NET 8.0 SDK** (Windows)
- **Visual Studio 2022** or **VS Code** with C# extension
- **WiX Toolset v5** (for MSI installer creation)
- **Windows 10/11** (WebView2 dependency)

### Clone & Build
```bash
# Clone the repository
git clone https://github.com/ramp-cyb/WorkHours.git
cd WorkHours/CybageMISAutomation

# Build the application
dotnet build

# Run in development
dotnet run

# Create optimized release build (65MB vs 163MB)
dotnet publish -c Release -r win-x64 --self-contained true
```

### Create MSI Installer
```powershell
# Navigate to installer directory
cd Installer

# Build optimized MSI (59MB package)
.\BuildMSI.ps1 -Clean
```

## 📁 Project Structure

```
CybageMISAutomation/
├── 📁 Models/                  # Data models and configuration
│   ├── AppConfig.cs           # Application configuration
│   ├── SwipeLogEntry.cs       # Swipe log data model
│   ├── WorkHoursCalculation.cs # Work hours calculation model
│   └── DayCalendarItem.cs     # Calendar day representation
├── 📁 Services/               # Business logic and data processing
│   ├── ConfigurationService.cs # Config file management
│   └── FullReportBuilder.cs   # Calendar report generation
├── 📁 Views/                  # UI dialogs and windows
│   ├── EmployeeIdInputDialog.* # Employee ID input popup
│   ├── FullReportWindow.*     # Monthly calendar display
│   ├── LogWindow.*            # Debug logging window
│   └── MonthlyWindow.*        # Monthly report viewer
├── 📁 Installer/              # MSI creation scripts
│   ├── CybageMISAutomation.wxs # WiX installer configuration
│   ├── BuildMSI.ps1           # Automated build script
│   └── README.md              # Installer documentation
├── MainWindow.*               # Primary application interface
├── cybageLogo.png            # Company branding
└── config.json              # Runtime configuration
```

## 🔧 Key Features & Architecture

### 🎯 **Core Functionality**

1. **Automated MIS Navigation**
   - WebView2-based browser automation
   - Windows authentication integration
   - Dynamic form interaction

2. **Data Extraction & Processing**
   - Real-time swipe log parsing
   - Work hours calculation
   - Multi-day data aggregation

3. **Calendar Visualization**
   - Monthly calendar grid display
   - Color-coded work hours (Red <40h, Green 40-45h, Blue 45h+)
   - Weekly totals summary

4. **Configuration Management**
   - Persistent JSON-based settings
   - Employee ID auto-save
   - Auto-start functionality

## 🏗️ Application Flow

### 🚦 **Startup Sequence**
```mermaid
graph TD
    A[App Start] --> B[Load Configuration]
    B --> C[Initialize WebView2]
    C --> D[Employee ID Check]
    D -->|Empty| E[Show Input Dialog]
    D -->|Valid| F[Auto-start Report?]
    E --> G[Save & Continue]
    F -->|Yes| H[Run Full Automation]
    F -->|No| I[Ready for Manual Use]
```

### 📊 **Report Generation Flow**
```mermaid
graph LR
    A[Start Automation] --> B[Navigate to MIS]
    B --> C[Select Employee]
    C --> D[Extract Monthly Data]
    D --> E[Get Today's Data]
    E --> F[Get Yesterday's Data]
    F --> G[Build Calendar]
    G --> H[Display Results]
```

## 🔑 Key Methods & Functions

### **MainWindow.xaml.cs** - Primary Controller (3,429 lines)

#### 🌟 **Critical Methods**

1. **`HandleAutoStartWithEmployeeIdValidation()`** *(Lines 1883-1924)*
   ```csharp
   // Validates employee ID before auto-start
   // Shows popup dialog if ID is missing
   // Saves ID and proceeds with automation
   ```

2. **`RunFullCalendarAutomation()`** *(Lines 3046-3074)*
   ```csharp
   // Main automation orchestrator
   // Sequence: Monthly → Today → Yesterday → Calendar
   // Handles all data aggregation
   ```

3. **`SelectEmployeeFromDropdown()`** *(Lines 2065-2255)*
   ```csharp
   // Dynamic employee selection using WebView2
   // Searches dropdown options by employee ID
   // Handles MIS form interactions
   ```

4. **`ExtractSwipeDataFromTable()`** *(Lines 2758-2900)*
   ```csharp
   // Parses HTML table data into structured objects
   // Converts time strings to TimeSpan objects
   // Calculates work hours and breaks
   ```

### **FullReportBuilder.cs** - Calendar Logic

#### 🗓️ **Calendar Generation**

1. **`Build()`** *(Main method)*
   ```csharp
   // Combines monthly, today, and yesterday data
   // Creates calendar grid structure
   // Applies color coding logic
   ```

2. **Color Coding Logic**:
   - **Red**: < 40 hours/week
   - **Green**: 40-45 hours/week  
   - **Blue**: 45+ hours/week
   - **Gray**: No data available

### **ConfigurationService.cs** - Settings Management

#### ⚙️ **Configuration Handling**

1. **`LoadConfigurationAsync()`**
   ```csharp
   // Loads settings from config.json
   // Creates default config if missing
   // Handles JSON deserialization errors
   ```

2. **`SaveConfigurationAsync()`**
   ```csharp
   // Persists settings to config.json
   // Thread-safe configuration updates
   // Auto-save on employee ID changes
   ```

## 🎨 UI Components

### **Employee ID Validation**
- **Purpose**: Ensures valid employee ID before automation
- **Trigger**: Auto-start with empty employee ID
- **Flow**: Popup → Validation → Save → Continue

### **Calendar Display** *(FullReportWindow)*
- **Size**: 60% of desktop (optimized from maximized)
- **Grid**: 7 days × 6 weeks + weekly totals column
- **Styling**: Professional color scheme with rounded corners

### **Logo Integration**
- **Location**: Top-left header with light blue background
- **Format**: PNG with transparent background
- **Visibility**: Enhanced with `#BBDEFB` background container

## 📦 Build Optimizations

### **Size Reduction Achievements**
| Component | Before | After | Reduction |
|-----------|--------|--------|-----------|
| Executable | 163.9 MB | 65.3 MB | **60%** |
| MSI Package | ~162 MB | 59.1 MB | **63%** |
| File Count | 200+ files | 3 files | **99%** |

### **Optimization Techniques**
1. **Single-file publishing** with compression
2. **Debug symbol removal** in release builds
3. **Language pack exclusion** (English only)
4. **Embedded CAB** in MSI installer
5. **Resource optimization** and trimming

## ⚡ Performance Features

### **WebView2 Integration**
- **Security**: Disabled web security for internal MIS access
- **Authentication**: Automatic Windows authentication
- **Navigation**: Programmatic form interaction
- **Error Handling**: Robust timeout and retry logic

### **Async Operations**
- **Non-blocking UI**: All automation runs asynchronously
- **Progress Feedback**: Real-time status updates
- **Cancellation Support**: User can interrupt operations

## 🔐 Security Considerations

### **Current Implementation**
- ⚠️ WebView2 runs with `--disable-web-security` flag
- ⚠️ No URL validation for navigation
- ⚠️ Credentials stored in Windows authentication

### **Recommended Improvements**
- Add URL whitelist validation
- Implement secure credential storage
- Enable proper HTTPS certificate validation
- Add request/response logging for audit trails

## 🚀 Deployment

### **MSI Installer Features**
- **Company**: Cybage Technology Group
- **Installation**: Program Files structure
- **Shortcuts**: Start Menu + Desktop with logo
- **Uninstall**: Clean removal support
- **Version**: Date-based versioning (v20250929)

### **Distribution**
```powershell
# Create production build
.\BuildMSI.ps1 -Configuration Release -Clean

# Output: CybageMISAutomation-Setup-v[DATE].msi
# Size: ~59MB (single-file distribution)
```

## 🐛 Debugging & Logs

### **Log Window**
- **Toggle**: Checkbox in main interface
- **Content**: Real-time automation progress
- **Levels**: Info, Warning, Error messages
- **Persistence**: Session-based (not saved to file)

### **Common Issues**
1. **WebView2 Initialization**: Requires Windows 10+
2. **Employee ID**: Must match MIS dropdown format
3. **Network**: Requires access to cybagemis.cybage.com
4. **Authentication**: Uses current Windows user credentials

## 📈 Future Enhancements

### **Planned Features**
- [ ] Export calendar to PDF/Excel
- [ ] Email report automation
- [ ] Multi-employee batch processing
- [ ] Historical data storage
- [ ] Advanced filtering and search

### **Technical Debt**
- [ ] Split 3,429-line MainWindow class
- [ ] Replace Task.Delay with proper event handling
- [ ] Add comprehensive unit tests
- [ ] Implement proper logging framework
- [ ] Add error recovery mechanisms

## 🤝 Contributing

### **Development Setup**
1. Fork the repository
2. Create feature branch: `git checkout -b feature/my-feature`
3. Follow existing code style and patterns
4. Test MSI build process before submitting
5. Submit pull request with detailed description

### **Code Style**
- Use async/await patterns for I/O operations
- Follow Microsoft C# coding conventions  
- Add XML documentation for public methods
- Include error handling for all external operations

## 📞 Support

For technical support or questions:
- **Internal**: Contact Cybage IT Support
- **Developer**: See inline code documentation
- **Issues**: Use GitHub issue tracker

---

*Built with ❤️ for Cybage Technology Group - Streamlining MIS reporting workflows*