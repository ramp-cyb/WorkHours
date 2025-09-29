# Developer Quick Start Guide

## 🎯 5-Minute Setup

### **Clone & Run**
```bash
git clone https://github.com/ramp-cyb/WorkHours.git
cd WorkHours/CybageMISAutomation
dotnet run
```

### **First Time Setup**
1. App will prompt for Employee ID
2. Enter your MIS employee ID
3. WebView2 will initialize automatically
4. Ready for automation!

## 🔧 Development Environment

### **Required Tools**
```powershell
# Install .NET 8.0 SDK
winget install Microsoft.DotNet.SDK.8

# Install VS Code with C# extension
winget install Microsoft.VisualStudioCode
code --install-extension ms-dotnettools.csharp
```

### **Optional Tools**
```powershell
# For MSI building
winget install WiXToolset.WiX

# For advanced debugging
winget install Microsoft.VisualStudio.2022.Community
```

## 🚀 Key Development Commands

### **Build & Test**
```bash
# Debug build
dotnet build

# Run with live reload
dotnet watch run

# Release build (optimized)
dotnet publish -c Release -r win-x64 --self-contained true

# Create MSI installer
cd Installer && .\BuildMSI.ps1
```

### **Code Structure to Know**
- `MainWindow.xaml.cs` - Main app logic (focus here)
- `FullReportWindow.xaml.cs` - Calendar display
- `Services/ConfigurationService.cs` - Settings management
- `Models/` - Data structures

## 🐛 Debugging Tips

### **Common Issues**
1. **WebView2 won't load**: Restart as Administrator
2. **Employee ID not found**: Check MIS dropdown format
3. **Build fails**: Clean and rebuild solution

### **Debug Mode**
- Enable "Show Log Window" checkbox
- Watch console output for automation steps
- Check `config.json` for saved settings

## 📝 Making Changes

### **UI Modifications**
- Edit `.xaml` files for visual changes
- Modify `.xaml.cs` files for behavior changes
- Test with `dotnet run` immediately

### **Adding Features**
1. Create new method in appropriate class
2. Add UI controls if needed
3. Wire up event handlers
4. Test automation flow

### **Performance Notes**
- Avoid `Task.Delay()` - use proper async patterns
- Cache MIS data when possible
- Update UI on main thread only

---

*Happy coding! 🚀 Remember to test MSI builds before releases.*
