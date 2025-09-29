# Cybage MIS Report Automation - Deployment Guide

## 📦 Installation Package
**File:** `CybageMISAutomation-AllFiles-v20250929.msi`  
**Size:** 51.31 MB (249 files, 147MB uncompressed)  
**Type:** Complete self-contained deployment  

## 🔧 Installation Details

### Installation Location
- **Path:** `C:\Program Files\Cybage Technology Group\MIS Report Automation\`
- **Permissions:** Full write access for Users and Everyone groups
- **Reason:** Application modifies `config.json` and may create temporary files

### What Gets Installed
✅ **Main Application:** CybageMISAutomation.exe  
✅ **Complete .NET 8 Runtime:** All system dependencies  
✅ **WPF Framework:** Full presentation libraries  
✅ **WebView2 Components:** Browser engine (Microsoft Edge WebView2)  
✅ **System Libraries:** All Windows API dependencies  
✅ **Configuration:** Default config.json  
✅ **Branding Assets:** Logo and icon files  

### Shortcuts Created
- **Desktop:** Cybage MIS Report Automation (with Cybage icon)
- **Start Menu:** Programs > Cybage Technology Group > Cybage MIS Report Automation

## 🛡️ Security & Permissions

### Folder Permissions
The installer grants write permissions to the installation folder because:
- **Config Modifications:** App saves employee ID and settings to `config.json`
- **Auto-Start Feature:** Stores last used employee ID for convenience
- **Error Logging:** Creates diagnostic logs in LocalAppData (separate from install folder)

### File Modifications
**Files the application writes to:**
1. **config.json** - User settings and employee ID storage
2. **Log files** - Error diagnostics (stored in `%LocalAppData%\CybageMISAutomation\`)
3. **Export files** - CSV reports (user-selected locations via Save Dialog)

## 🚀 Deployment Steps

### 1. Prerequisites Check
- **Windows Version:** Windows 10/11 (any recent version)
- **User Rights:** Standard user account (no admin required after installation)
- **Network:** Internet connectivity for WebView2 functionality
- **Dependencies:** All included in MSI (no external requirements)

### 2. Installation Process
1. **Run MSI:** Double-click `CybageMISAutomation-AllFiles-v20250929.msi`
2. **Admin Rights:** Required during installation only (standard MSI behavior)
3. **Installation Path:** Default recommended (Program Files)
4. **Complete Installation:** All 249 files deployed automatically

### 3. First Launch
1. **Launch:** Use Desktop or Start Menu shortcut
2. **Employee ID:** App will prompt for employee ID on first run
3. **Auto-Save:** Employee ID saved for future sessions
4. **WebView2:** Browser component initializes automatically

## 🔍 Troubleshooting

### Common Issues Resolved
✅ **Kernelbase.dll crashes:** Fixed with complete dependency set  
✅ **Missing WebView2:** All components included in installer  
✅ **Permission errors:** Installation folder has full write access  
✅ **Missing .NET runtime:** Complete self-contained deployment  

### If Application Won't Start
1. **Check Event Viewer:** Windows Logs > Application
2. **Run as Admin:** Temporary test (shouldn't be needed)
3. **Reinstall:** Uninstall and reinstall MSI package
4. **Contact Support:** Provide Event Viewer error details

### Configuration Issues
- **Config Reset:** Delete `config.json` in installation folder
- **Employee ID:** Will be prompted again on next launch
- **Log Files:** Check `%LocalAppData%\CybageMISAutomation\` for diagnostic logs

## 📋 Deployment Checklist

**Pre-Installation:**
- [ ] Download MSI file to target machine
- [ ] Verify file size (should be ~51MB)
- [ ] Ensure user has local admin rights for installation

**Installation:**
- [ ] Right-click MSI → "Run as administrator"
- [ ] Follow installation wizard
- [ ] Verify shortcuts created (Desktop + Start Menu)

**Post-Installation:**
- [ ] Launch application from shortcut
- [ ] Enter employee ID when prompted
- [ ] Verify WebView2 loads correctly
- [ ] Test basic functionality (report generation)

**Verification:**
- [ ] Application starts without errors
- [ ] No crashes in Event Viewer
- [ ] Config.json saves employee ID correctly
- [ ] Export functionality works

## 🎯 Success Criteria
✅ **No Event Viewer Errors:** Application starts cleanly  
✅ **Full Functionality:** All features work as expected  
✅ **User Convenience:** Employee ID remembered between sessions  
✅ **Professional Appearance:** Proper branding and shortcuts  
✅ **Self-Contained:** No dependency installation required  

## 📞 Support Information
- **Version:** 1.0.0 (September 2025)
- **Build Date:** September 29, 2025
- **Dependencies:** All included (self-contained)
- **Compatibility:** Windows 10/11, .NET 8 Runtime (included)

---
*This deployment package resolves all previously identified kernelbase.dll crashes by including the complete dependency set required for standalone operation.*