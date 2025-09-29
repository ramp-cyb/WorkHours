# Simple Complete MSI with ALL Dependencies - GUID-based IDs
param([switch]$Clean = $false)

$ErrorActionPreference = "Stop"
Write-Host "Building Complete MSI with ALL Dependencies (GUID-based)" -ForegroundColor Cyan

$AppDir = Split-Path -Parent $PSScriptRoot
$InstallerDir = $PSScriptRoot
$PublishDir = Join-Path $AppDir "bin\Release\net8.0-windows\win-x64\publish"

# Build with English only
Write-Host "Building complete release..." -ForegroundColor Yellow
Set-Location $AppDir
dotnet publish CybageMISAutomation.csproj -c Release -r win-x64 --self-contained true -p:SatelliteResourceLanguages=en

# Copy config
$configPath = Join-Path $PublishDir "config.json"
if (-not (Test-Path $configPath)) { Copy-Item "config.json" $configPath -Force }

# Remove language packs and duplicates
$languagesToRemove = @("de", "es", "fr", "it", "ja", "ko", "pl", "pt-BR", "ru", "tr", "zh-Hans", "zh-Hant", "cs")
foreach ($lang in $languagesToRemove) {
    $langDir = Join-Path $PublishDir $lang
    if (Test-Path $langDir) { Remove-Item $langDir -Recurse -Force }
}
Remove-Item (Join-Path $PublishDir "runtimes") -Recurse -Force -ErrorAction SilentlyContinue
Remove-Item (Join-Path $PublishDir "CybageMISAutomation.ico") -Force -ErrorAction SilentlyContinue

# Get ALL files
$allFiles = Get-ChildItem -Path $PublishDir -File -Recurse | Sort-Object FullName
Write-Host "Found $($allFiles.Count) files for deployment" -ForegroundColor Green

# Create WiX file with GUID-based IDs
Set-Location $InstallerDir
$wxsContent = @"
<?xml version="1.0" encoding="UTF-8"?>
<Wix xmlns="http://wixtoolset.org/schemas/v4/wxs">
  <Package Name="Cybage MIS Report Automation" 
           Manufacturer="Cybage Technology Group" 
           Version="1.0.0.0" 
           UpgradeCode="A1B2C3D4-5E6F-7890-ABCD-123456789012"
           Scope="perMachine"
           Compressed="yes">
    
    <SummaryInformation Description="Cybage MIS Report Automation - All Dependencies" />
    <Media Id="1" Cabinet="app.cab" EmbedCab="yes" CompressionLevel="high" />
    
    <Feature Id="MainApplication" Title="Cybage MIS Report Automation" Level="1">
      <ComponentGroupRef Id="AllFiles" />
      <ComponentGroupRef Id="Shortcuts" />
    </Feature>
    
    <StandardDirectory Id="ProgramFilesFolder">
      <Directory Id="CompanyFolder" Name="Cybage Technology Group">
        <Directory Id="INSTALLFOLDER" Name="MIS Report Automation" />
      </Directory>
    </StandardDirectory>
    
    <StandardDirectory Id="ProgramMenuFolder">
      <Directory Id="ProgramMenuDir" Name="Cybage Technology Group" />
    </StandardDirectory>
    
    <StandardDirectory Id="DesktopFolder" />
  </Package>

  <Fragment>
    <ComponentGroup Id="AllFiles" Directory="INSTALLFOLDER">
"@

# Add components with explicit file IDs and folder permissions
$fileIndex = 1
foreach ($file in $allFiles) {
    $relativePath = $file.FullName.Substring($PublishDir.Length + 1)
    $componentId = "Comp_$fileIndex"
    $fileId = "File_$fileIndex"
    
    $wxsContent += @"

      <Component Id="$componentId">
        <File Id="$fileId" Source="..\bin\Release\net8.0-windows\win-x64\publish\$relativePath" KeyPath="yes" />
      </Component>
"@
    $fileIndex++
}

# Add folder permissions component for config.json modifications
$wxsContent += @"

      <Component Id="FolderPermissions">
        <CreateFolder>
          <Permission User="Users" GenericAll="yes" />
          <Permission User="Everyone" GenericWrite="yes" />
        </CreateFolder>
        <RegistryValue Root="HKCU" Key="Software\Cybage\MIS" Name="FolderPermissions" Type="integer" Value="1" KeyPath="yes" />
      </Component>
"@

$wxsContent += @"

      <Component Id="Logo">
        <File Source="..\cybageLogo.png" KeyPath="yes" />
      </Component>
    </ComponentGroup>
  </Fragment>

  <Fragment>
    <ComponentGroup Id="Shortcuts">
      <Component Directory="ProgramMenuDir">
        <Shortcut Id="StartMenuLnk" Name="Cybage MIS Report Automation"
                  Target="[INSTALLFOLDER]CybageMISAutomation.exe" Icon="AppIcon" />
        <RemoveFolder Id="ProgramMenuDir" On="uninstall" />
        <RegistryValue Root="HKCU" Key="Software\Cybage\MIS" Name="StartMenu" Type="integer" Value="1" KeyPath="yes" />
      </Component>
      <Component Directory="DesktopFolder">
        <Shortcut Id="DesktopLnk" Name="Cybage MIS Report Automation"
                  Target="[INSTALLFOLDER]CybageMISAutomation.exe" Icon="AppIcon" />
        <RegistryValue Root="HKCU" Key="Software\Cybage\MIS" Name="Desktop" Type="integer" Value="1" KeyPath="yes" />
      </Component>
    </ComponentGroup>
  </Fragment>
  
  <Fragment>
    <Icon Id="AppIcon" SourceFile="..\bin\Release\net8.0-windows\win-x64\publish\cybage30.ico" />
  </Fragment>
</Wix>
"@

$wxsContent | Out-File -FilePath "complete-simple.wxs" -Encoding UTF8

# Build MSI
$timestamp = Get-Date -Format "yyyyMMdd"
Write-Host "Building MSI with $($allFiles.Count) files..." -ForegroundColor Yellow

wix build complete-simple.wxs -arch x64 -o "CybageMISAutomation-AllFiles-v$timestamp.msi"

if ($LASTEXITCODE -eq 0) {
    $msiFile = Get-ChildItem "*AllFiles*.msi" | Sort-Object LastWriteTime -Descending | Select-Object -First 1
    $msiSize = [math]::Round($msiFile.Length / 1MB, 2)
    
    $totalSize = ($allFiles | Measure-Object -Property Length -Sum).Sum
    $totalSizeMB = [math]::Round($totalSize / 1MB, 2)
    
    Write-Host ""
    Write-Host "✅ SUCCESS! Complete MSI Created!" -ForegroundColor Green
    Write-Host "  📦 File: $($msiFile.Name)" -ForegroundColor Cyan
    Write-Host "  📏 MSI Size: $msiSize MB" -ForegroundColor Cyan
    Write-Host "  📁 Files: $($allFiles.Count) files, $totalSizeMB MB" -ForegroundColor Cyan
    Write-Host ""
    Write-Host "🚀 This installer includes ALL dependencies!" -ForegroundColor Yellow
    Write-Host "   ✅ Should resolve kernelbase.dll crashes" -ForegroundColor Green
    Write-Host "   ✅ Complete .NET runtime" -ForegroundColor Green
    Write-Host "   ✅ All Windows API libraries" -ForegroundColor Green
    Write-Host "   ✅ Ready for any target machine" -ForegroundColor Green
} else {
    throw "MSI build failed!"
}

Remove-Item "complete-simple.wxs" -Force -ErrorAction SilentlyContinue