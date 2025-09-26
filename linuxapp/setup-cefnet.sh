#!/bin/bash

# Setup script to enable CefNet web browser integration
# Run this script to add full web automation capabilities

echo "Setting up CefNet integration for Cybage MIS Automation Linux..."

# Add CefNet packages to the project
cd CybageMISAutomationLinux

echo "Adding CefNet packages..."
dotnet add package CefNet --version 105.3.22248.142
dotnet add package CefNet.Avalonia --version 105.3.22248.142

# Backup current MainWindow files
echo "Backing up current files..."
cp Views/MainWindow.axaml Views/MainWindow.axaml.backup
cp Views/MainWindow.axaml.cs Views/MainWindow.axaml.cs.backup

echo "CefNet packages added successfully!"
echo ""
echo "Next steps:"
echo "1. Uncomment CefNet namespaces in MainWindow.axaml"
echo "2. Replace placeholder web view with CefNet WebView"
echo "3. Update InitializeWebView() method to use CefNet APIs"
echo "4. Test the web automation functionality"
echo ""
echo "Refer to CefNet documentation: https://github.com/CefNet/CefNet"
echo "Original implementation backup saved as *.backup files"