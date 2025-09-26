using Avalonia.Controls;
using Avalonia.Interactivity;
using CybageMISAutomationLinux.Models;
using CybageMISAutomationLinux.Services;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Threading.Tasks;

#pragma warning disable CS8602, CS8604, CS8600, CS8601

namespace CybageMISAutomationLinux.Views;

public partial class MainWindow : Window
{
    private AppConfig _config = new AppConfig();
    private bool _isWebViewInitialized = false;
    private ObservableCollection<SwipeLogEntry> _swipeLogData = new();
    private LogWindow? _logWindow;
    
    // Data storage for full automation
    private List<SwipeLogEntry> _todayData = new List<SwipeLogEntry>();
    private List<SwipeLogEntry> _yesterdayData = new List<SwipeLogEntry>();
    private bool _isFullAutomationRunning = false;

    public MainWindow()
    {
        InitializeComponent();
        
        // Wire up button events
        btnNavigate.Click += BtnNavigate_Click;
        btnStartFullAutomation.Click += BtnStartFullAutomation_Click;
        btnReset.Click += BtnReset_Click;
        btnMonthlyReport.Click += BtnMonthlyReport_Click;
        btnFullReport.Click += BtnFullReport_Click;
        chkShowLogs.Checked += ChkShowLogs_Checked;
        chkShowLogs.Unchecked += ChkShowLogs_Unchecked;
        
        LoadConfigurationAsync();
    }
    
    private async void LoadConfigurationAsync()
    {
        try
        {
            _config = await ConfigurationService.LoadConfigurationAsync();
            
            // Apply configuration
            txtEmployeeId.Text = _config.EmployeeId;
            Title = _config.WindowTitle + " (Linux)";
            
            // Apply visibility settings
            if (_config.ShowLogWindow)
            {
                chkShowLogs.IsChecked = true;
                ShowLogWindow();
            }
            
            // Create log window with default setting (hidden initially)
            _logWindow = new LogWindow();
            
            InitializeWebView();
        }
        catch (Exception ex)
        {
            LogMessage($"Failed to load configuration: {ex.Message}");
            // Create log window with default setting (hidden)
            _logWindow = new LogWindow();
            InitializeWebView();
        }
    }

    private async void InitializeWebView()
    {
        try
        {
            UpdateStatus("Initializing Web View (Placeholder)...", 10);
            
            LogMessage("Note: This is a placeholder implementation.");
            LogMessage("For full functionality, CefNet or similar Chromium wrapper needs to be integrated.");
            LogMessage("The web automation features will require a proper browser component.");
            
            await Task.Delay(1000); // Simulate initialization
            
            _isWebViewInitialized = true;
            UpdateStatus("Web View placeholder initialized.", 20);
            LogMessage("Placeholder WebView initialization complete.");

            // Enable buttons after successful initialization
            btnStartFullAutomation.IsEnabled = true;
            if (_config.ShowMonthly)
                btnMonthlyReport.IsEnabled = true;
            btnFullReport.IsEnabled = true;
            
            // Set default URL
            txtUrlInput.Text = _config.MisUrl;
        }
        catch (Exception ex)
        {
            UpdateStatus($"Error initializing WebView: {ex.Message}", 0);
            LogMessage($"ERROR: WebView initialization failed - {ex.Message}");
        }
    }

    private void BtnNavigate_Click(object? sender, RoutedEventArgs e)
    {
        var url = txtUrlInput.Text;
        if (!string.IsNullOrEmpty(url))
        {
            LogMessage($"Navigate to: {url}");
            txtWebContent.Text = $"Simulating navigation to: {url}\n\nIn a full implementation, this would load the actual webpage using CefNet or similar browser component.";
        }
    }

    private void UpdateStatus(string message, int progress)
    {
        txtStatus.Text = message;
        progressBar.Value = progress;
        LogMessage($"Status: {message} ({progress}%)");
    }

    private void LogMessage(string message)
    {
        var timestamp = DateTime.Now.ToString("HH:mm:ss");
        var logEntry = $"[{timestamp}] {message}";
        
        // Add to log window if available
        _logWindow?.AddLogEntry(logEntry);
        
        // Also log to console for debugging
        Console.WriteLine(logEntry);
    }

    private void ShowLogWindow()
    {
        if (_logWindow != null)
        {
            _logWindow.Show();
            _logWindow.WindowState = WindowState.Normal;
        }
    }

    private void HideLogWindow()
    {
        _logWindow?.Hide();
    }

    // Event handlers for buttons
    private async void BtnStartFullAutomation_Click(object? sender, RoutedEventArgs e)
    {
        if (!_isWebViewInitialized)
        {
            LogMessage("ERROR: WebView not initialized yet");
            return;
        }

        try
        {
            LogMessage("Starting full automation (Today/Yesterday)...");
            await StartFullAutomation();
        }
        catch (Exception ex)
        {
            LogMessage($"ERROR: Full automation failed - {ex.Message}");
        }
    }

    private void BtnReset_Click(object? sender, RoutedEventArgs e)
    {
        LogMessage("Resetting application...");
        
        // Clear data
        _todayData.Clear();
        _yesterdayData.Clear();
        _isFullAutomationRunning = false;
        
        // Reset UI
        UpdateStatus("Application reset", 0);
        
        // Navigate to home page if configured
        if (!string.IsNullOrEmpty(_config.MisUrl))
        {
            txtUrlInput.Text = _config.MisUrl;
            LogMessage($"Reset to default URL: {_config.MisUrl}");
        }
    }

    private void BtnMonthlyReport_Click(object? sender, RoutedEventArgs e)
    {
        LogMessage("Opening monthly report window...");
        // TODO: Implement monthly report window
    }

    private void BtnFullReport_Click(object? sender, RoutedEventArgs e)
    {
        LogMessage("Opening full report window...");
        // TODO: Implement full report window
    }

    private void ChkShowLogs_Checked(object? sender, RoutedEventArgs e)
    {
        ShowLogWindow();
    }

    private void ChkShowLogs_Unchecked(object? sender, RoutedEventArgs e)
    {
        HideLogWindow();
    }

    // Placeholder for full automation logic
    public async Task StartFullAutomation()
    {
        if (_isFullAutomationRunning)
        {
            LogMessage("Full automation is already running!");
            return;
        }

        _isFullAutomationRunning = true;
        SetButtonsEnabled(false);

        try
        {
            LogMessage("🚀 Starting Full Automation (Today + Yesterday)...");
            UpdateStatus("Running full automation...", 0);

            // Execute Today automation
            LogMessage("📋 Phase 1: Executing Today automation...");
            await ExecuteTodayAutomation();

            // Execute Yesterday automation  
            LogMessage("📋 Phase 2: Executing Yesterday automation...");
            await ExecuteYesterdayAutomation();

            // Show comparison results
            LogMessage("📊 Phase 3: Showing comparison results...");
            ShowComparisonResults();

            LogMessage("✅ Full automation completed successfully!");
            UpdateStatus("Full automation completed successfully", 100);
        }
        catch (Exception ex)
        {
            LogMessage($"❌ Full automation failed: {ex.Message}");
            UpdateStatus($"Full automation failed: {ex.Message}", 0);
        }
        finally
        {
            _isFullAutomationRunning = false;
            SetButtonsEnabled(true);
        }
    }

    private async Task ExecuteTodayAutomation()
    {
        // Placeholder for today automation logic
        LogMessage("Executing today automation...");
        await Task.Delay(1000); // Simulate work
    }

    private async Task ExecuteYesterdayAutomation()
    {
        // Placeholder for yesterday automation logic
        LogMessage("Executing yesterday automation...");
        await Task.Delay(1000); // Simulate work
    }

    private void ShowComparisonResults()
    {
        // Placeholder for showing results
        LogMessage("Showing comparison results...");
    }

    private void SetButtonsEnabled(bool enabled)
    {
        btnStartFullAutomation.IsEnabled = enabled;
        btnMonthlyReport.IsEnabled = enabled && _config.ShowMonthly;
        btnFullReport.IsEnabled = enabled;
    }
}