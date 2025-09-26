using Avalonia.Controls;
using Avalonia.Interactivity;
using System;
using System.IO;
using System.Threading.Tasks;

namespace CybageMISAutomationLinux.Views;

public partial class LogWindow : Window
{
    public LogWindow()
    {
        InitializeComponent();
        
        // Wire up event handlers
        btnClearLogs.Click += BtnClearLogs_Click;
        btnSaveLogs.Click += BtnSaveLogs_Click;
        btnClose.Click += BtnClose_Click;
    }

    public void AddLogEntry(string logEntry)
    {
        txtLogs.Text += logEntry + Environment.NewLine;
        
        // Auto-scroll to bottom
        // Note: In Avalonia, we need to access the ScrollViewer to scroll
        // This is a simplified version - in production you might want to use a proper logging control
    }

    private void BtnClearLogs_Click(object? sender, RoutedEventArgs e)
    {
        txtLogs.Text = string.Empty;
    }

    private async void BtnSaveLogs_Click(object? sender, RoutedEventArgs e)
    {
        try
        {
            var saveDialog = new SaveFileDialog
            {
                Title = "Save Logs",
                DefaultExtension = "txt",
                Filters = new() 
                { 
                    new FileDialogFilter 
                    { 
                        Name = "Text files", 
                        Extensions = { "txt" } 
                    },
                    new FileDialogFilter 
                    { 
                        Name = "All files", 
                        Extensions = { "*" } 
                    }
                }
            };

            var result = await saveDialog.ShowAsync(this);
            if (!string.IsNullOrEmpty(result))
            {
                await File.WriteAllTextAsync(result, txtLogs.Text);
                // Show success message - in a real app you might use a message box
                AddLogEntry($"Logs saved to: {result}");
            }
        }
        catch (Exception ex)
        {
            AddLogEntry($"Error saving logs: {ex.Message}");
        }
    }

    private void BtnClose_Click(object? sender, RoutedEventArgs e)
    {
        Hide();
    }
}