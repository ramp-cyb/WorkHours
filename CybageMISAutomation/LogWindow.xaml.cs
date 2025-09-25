using System.Collections.ObjectModel;
using System.IO;
using System.Windows;

namespace CybageMISAutomation
{
    public partial class LogWindow : Window
    {
        private ObservableCollection<LogEntry> _logEntries = new();

        public LogWindow()
        {
            InitializeComponent();
            dataGridLogs.ItemsSource = _logEntries;
        }

        public void AddLogEntry(string message)
        {
            var timestamp = DateTime.Now.ToString("HH:mm:ss.fff");
            var logEntry = new LogEntry
            {
                Timestamp = timestamp,
                Message = message,
                FullText = $"[{timestamp}] {message}"
            };

            // Add to collection (will automatically update UI)
            Dispatcher.Invoke(() =>
            {
                _logEntries.Add(logEntry);
                
                // Auto-scroll to bottom
                if (dataGridLogs.Items.Count > 0)
                {
                    dataGridLogs.ScrollIntoView(dataGridLogs.Items[dataGridLogs.Items.Count - 1]);
                }

                // Keep only last 500 entries to prevent memory issues
                if (_logEntries.Count > 500)
                {
                    _logEntries.RemoveAt(0);
                }
            });
        }

        public void ClearLogs()
        {
            Dispatcher.Invoke(() => _logEntries.Clear());
        }

        private void BtnClear_Click(object sender, RoutedEventArgs e)
        {
            ClearLogs();
        }

        private void BtnCopy_Click(object sender, RoutedEventArgs e)
        {
            var selectedItems = dataGridLogs.SelectedItems.Cast<LogEntry>().ToList();
            
            if (selectedItems.Any())
            {
                var selectedText = string.Join(Environment.NewLine, selectedItems.Select(x => x.FullText));
                Clipboard.SetText(selectedText);
                MessageBox.Show($"Copied {selectedItems.Count} log entries to clipboard.", "Copied", 
                    MessageBoxButton.OK, MessageBoxImage.Information);
            }
            else
            {
                // Copy all if none selected
                var allText = string.Join(Environment.NewLine, _logEntries.Select(x => x.FullText));
                if (!string.IsNullOrEmpty(allText))
                {
                    Clipboard.SetText(allText);
                    MessageBox.Show($"Copied all {_logEntries.Count} log entries to clipboard.", "Copied", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
        }

        private void BtnSave_Click(object sender, RoutedEventArgs e)
        {
            var saveDialog = new Microsoft.Win32.SaveFileDialog
            {
                Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*",
                DefaultExt = "txt",
                FileName = $"CybageMIS_Log_{DateTime.Now:yyyyMMdd_HHmmss}.txt"
            };

            if (saveDialog.ShowDialog() == true)
            {
                try
                {
                    var allText = string.Join(Environment.NewLine, _logEntries.Select(x => x.FullText));
                    File.WriteAllText(saveDialog.FileName, allText);
                    MessageBox.Show($"Log saved to {saveDialog.FileName}", "Saved", 
                        MessageBoxButton.OK, MessageBoxImage.Information);
                }
                catch (Exception ex)
                {
                    MessageBox.Show($"Error saving log: {ex.Message}", "Error", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
        }
    }

    public class LogEntry
    {
        public string Timestamp { get; set; } = string.Empty;
        public string Message { get; set; } = string.Empty;
        public string FullText { get; set; } = string.Empty;
    }
}