#nullable enable

using System;
using System.Threading;
using System.Threading.Tasks;
using CybageMISAutomation.Models;
using CybageMISAutomation.Services;
using Gtk;
using WebKit;

namespace CybageMISAutomation.Gtk
{
    public sealed class MainWindow : Window
    {
        private readonly Entry _employeeEntry;
        private readonly Button _startButton;
        private readonly Button _fullAutomationButton;
        private readonly Button _resetButton;
        private readonly CheckButton _showLogsToggle;
        private readonly CheckButton _autoStartToggle;
        private readonly Label _statusLabel;
        private readonly ProgressBar _progressBar;
        private readonly TextView _logView;
        private readonly Frame _logFrame;
        private readonly Paned _contentPaned;
        private readonly WebView _webView;

        private readonly object _logLock = new();
        private CancellationTokenSource? _configSaveDebounce;
        private AppConfig _config = new();

        public MainWindow() : base("Cybage MIS Automation (GTK)")
        {
            SetDefaultSize(1280, 880);
            SetPosition(WindowPosition.Center);
            DeleteEvent += (_, _) => Application.Quit();

            _employeeEntry = new Entry { PlaceholderText = "Enter employee ID" };
            _startButton = new Button("Start Automation");
            _fullAutomationButton = new Button("Full Automation");
            _resetButton = new Button("Reset");
            _showLogsToggle = new CheckButton("Show Logs");
            _autoStartToggle = new CheckButton("Auto-start full report");

            var root = new Box(Orientation.Vertical, 6) { BorderWidth = 6 };
            Add(root);

            var controlRow = BuildControlRow();
            root.PackStart(controlRow, false, false, 0);

            _contentPaned = new Paned(Orientation.Vertical);
            root.PackStart(_contentPaned, true, true, 0);

            var webScrolled = new ScrolledWindow();
            _webView = new WebView();
            ConfigureWebView();
            webScrolled.Add(_webView);
            _contentPaned.Pack1(webScrolled, true, false);

            _logFrame = BuildLogFrame(out _logView);
            _contentPaned.Pack2(_logFrame, false, true);
            _contentPaned.Position = 650;

            var statusBar = BuildStatusBar(out _statusLabel, out _progressBar);
            root.PackStart(statusBar, false, false, 0);

            ShowAll();
            _logFrame.Hide();

            HookEvents();
            LoadConfigurationAsync();
        }

        private void HookEvents()
        {
            _employeeEntry.Changed += (_, _) => OnEmployeeIdChanged();
            _startButton.Clicked += (_, _) => StartNavigation();
            _fullAutomationButton.Clicked += (_, _) => StartFullAutomation();
            _resetButton.Clicked += (_, _) => ResetBrowser();
            _showLogsToggle.Toggled += (_, _) => OnLogToggleChanged();
            _autoStartToggle.Toggled += (_, _) => OnAutoStartToggleChanged();
        }

        private Box BuildControlRow()
        {
            var row = new Box(Orientation.Horizontal, 6);

            var employeeLabel = new Label("Employee ID:") { Xalign = 0f };
            row.PackStart(employeeLabel, false, false, 0);

            row.PackStart(_employeeEntry, false, false, 0);
            row.PackStart(_startButton, false, false, 0);
            row.PackStart(_fullAutomationButton, false, false, 0);
            row.PackStart(_resetButton, false, false, 0);
            row.PackStart(_showLogsToggle, false, false, 0);
            row.PackStart(_autoStartToggle, false, false, 0);

            return row;
        }

        private static Frame BuildLogFrame(out TextView logView)
        {
            var frame = new Frame("Automation Log");
            var scroll = new ScrolledWindow { ShadowType = ShadowType.EtchedIn };
            logView = new TextView
            {
                Editable = false,
                Monospace = true,
                WrapMode = WrapMode.WordChar
            };
            scroll.Add(logView);
            frame.Add(scroll);
            return frame;
        }

        private static Box BuildStatusBar(out Label statusLabel, out ProgressBar progressBar)
        {
            var bar = new Box(Orientation.Horizontal, 6);
            statusLabel = new Label("Ready") { Xalign = 0f };
            progressBar = new ProgressBar { Fraction = 0 };
            bar.PackStart(statusLabel, true, true, 0);
            bar.PackEnd(progressBar, false, false, 0);
            return bar;
        }

        private void ConfigureWebView()
        {
            _webView.Settings.EnableDeveloperExtras = true;
            _webView.LoadChanged += (_, args) => HandleWebLoadChanged(args);
            _webView.LoadFailed += (_, args) => HandleWebLoadFailed(args);
        }

        private void HandleWebLoadChanged(LoadChangedArgs args)
        {
            switch (args.LoadEvent)
            {
                case LoadEvent.Started:
                    UpdateStatus("Starting navigation...", 0.2);
                    break;
                case LoadEvent.Committed:
                    UpdateStatus("Content loading...", 0.6);
                    break;
                case LoadEvent.Finished:
                    UpdateStatus("Navigation complete.", null);
                    Log($"Page loaded: {_webView.Uri}");
                    break;
            }
        }

        private void HandleWebLoadFailed(LoadFailedArgs args)
        {
            var failingUri = string.IsNullOrWhiteSpace(_webView.Uri) ? "(unknown URI)" : _webView.Uri;
            UpdateStatus($"Navigation failed: {failingUri}", null);
            Log($"ERROR: Navigation failed for {failingUri}");
        }

        private void LoadConfigurationAsync()
        {
            UpdateStatus("Loading configuration...", 0.1);

            Task.Run(async () =>
            {
                try
                {
                    var cfg = await ConfigurationService.LoadConfigurationAsync();
                    Application.Invoke(delegate
                    {
                        _config = cfg ?? new AppConfig();

                        Title = string.IsNullOrWhiteSpace(_config.WindowTitle)
                            ? "Cybage MIS Automation (GTK)"
                            : _config.WindowTitle;

                        _employeeEntry.Text = _config.EmployeeId;
                        _showLogsToggle.Active = _config.ShowLogWindow;
                        _autoStartToggle.Active = _config.AutoStartFullReport;

                        SetLogVisibility(_config.ShowLogWindow);
                        UpdateStatus("Configuration loaded.", null);

                        if (_config.AutoStartFullReport)
                        {
                            ScheduleFullAutomationKickoff();
                        }
                    });
                }
                catch (Exception ex)
                {
                    Application.Invoke(delegate
                    {
                        Log($"Failed to load configuration: {ex.Message}");
                        UpdateStatus("Configuration load failed; using defaults.", null);
                    });
                }
            });
        }

        private void ScheduleFullAutomationKickoff()
        {
            GLib.Timeout.Add(500, () =>
            {
                StartFullAutomation();
                return false;
            });
        }

        private void OnEmployeeIdChanged()
        {
            if (_config.EmployeeId == _employeeEntry.Text)
            {
                return;
            }

            _config.EmployeeId = _employeeEntry.Text;
            ScheduleConfigurationSave();
        }

        private void OnLogToggleChanged()
        {
            var showLogs = _showLogsToggle.Active;
            SetLogVisibility(showLogs);

            if (_config.ShowLogWindow != showLogs)
            {
                _config.ShowLogWindow = showLogs;
                ScheduleConfigurationSave();
            }
        }

        private void OnAutoStartToggleChanged()
        {
            var autoStart = _autoStartToggle.Active;

            if (_config.AutoStartFullReport != autoStart)
            {
                _config.AutoStartFullReport = autoStart;
                ScheduleConfigurationSave();
            }
        }

        private void SetLogVisibility(bool isVisible)
        {
            if (isVisible)
            {
                _logFrame.ShowAll();
                var height = _contentPaned.Allocation.Height;
                var desired = height > 0 ? Math.Max(height - 220, 200) : 650;
                _contentPaned.Position = desired;
            }
            else
            {
                _logFrame.Hide();
                _contentPaned.Position = _contentPaned.Allocation.Height;
            }
        }

        private void StartNavigation()
        {
            if (string.IsNullOrWhiteSpace(_config.MisUrl))
            {
                Log("No MIS URL configured.");
                return;
            }

            Log($"Navigating to {_config.MisUrl}");
            UpdateStatus("Navigating to MIS portal...", 0.2);
            _webView.LoadUri(_config.MisUrl);
        }

        private void StartFullAutomation()
        {
            Log("Full automation is not yet implemented in the GTK head.");
            UpdateStatus("Full automation (placeholder)", null);
        }

        private void ResetBrowser()
        {
            UpdateStatus("Resetting view...", 0.1);
            _webView.LoadUri("about:blank");
            Log("Browser reset to blank page.");
        }

        private void UpdateStatus(string message, double? progress)
        {
            Application.Invoke(delegate
            {
                _statusLabel.Text = message;
                if (progress.HasValue)
                {
                    _progressBar.Show();
                    _progressBar.Fraction = Math.Clamp(progress.Value, 0d, 1d);
                }
                else
                {
                    _progressBar.Fraction = 0;
                    _progressBar.Hide();
                }
            });
        }

        private void Log(string message)
        {
            Application.Invoke(delegate
            {
                lock (_logLock)
                {
                    var buffer = _logView.Buffer;
                    var endIter = buffer.EndIter;
                    buffer.Insert(ref endIter, $"[{DateTime.Now:HH:mm:ss}] {message}{Environment.NewLine}");
                    _logView.ScrollToIter(buffer.EndIter, 0, false, 0, 0);
                }
            });
        }

        private void ScheduleConfigurationSave()
        {
            _configSaveDebounce?.Cancel();
            _configSaveDebounce = new CancellationTokenSource();
            var token = _configSaveDebounce.Token;

            Task.Run(async () =>
            {
                try
                {
                    await Task.Delay(350, token);
                    await ConfigurationService.SaveConfigurationAsync(_config);
                    Log("Configuration saved.");
                }
                catch (TaskCanceledException)
                {
                    // ignore
                }
                catch (Exception ex)
                {
                    Log($"Failed to save configuration: {ex.Message}");
                }
            }, token);
        }
    }
}
