using System.Windows;
using CybageMISAutomation.Models;
using CybageMISAutomation.Services;
using CybageMISAutomation.Views;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using System.Collections.ObjectModel;
using Newtonsoft.Json;

// NOTE: Dynamic JSON parsing below triggers many nullable warnings (CS8602 etc.).
// The dynamic objects are always produced by ExecuteScriptAsync returning JSON.
// Suppressing to keep build clean while preserving concise pattern; consider refactoring to strong types later.
#pragma warning disable CS8602, CS8604, CS8600, CS8601

namespace CybageMISAutomation
{
    /// <summary>
    /// Interaction logic for MainWindow.xaml
    /// </summary>
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
            try
            {
                InitializeComponent();
                // dataGridResults grid is currently commented out in XAML; attach if reinstated
                if (this.FindName("dataGridResults") is DataGrid dg)
                    dg.ItemsSource = _swipeLogData;
                
                LoadConfigurationAsync();
            }
            catch (Exception ex)
            {
                LogMessage($"MainWindow constructor failed: {ex}");
                MessageBox.Show($"Failed to initialize main window: {ex.Message}\n\nDetailed error logged to LocalAppData\\CybageMISAutomation\\startup.log",
                    "Cybage MIS Automation - Startup Error", MessageBoxButton.OK, MessageBoxImage.Error);
                throw;
            }
        }
        
        private async void LoadConfigurationAsync()
        {
            try
            {
                _config = await ConfigurationService.LoadConfigurationAsync();
                
                // Apply configuration
                txtEmployeeId.Text = _config.EmployeeId;
                Title = _config.WindowTitle;
                
                // Apply UI visibility settings
                btnMonthlyReport.Visibility = _config.ShowMonthly ? Visibility.Visible : Visibility.Collapsed;
                btnFullReport.IsEnabled = false; // Will be enabled after WebView initialization
                
                // Set checkbox states from config
                chkShowLogs.IsChecked = _config.ShowLogWindow;
                chkAutoStartFullReport.IsChecked = _config.AutoStartFullReport;
                
                // Add event handler for employee ID changes
                txtEmployeeId.TextChanged += TxtEmployeeId_TextChanged;
                
                // Create and show/hide log window based on config
                _logWindow = new LogWindow();
                if (_config.ShowLogWindow)
                {
                    _logWindow.Show();
                }
                else
                {
                    _logWindow.Hide();
                }
                
                LogMessage($"Configuration loaded from {ConfigurationService.ConfigFilePath}: Employee ID = {_config.EmployeeId}");
                if (ConfigurationService.LastLoadError != null)
                {
                    LogMessage($"⚠ {ConfigurationService.LastLoadError}");
                    MessageBox.Show(ConfigurationService.LastLoadError, "Settings Problem",
                        MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                InitializeWebView();
            }
            catch (Exception ex)
            {
                // Create the log window first so the failure and everything after it is recorded.
                _logWindow ??= new LogWindow();
                LogMessage($"Failed to load configuration: {ex}");
                MessageBox.Show($"Configuration Error: {ex.Message}\n\nThe application will try to continue with defaults.",
                    "Configuration Warning", MessageBoxButton.OK, MessageBoxImage.Warning);

                // Continue with default config; initialize the browser exactly once.
                _config = new AppConfig();
                txtEmployeeId.TextChanged -= TxtEmployeeId_TextChanged;
                txtEmployeeId.TextChanged += TxtEmployeeId_TextChanged;
                if (!_isWebViewInitialized)
                    InitializeWebView(); // reports its own failures
            }
        }

        private async void InitializeWebView()
        {
            try
            {
                LogMessage("Starting WebView2 initialization...");
                UpdateStatus("Checking WebView2 Runtime availability...", 10);
                
                // Check if WebView2 runtime is available
                try
                {
                    var version = CoreWebView2Environment.GetAvailableBrowserVersionString();
                    LogMessage($"WebView2 Runtime found: {version}");
                }
                catch (WebView2RuntimeNotFoundException)
                {
                    string errorMsg = "WebView2 Runtime is not installed on this machine.\\n\\n" +
                                    "Please install it from:\\n" +
                                    "https://developer.microsoft.com/microsoft-edge/webview2/\\n\\n" +
                                    "Or use the 'Evergreen Bootstrapper' for automatic installation.";
                    LogMessage("WebView2 Runtime not found!");
                    MessageBox.Show(errorMsg, "Missing WebView2 Runtime", 
                        MessageBoxButton.OK, MessageBoxImage.Error);
                    return;
                }

                UpdateStatus("Initializing WebView2 with Windows Authentication...", 20);

                // Configure WebView2 environment with Windows authentication
                var options = new CoreWebView2EnvironmentOptions();
                options.AdditionalBrowserArguments = "--enable-features=msIntegratedAuth --auth-server-allowlist=*.cybage.com --auth-negotiate-delegate-allowlist=*.cybage.com --disable-web-security --disable-site-isolation-trials";
                options.AllowSingleSignOnUsingOSPrimaryAccount = true;
                
                var userDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\CybageReportViewer";
                LogMessage($"Creating WebView2 environment with user data folder: {userDataFolder}");
                
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
                LogMessage("WebView2 environment created successfully");
                
                await webView.EnsureCoreWebView2Async(env);
                LogMessage("WebView2 core initialized successfully");

                // Set up event handlers
                webView.CoreWebView2.NavigationStarting += WebView_NavigationStarting;
                webView.CoreWebView2.NavigationCompleted += WebView_NavigationCompleted;
                webView.CoreWebView2.DOMContentLoaded += WebView_DOMContentLoaded;
                webView.CoreWebView2.DocumentTitleChanged += WebView_DocumentTitleChanged;

                _isWebViewInitialized = true;
                UpdateStatus("WebView2 initialized successfully with Windows Authentication.", 20);
                LogMessage("WebView2 initialization complete. Ready for navigation.");

                // Enable buttons after successful WebView initialization
                btnStartAutomation.IsEnabled = true;
                btnStartFullAutomation.IsEnabled = true; // renamed caption only
                if (_config.ShowMonthly)
                    btnMonthlyReport.IsEnabled = true;
                btnFullReport.IsEnabled = true; // Enable Full Report after init complete
                
                // Auto start full report if configured
                if (_config.AutoStartFullReport)
                {
                    await HandleAutoStartWithEmployeeIdValidation();
                }
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error initializing WebView2: {ex.Message}", 0);
                LogMessage($"ERROR: WebView2 initialization failed - {ex.Message}");
                MessageBox.Show($"Failed to initialize WebView2: {ex.Message}", "Initialization Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private void WebView_NavigationStarting(object? sender, CoreWebView2NavigationStartingEventArgs e)
        {
            UpdateStatus($"Navigating to: {e.Uri}", 30);
            LogMessage($"Navigation started to: {e.Uri}");
        }

        private void WebView_NavigationCompleted(object? sender, CoreWebView2NavigationCompletedEventArgs e)
        {
            if (e.IsSuccess)
            {
                UpdateStatus("Navigation completed successfully.", 50);
                LogMessage("Navigation completed successfully.");
            }
            else
            {
                UpdateStatus($"Navigation failed: {e.WebErrorStatus}", 0);
                LogMessage($"ERROR: Navigation failed - {e.WebErrorStatus}");
            }
        }

        private async void WebView_DOMContentLoaded(object? sender, CoreWebView2DOMContentLoadedEventArgs e)
        {
            // During automation the running flow owns the status bar and buttons; don't report
            // intermediate pages as "loaded successfully" or re-enable manual buttons.
            if (_isAutomationBusy)
            {
                LogMessage($"DOM content loaded: {webView.CoreWebView2.Source}");
                return;
            }

            UpdateStatus("Page DOM loaded. Checking page content...", 70);
            LogMessage("DOM content loaded. Page is ready for interaction.");

            try
            {
                // Check if we're on the correct MIS page
                var title = await webView.CoreWebView2.ExecuteScriptAsync("document.title");
                var url = webView.CoreWebView2.Source;
                
                LogMessage($"Page Title: {title}");
                LogMessage($"Current URL: {url}");

                // Check for key elements that indicate the page is fully loaded
                var treeViewExists = await webView.CoreWebView2.ExecuteScriptAsync(
                    "document.getElementById('TempleteTreeView') !== null"
                );

                if (treeViewExists == "true")
                {
                    UpdateStatus("MIS Report Page loaded successfully. Tree view detected.", 100);
                    LogMessage("SUCCESS: MIS page loaded with tree view elements.");
                    btnExtractData.IsEnabled = true;
                    
                    // Also check if Leave Management System node exists to enable expand button
                    var leaveNodeExists = await webView.CoreWebView2.ExecuteScriptAsync(
                        "document.getElementById('TempleteTreeViewn21') !== null"
                    );
                    
                    if (leaveNodeExists == "true")
                    {
                        LogMessage("✓ Leave Management System node detected - ready for tree expansion");
                        btnExpandTree.IsEnabled = true;
                    }
                }
                else
                {
                    UpdateStatus("Page loaded but tree view not found. May need authentication.", 80);
                    LogMessage("WARNING: Page loaded but expected tree view elements not found.");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR checking page content: {ex.Message}");
            }
        }

        private void WebView_DocumentTitleChanged(object? sender, object e)
        {
            // Update status when document title changes
            var title = webView.CoreWebView2.DocumentTitle;
            LogMessage($"Document title changed to: {title}");
        }

        private async void BtnStartAutomation_Click(object sender, RoutedEventArgs e)
        {
            if (!_isWebViewInitialized)
            {
                MessageBox.Show("WebView2 is not initialized yet. Please wait.", "Not Ready", 
                    MessageBoxButton.OK, MessageBoxImage.Warning);
                return;
            }

            try
            {
                btnStartAutomation.IsEnabled = false;
                btnExtractData.IsEnabled = false;
                
                UpdateStatus("Starting navigation to Cybage MIS...", 25);
                LogMessage("Starting automation process...");
                LogMessage($"Target URL: {_config.MisUrl}");

                // Navigate to the MIS page
                webView.CoreWebView2.Navigate(_config.MisUrl);
            }
            catch (Exception ex)
            {
                UpdateStatus($"Error during navigation: {ex.Message}", 0);
                LogMessage($"ERROR: Navigation failed - {ex.Message}");
                MessageBox.Show($"Navigation failed: {ex.Message}", "Navigation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
                btnStartAutomation.IsEnabled = true;
            }
        }



        private void UpdateStatus(string message, int progress)
        {
            txtStatus.Text = message;
            progressBar.Value = progress;
            LogMessage($"STATUS: {message}");
        }

        private void LogMessage(string message)
        {
            _logWindow?.AddLogEntry(message);
        }

        /// <summary>
        /// ExecuteScriptAsync returns the script's value JSON-encoded. Our scripts return
        /// JSON.stringify(...), so the result is a JSON string literal wrapping JSON text.
        /// Decode it properly (handles quotes, backslashes, newlines, unicode) instead of
        /// trimming quotes by hand. A null/undefined result becomes a failure object so
        /// callers reading .success never hit a null reference.
        /// </summary>
        private static string UnwrapScriptJson(string? scriptResult)
        {
            const string noResult = "{\"success\":false,\"error\":\"Script returned no result\"}";
            if (string.IsNullOrWhiteSpace(scriptResult)) return noResult;
            var trimmed = scriptResult.Trim();
            if (trimmed == "null" || trimmed == "undefined") return noResult;
            if (trimmed.StartsWith("\""))
            {
                var inner = JsonConvert.DeserializeObject<string>(trimmed);
                return string.IsNullOrWhiteSpace(inner) ? noResult : inner;
            }
            return trimmed; // script returned an object directly
        }

        private void BtnShowLogs_Click(object sender, RoutedEventArgs e)
        {
            if (_logWindow != null)
            {
                _logWindow.Show();
                _logWindow.WindowState = WindowState.Normal;
                _logWindow.Activate();
            }
        }

        private async void BtnTestPage_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                UpdateStatus("Testing page interaction...", 90);
                LogMessage("Testing JavaScript execution on loaded page...");

                // Test JavaScript execution
                var pageInfoJson = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    JSON.stringify({
                        title: document.title,
                        url: window.location.href,
                        hasTreeView: document.getElementById('TempleteTreeView') !== null,
                        treeViewVisible: document.getElementById('TempleteTreeView') ? 
                            window.getComputedStyle(document.getElementById('TempleteTreeView')).display !== 'none' : false,
                        leaveManagementNode: document.getElementById('TempleteTreeViewn21') !== null
                    })
                ");

                LogMessage($"Raw Page Info JSON: {pageInfoJson}");

                // Parse the JSON string (WebView2 returns JSON as a quoted string)
                var cleanJson = UnwrapScriptJson(pageInfoJson);
                var info = JsonConvert.DeserializeObject<dynamic>(cleanJson);
                
                LogMessage($"Parsed Page Info: Title='{info.title}', HasTreeView={info.hasTreeView}, LeaveNode={info.leaveManagementNode}");

                if ((bool)info.hasTreeView == true)
                {
                    UpdateStatus("Page verification successful. Ready for automation steps.", 100);
                    LogMessage("SUCCESS: Page has required elements for automation.");
                    
                    if ((bool)info.leaveManagementNode == true)
                    {
                        LogMessage("✓ Leave Management System node found (ID: TempleteTreeViewn21)");
                        btnExpandTree.IsEnabled = true;
                    }
                    else
                    {
                        LogMessage("⚠ Leave Management System node not found");
                    }
                    
                    if ((bool)info.treeViewVisible == true)
                    {
                        LogMessage("✓ Tree view is visible on page");
                    }
                    else
                    {
                        LogMessage("⚠ Tree view exists but may not be visible");
                    }
                }
                else
                {
                    LogMessage("✗ Required tree view elements not found on page");
                    UpdateStatus("Page missing required elements", 50);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Page verification failed - {ex.Message}");
                UpdateStatus("Page verification failed", 0);
            }
        }

        private async void BtnExpandTree_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnExpandTree.IsEnabled = false;
                UpdateStatus("Expanding Leave Management System tree node...", 20);
                LogMessage("Starting tree expansion process...");

                // First, check if the node is already expanded
                var isExpandedResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var node = document.getElementById('TempleteTreeViewn21');
                        if (!node) return JSON.stringify({success: false, error: 'Node not found'});
                        
                        var img = node.querySelector('img');
                        var isExpanded = img && img.src.includes('minus');
                        
                        return JSON.stringify({
                            success: true, 
                            isExpanded: isExpanded,
                            imgSrc: img ? img.src : 'no image',
                            nodeText: node.textContent || 'no text'
                        });
                    })()
                ");

                var expandCheck = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(isExpandedResult));
                LogMessage($"Tree node check: Success={expandCheck.success}, IsExpanded={expandCheck.isExpanded}");

                if (!(bool)expandCheck.success)
                {
                    LogMessage($"ERROR: {expandCheck.error}");
                    UpdateStatus("Tree expansion failed - node not found", 0);
                    return;
                }

                if ((bool)expandCheck.isExpanded)
                {
                    LogMessage("✓ Leave Management System tree is already expanded");
                    UpdateStatus("Tree already expanded. Looking for swipe log option...", 50);
                    await CheckForSwipeLogOption();
                }
                else
                {
                    LogMessage("🔄 Expanding Leave Management System tree node...");
                    
                    // Click the expand icon to expand the tree
                    var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                        (function() {
                            try {
                                var node = document.getElementById('TempleteTreeViewn21');
                                if (!node) return JSON.stringify({success: false, error: 'Node not found'});
                                
                                // Simulate the click event
                                node.click();
                                
                                return JSON.stringify({success: true, message: 'Click executed'});
                            } catch (ex) {
                                return JSON.stringify({success: false, error: ex.message});
                            }
                        })()
                    ");

                    var clickResponse = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(clickResult));
                    LogMessage($"Click result: Success={clickResponse.success}");

                    if ((bool)clickResponse.success)
                    {
                        LogMessage("✓ Tree node clicked successfully");
                        UpdateStatus("Waiting for tree expansion to complete...", 60);
                        
                        // Wait a moment for the tree to expand
                        await Task.Delay(2000);
                        
                        // Verify expansion occurred and look for swipe log option
                        await CheckForSwipeLogOption();
                    }
                    else
                    {
                        LogMessage($"ERROR clicking tree node: {clickResponse.error}");
                        UpdateStatus("Tree expansion failed", 0);
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Tree expansion failed - {ex.Message}");
                UpdateStatus("Tree expansion error", 0);
            }
            finally
            {
                btnExpandTree.IsEnabled = true;
            }
        }

        private async Task CheckForSwipeLogOption()
        {
            try
            {
                LogMessage("🔍 Looking for 'Today's and Yesterday's Swipe Log' option...");

                var swipeLogCheck = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var swipeLogLink = document.getElementById('TempleteTreeViewt32');
                        if (swipeLogLink) {
                            return JSON.stringify({
                                found: true,
                                text: swipeLogLink.textContent || swipeLogLink.innerText,
                                visible: window.getComputedStyle(swipeLogLink).display !== 'none',
                                href: swipeLogLink.getAttribute('href') || 'no href'
                            });
                        } else {
                            // Try to find it by text content
                            var allLinks = document.querySelectorAll('a');
                            for (var i = 0; i < allLinks.length; i++) {
                                var link = allLinks[i];
                                var text = link.textContent || link.innerText;
                                if (text && text.toLowerCase().includes('swipe log')) {
                                    return JSON.stringify({
                                        found: true,
                                        text: text,
                                        visible: window.getComputedStyle(link).display !== 'none',
                                        id: link.id || 'no id',
                                        href: link.getAttribute('href') || 'no href'
                                    });
                                }
                            }
                            return JSON.stringify({found: false});
                        }
                    })()
                ");

                var swipeLogInfo = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(swipeLogCheck));

                if ((bool)swipeLogInfo.found)
                {
                    LogMessage($"✓ Swipe Log option found: '{swipeLogInfo.text}'");
                    LogMessage($"  Visible: {swipeLogInfo.visible}");
                    
                    UpdateStatus("Tree expanded successfully. Swipe Log option available.", 100);
                    
                    // Enable the swipe log click button
                    btnClickSwipeLog.IsEnabled = true;
                }
                else
                {
                    LogMessage("⚠ Swipe Log option not found after tree expansion");
                    UpdateStatus("Tree expanded but Swipe Log option not found", 75);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR checking for swipe log option: {ex.Message}");
            }
        }

        private async void BtnClickSwipeLog_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnClickSwipeLog.IsEnabled = false;
                UpdateStatus("Clicking 'Today's and Yesterday's Swipe Log' link...", 30);
                LogMessage("Starting swipe log link click process...");

                // First verify the swipe log link is still available and get its details
                    // Locate and click the swipe log link in a single script execution
                    var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                        (function() {
                            var result = {
                                success: false,
                                found: false,
                                clicked: false,
                                error: null,
                                details: null
                            };

                            try {
                                var anchors = document.querySelectorAll('a');
                                var targetText = ""today's and yesterday's swipe log"";

                                for (var i = 0; i < anchors.length; i++) {
                                    var text = (anchors[i].textContent || anchors[i].innerText || '').trim();
                                    var id = anchors[i].id || '';

                                    if (text && text.toLowerCase() === targetText && id.startsWith('TempleteTreeView')) {
                                        result.found = true;
                                        result.details = {
                                            id: id || 'no-id',
                                            text: text,
                                            href: anchors[i].getAttribute('href') || '',
                                            onclick: anchors[i].getAttribute('onclick') || '',
                                            visible: window.getComputedStyle(anchors[i]).display !== 'none',
                                            className: anchors[i].className || ''
                                        };

                                        anchors[i].click();
                                        result.clicked = true;
                                        result.success = true;
                                        return JSON.stringify(result);
                                    }
                                }

                                result.error = 'Link not found';
                                return JSON.stringify(result);
                            } catch (ex) {
                                result.error = ex.message || ex.toString();
                                return JSON.stringify(result);
                            }
                        })()
                    ");

                    var clickResponse = JsonConvert.DeserializeObject<dynamic>(
                        UnwrapScriptJson(clickResult));

                    if ((bool)clickResponse.success && (bool)clickResponse.clicked)
                    {
                        if (clickResponse.details != null)
                        {
                            LogMessage($"✓ Swipe Log link clicked successfully: '{clickResponse.details.text}'");
                            LogMessage($"  ID: {clickResponse.details.id}, Class: {clickResponse.details.className}, OnClick: {clickResponse.details.onclick}");
                        }
                        else
                        {
                            LogMessage("✓ Swipe Log link clicked successfully (no additional details)");
                        }

                        UpdateStatus("Waiting for report page to load...", 60);

                        // Wait for the page to navigate/change after the postback
                        await Task.Delay(3000);

                        // Verify we've navigated to the report configuration page
                        await VerifyReportPage();
                    }
                    else
                    {
                        var errorMsg = clickResponse.error?.ToString() ?? "Unknown error";
                        LogMessage($"ERROR clicking swipe log link: {errorMsg}");
                        UpdateStatus($"Swipe log click failed: {errorMsg}", 0);
                        throw new Exception($"Swipe log click failed: {errorMsg}");
                    }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Swipe log click failed - {ex.Message}");
                UpdateStatus("Swipe log click error", 0);
            }
            finally
            {
                btnClickSwipeLog.IsEnabled = true;
            }
        }

        private async Task VerifyReportPage()
        {
            try
            {
                LogMessage("🔍 Verifying report configuration page loaded...");

                var pageVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        // Look for key elements that indicate we're on the report page
                        var title = document.title;
                        var wizardPanel = document.getElementById('WizardPanel');
                        var employeeDropdown = document.getElementById('EmployeeIDDropDownList7322');
                        var dayDropdown = document.getElementById('DayDropDownList8665');
                        var generateButton = document.getElementById('ViewReportImageButton');
                        
                        return JSON.stringify({
                            title: title,
                            hasWizardPanel: wizardPanel !== null,
                            hasEmployeeDropdown: employeeDropdown !== null,
                            hasDayDropdown: dayDropdown !== null,
                            hasGenerateButton: generateButton !== null,
                            url: window.location.href
                        });
                    })()
                ");

                var pageInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(pageVerificationResult));

                LogMessage($"Page verification - Title: '{pageInfo.title}'");
                LogMessage($"  Has WizardPanel: {pageInfo.hasWizardPanel}");
                LogMessage($"  Has Employee Dropdown: {pageInfo.hasEmployeeDropdown}");
                LogMessage($"  Has Day Dropdown: {pageInfo.hasDayDropdown}");
                LogMessage($"  Has Generate Button: {pageInfo.hasGenerateButton}");

                bool isReportPage = (bool)pageInfo.hasWizardPanel && 
                                   (bool)pageInfo.hasEmployeeDropdown && 
                                   (bool)pageInfo.hasDayDropdown && 
                                   (bool)pageInfo.hasGenerateButton;

                if (isReportPage)
                {
                    LogMessage("✅ SUCCESS: Report configuration page loaded successfully!");
                    UpdateStatus("Report page ready. Dropdown selection available.", 100);
                    
                    // Enable report generation button
                    btnGenerateReport.IsEnabled = true;
                }
                else
                {
                    LogMessage("⚠ Report page elements not fully loaded or incorrect page");
                    UpdateStatus("Page loaded but missing expected elements", 75);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR verifying report page: {ex.Message}");
            }
        }

        private async void BtnGenerateReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                btnGenerateReport.IsEnabled = false;
                UpdateStatus("Setting up report parameters and generating report...", 40);
                LogMessage("Starting report generation process...");

                // Select the requested employee before setting the day
                await SelectEmployeeInSwipeLogReport();

                // Get the selected report type from UI
                var selectedItem = cmbReportType.SelectedItem as ComboBoxItem;
                var reportTypeValue = selectedItem?.Tag?.ToString() ?? "1"; // Default to Today
                var reportTypeName = selectedItem?.Content?.ToString() ?? "Today";

                LogMessage($"Selected report type: {reportTypeName} (Value: {reportTypeValue})");

                // First, set the day dropdown selection
                var dropdownResult = await webView.CoreWebView2.ExecuteScriptAsync($@"
                    (function() {{
                        try {{
                            var dayDropdown = document.getElementById('DayDropDownList8665');
                            if (!dayDropdown) {{
                                return JSON.stringify({{success: false, error: 'Day dropdown not found'}});
                            }}

                            // Check current selection
                            var currentValue = dayDropdown.value;
                            var targetValue = '{reportTypeValue}'; // 1 = Today, 0 = Yesterday
                            
                            // Set the dropdown value
                            dayDropdown.value = targetValue;
                            
                            // Trigger change event to ensure any JavaScript handlers are called
                            var changeEvent = new Event('change', {{ bubbles: true }});
                            dayDropdown.dispatchEvent(changeEvent);

                            // Verify the selection was set
                            var finalValue = dayDropdown.value;
                            var selectedText = dayDropdown.options[dayDropdown.selectedIndex].text;

                            return JSON.stringify({{
                                success: true,
                                previousValue: currentValue,
                                targetValue: targetValue,
                                finalValue: finalValue,
                                selectedText: selectedText,
                                selectionChanged: currentValue !== finalValue
                            }});
                        }} catch (ex) {{
                            return JSON.stringify({{success: false, error: ex.message}});
                        }}
                    }})()
                ");

                var dropdownResponse = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(dropdownResult));

                if ((bool)dropdownResponse.success)
                {
                    LogMessage($"✓ Day dropdown set successfully: '{dropdownResponse.selectedText}'");
                    LogMessage($"  Previous: {dropdownResponse.previousValue} -> New: {dropdownResponse.finalValue}");
                    
                    if ((bool)dropdownResponse.selectionChanged)
                    {
                        LogMessage("✓ Dropdown selection was changed");
                    }
                    else
                    {
                        LogMessage("ℹ Dropdown was already set to correct value");
                    }

                    UpdateStatus("Dropdown set. Clicking Generate Report button...", 70);

                    // Wait a moment for any postback or JavaScript to complete
                    await Task.Delay(1000);

                    // Now click the Generate Report button
                    var generateResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                        (function() {
                            try {
                                var generateButton = document.getElementById('ViewReportImageButton');
                                if (!generateButton) {
                                    return JSON.stringify({success: false, error: 'Generate Report button not found'});
                                }

                                // Check if button is enabled
                                var isDisabled = generateButton.disabled || 
                                               generateButton.style.display === 'none' ||
                                               window.getComputedStyle(generateButton).display === 'none';

                                if (isDisabled) {
                                    return JSON.stringify({success: false, error: 'Generate Report button is disabled or hidden'});
                                }

                                // Click the generate button
                                generateButton.click();

                                return JSON.stringify({
                                    success: true,
                                    message: 'Generate Report button clicked successfully'
                                });
                            } catch (ex) {
                                return JSON.stringify({success: false, error: ex.message});
                            }
                        })()
                    ");

                    var generateResponse = JsonConvert.DeserializeObject<dynamic>(
                        UnwrapScriptJson(generateResult));

                    if ((bool)generateResponse.success)
                    {
                        LogMessage("✓ Generate Report button clicked successfully");
                        UpdateStatus("Report generation started. Waiting for results...", 85);

                        // Wait for the report to generate and load
                        await Task.Delay(5000);

                        // Verify the report has loaded
                        await VerifyReportGenerated();
                    }
                    else
                    {
                        LogMessage($"ERROR clicking Generate Report button: {generateResponse.error}");
                        UpdateStatus("Failed to click Generate Report button", 0);
                    }
                }
                else
                {
                    LogMessage($"ERROR setting day dropdown: {dropdownResponse.error}");
                    UpdateStatus("Failed to set dropdown selection", 0);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Report generation failed - {ex.Message}");
                UpdateStatus("Report generation error", 0);
            }
            finally
            {
                btnGenerateReport.IsEnabled = true;
            }
        }

        private async Task VerifyReportGenerated()
        {
            try
            {
                LogMessage("🔍 Verifying report has been generated...");

                var reportVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        // Look for the ReportViewer1 control and report data
                        var reportViewer = document.getElementById('ReportViewer1');
                        var reportData = null;
                        var hasReportData = false;
                        
                        if (reportViewer) {
                            // Check for typical report elements
                            var reportTables = reportViewer.querySelectorAll('table');
                            var reportCells = reportViewer.querySelectorAll('td, th');
                            hasReportData = reportTables.length > 0 || reportCells.length > 0;
                        }

                        // Also check for error messages
                        var errorElements = document.querySelectorAll('.error, .message, [id*=error], [id*=Error]');
                        var errorMessages = [];
                        for (var i = 0; i < errorElements.length; i++) {
                            var text = errorElements[i].textContent || errorElements[i].innerText;
                            if (text && text.trim() !== '') {
                                errorMessages.push(text.trim());
                            }
                        }

                        return JSON.stringify({
                            hasReportViewer: reportViewer !== null,
                            hasReportData: hasReportData,
                            tableCount: reportViewer ? reportViewer.querySelectorAll('table').length : 0,
                            cellCount: reportViewer ? reportViewer.querySelectorAll('td, th').length : 0,
                            errorMessages: errorMessages,
                            pageTitle: document.title,
                            url: window.location.href
                        });
                    })()
                ");

                var reportInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(reportVerificationResult));

                LogMessage($"Report verification - Has ReportViewer: {reportInfo.hasReportViewer}");
                LogMessage($"  Has Report Data: {reportInfo.hasReportData}");
                LogMessage($"  Table Count: {reportInfo.tableCount}");
                LogMessage($"  Cell Count: {reportInfo.cellCount}");

                if (reportInfo.errorMessages != null && ((Newtonsoft.Json.Linq.JArray)reportInfo.errorMessages).Count > 0)
                {
                    LogMessage("⚠ Error messages found:");
                    foreach (var error in (Newtonsoft.Json.Linq.JArray)reportInfo.errorMessages)
                    {
                        LogMessage($"  - {error}");
                    }
                }

                if ((bool)reportInfo.hasReportViewer && (bool)reportInfo.hasReportData)
                {
                    LogMessage("✅ SUCCESS: Report generated and data available!");
                    UpdateStatus("Report generated successfully. Ready for data extraction.", 100);
                    
                    // Enable data extraction functionality here
                    btnExtractData.IsEnabled = true;
                }
                else if ((bool)reportInfo.hasReportViewer)
                {
                    LogMessage("⚠ Report viewer found but no data visible yet. May still be loading...");
                    UpdateStatus("Report viewer loaded, checking for data...", 90);
                }
                else
                {
                    LogMessage("❌ No report viewer found. Report may have failed to generate.");
                    UpdateStatus("Report generation may have failed", 50);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR verifying report generation: {ex.Message}");
            }
        }

        private async void BtnExtractData_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("🔍 Starting data extraction from report...");
                UpdateStatus("Extracting data from report...", 0);

                // Enhanced JavaScript to extract table data from ReportViewer1
                var extractDataScript = @"
                    (function() {
                        try {
                            var result = {
                                success: false,
                                entries: [],
                                error: null,
                                debug: {
                                    reportViewerFound: false,
                                    tablesFound: 0,
                                    candidateTables: [],
                                    selectedTable: null,
                                    tableRows: 0,
                                    processedRows: 0
                                }
                            };

                            // Find the ReportViewer1 div
                            var reportViewer = document.querySelector('#ReportViewer1');
                            if (!reportViewer) {
                                result.error = 'ReportViewer1 not found';
                                return JSON.stringify(result);
                            }
                            result.debug.reportViewerFound = true;

                            // Look for all tables within the report viewer
                            var tables = reportViewer.querySelectorAll('table');
                            result.debug.tablesFound = tables.length;
                            console.log('Found ' + tables.length + ' tables in report viewer');

                            // Analyze each table to find the data table
                            var bestTable = null;
                            var maxDataRows = 0;
                            
                            for (var i = 0; i < tables.length; i++) {
                                var table = tables[i];
                                var rows = table.querySelectorAll('tr');
                                var tableInfo = {
                                    index: i,
                                    rows: rows.length,
                                    cells: 0,
                                    hasNumericData: false,
                                    hasEmployeeData: false
                                };
                                
                                if (rows.length > 0) {
                                    var firstRow = rows[0];
                                    var cells = firstRow.querySelectorAll('td, th');
                                    tableInfo.cells = cells.length;
                                    
                                    // Check for potential employee data patterns
                                    for (var j = 0; j < Math.min(3, rows.length); j++) {
                                        var rowCells = rows[j].querySelectorAll('td, th');
                                        for (var k = 0; k < rowCells.length; k++) {
                                            var text = rowCells[k].textContent.trim();
                                            // Look for employee ID patterns (numbers) or time patterns
                                            if (/^\d{3,6}$/.test(text)) {
                                                tableInfo.hasEmployeeData = true;
                                            }
                                            if (/\d{1,2}:\d{2}/.test(text)) {
                                                tableInfo.hasNumericData = true;
                                            }
                                        }
                                    }
                                }
                                
                                result.debug.candidateTables.push(tableInfo);
                                
                                // Select best table based on criteria
                                if (rows.length > 1 && tableInfo.cells >= 3 && 
                                    (tableInfo.hasEmployeeData || tableInfo.hasNumericData)) {
                                    if (rows.length > maxDataRows) {
                                        maxDataRows = rows.length;
                                        bestTable = table;
                                        result.debug.selectedTable = tableInfo;
                                    }
                                }
                            }

                            if (!bestTable) {
                                result.error = 'No suitable data table found. Tables analyzed: ' + tables.length;
                                return JSON.stringify(result);
                            }

                            var rows = bestTable.querySelectorAll('tr');
                            result.debug.tableRows = rows.length;
                            console.log('Selected table with ' + rows.length + ' rows');

                            // Extract data from each row
                            for (var i = 0; i < rows.length; i++) {
                                var row = rows[i];
                                var cells = row.querySelectorAll('td, th');
                                
                                if (cells.length >= 2) {
                                    var rowData = [];
                                    for (var j = 0; j < cells.length; j++) {
                                        rowData.push(cells[j].textContent.trim());
                                    }
                                    
                                    // Check if this looks like a data row (not header)
                                    var firstCell = rowData[0];
                                    var isDataRow = /^\d/.test(firstCell) && firstCell.length >= 3;
                                    
                                    if (isDataRow) {
                                        var entry = {
                                            employeeId: rowData[0] || '',
                                            date: rowData[1] || '',
                                            machineName: rowData[2] || '',
                                            direction: rowData[3] || '',
                                            time: rowData[4] || '',
                                            rawData: rowData // Include all cell data for debugging
                                        };
                                        
                                        result.entries.push(entry);
                                        result.debug.processedRows++;
                                    }
                                }
                            }

                            result.success = true;
                            console.log('Extracted ' + result.entries.length + ' entries from ' + result.debug.processedRows + ' data rows');
                            return JSON.stringify(result);
                        }
                        catch (ex) {
                            return JSON.stringify({
                                success: false,
                                entries: [],
                                error: ex.message,
                                debug: result.debug
                            });
                        }
                    })()
                ";

                var extractResult = await webView.CoreWebView2.ExecuteScriptAsync(extractDataScript);
                var extractInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(extractResult));

                // Log debug information
                if (extractInfo.debug != null)
                {
                    LogMessage($"🔍 Debug Info: ReportViewer Found: {extractInfo.debug.reportViewerFound}");
                    LogMessage($"  Tables Found: {extractInfo.debug.tablesFound}");
                    LogMessage($"  Selected Table Rows: {extractInfo.debug.tableRows}");
                    LogMessage($"  Processed Rows: {extractInfo.debug.processedRows}");
                    
                    if (extractInfo.debug.candidateTables != null)
                    {
                        LogMessage($"  Table Analysis:");
                        foreach (var tableInfo in (Newtonsoft.Json.Linq.JArray)extractInfo.debug.candidateTables)
                        {
                            LogMessage($"    Table {tableInfo["index"]}: {tableInfo["rows"]} rows, {tableInfo["cells"]} cells, HasEmployeeData: {tableInfo["hasEmployeeData"]}, HasNumericData: {tableInfo["hasNumericData"]}");
                        }
                    }
                }

                if ((bool)extractInfo.success)
                {
                    var entries = new List<SwipeLogEntry>();
                    
                    foreach (var entry in (Newtonsoft.Json.Linq.JArray)extractInfo.entries)
                    {
                        entries.Add(new SwipeLogEntry
                        {
                            EmployeeId = entry["employeeId"]?.ToString() ?? "",
                            Date = entry["date"]?.ToString() ?? "",
                            MachineName = entry["machineName"]?.ToString() ?? "",
                            Direction = entry["direction"]?.ToString() ?? "",
                            Time = entry["time"]?.ToString() ?? "",
                            Activity = "" // Will be set later by ComparisonWindow labeling
                        });
                    }

                    LogMessage($"✅ Successfully extracted {entries.Count} swipe log entries");
                    
                    // Show sample of raw data for debugging
                    if (entries.Count > 0 && extractInfo.entries.Count > 0)
                    {
                        var firstEntry = ((Newtonsoft.Json.Linq.JArray)extractInfo.entries)[0];
                        if (firstEntry["rawData"] != null)
                        {
                            var rawData = string.Join(" | ", ((Newtonsoft.Json.Linq.JArray)firstEntry["rawData"]).Select(x => x.ToString()));
                            LogMessage($"  Sample Raw Data: {rawData}");
                        }
                    }
                    
                    UpdateStatus($"Data extraction completed - {entries.Count} records", 100);

                    // Open data extraction window with results
                    var dataWindow = new DataExtractionWindow();
                    var reportType = cmbReportType.SelectedItem?.ToString() ?? "Unknown";
                    dataWindow.LoadSwipeLogData(entries, reportType);
                    dataWindow.Show();
                }
                else
                {
                    var errorMsg = extractInfo.error?.ToString() ?? "Unknown error";
                    LogMessage($"❌ Data extraction failed: {errorMsg}");
                    UpdateStatus("Data extraction failed", 0);
                    MessageBox.Show($"Failed to extract data: {errorMsg}", "Extraction Error", 
                                    MessageBoxButton.OK, MessageBoxImage.Error);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR during data extraction: {ex.Message}");
                UpdateStatus("Data extraction error", 0);
                MessageBox.Show($"Error extracting data: {ex.Message}", "Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnReset_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                LogMessage("🔄 Resetting navigation to main page...");
                UpdateStatus("Resetting to main page...", 0);

                // Navigate back to the main report page
                webView.CoreWebView2.Navigate(_config.MisUrl);
                
                // Reset all button states
                btnExpandTree.IsEnabled = false;
                btnClickSwipeLog.IsEnabled = false;
                btnGenerateReport.IsEnabled = false;
                btnExtractData.IsEnabled = false;

                LogMessage("✅ Reset completed - ready for new automation cycle");
                UpdateStatus("Reset completed - page reloaded", 100);
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR during reset: {ex.Message}");
                UpdateStatus("Reset failed", 0);
                MessageBox.Show($"Error during reset: {ex.Message}", "Reset Error", 
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }

        private async void BtnFullAutomation_Click(object sender, RoutedEventArgs e)
        {
            await StartFullAutomation();
        }

        public async Task StartFullAutomation()
        {
            if (!TryBeginAutomation("Full automation")) return;
            try
            {
                _todayData = new List<SwipeLogEntry>();
                _yesterdayData = new List<SwipeLogEntry>();

                LogMessage("🚀 Starting FULL AUTOMATION workflow (Today + Yesterday)...");
                UpdateStatus("Full automation started...", 0);

                LogMessage("📅 PHASE 1: Extracting Today's Data");
                _todayData = await RunDailyReportAsync(isToday: true);

                LogMessage("📅 PHASE 2: Extracting Yesterday's Data");
                _yesterdayData = await RunDailyReportAsync(isToday: false);

                LogMessage("📊 PHASE 3: Displaying Results");
                ShowComparisonResults();

                LogMessage($"✅ FULL AUTOMATION COMPLETED - Today: {_todayData.Count} swipes, Yesterday: {_yesterdayData.Count} swipes");
                UpdateStatus($"Full automation completed - Today: {_todayData.Count}, Yesterday: {_yesterdayData.Count} swipes", 100);
            }
            catch (Exception ex)
            {
                LogMessage($"❌ FULL AUTOMATION FAILED: {ex.Message}");
                UpdateStatus($"Full automation failed: {ex.Message}", 0);
                MessageBox.Show($"Full automation failed: {ex.Message}", "Automation Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                EndAutomation();
            }
        }

        #region Busy guard

        private bool _isAutomationBusy = false;

        /// <summary>
        /// Only one automation may drive the WebView at a time. Returns false (and tells the user)
        /// when another run is in progress.
        /// </summary>
        private bool TryBeginAutomation(string what)
        {
            if (_isAutomationBusy)
            {
                LogMessage($"❌ {what} not started: another automation is already running.");
                UpdateStatus("Another automation is already running - please wait for it to finish.", (int)progressBar.Value);
                return false;
            }
            _isAutomationBusy = true;
            _isFullAutomationRunning = true;
            SetButtonsEnabled(false);
            btnStartFullAutomation.IsEnabled = false;
            if (this.FindName("btnFullReport") is Button fr) fr.IsEnabled = false;
            return true;
        }

        private void EndAutomation()
        {
            _isAutomationBusy = false;
            _isFullAutomationRunning = false;
            SetButtonsEnabled(true);
            btnStartFullAutomation.IsEnabled = true;
            if (this.FindName("btnFullReport") is Button fr) fr.IsEnabled = _isWebViewInitialized;
        }

        #endregion

        #region Reliable page waiting

        /// <summary>
        /// Stamps the current document (and any report tables) with a token so we can tell
        /// when a navigation or postback has actually replaced it.
        /// </summary>
        private async Task<string> MarkCurrentPageAsync()
        {
            var token = Guid.NewGuid().ToString("N");
            await webView.CoreWebView2.ExecuteScriptAsync(
                "window.__cmisPage = '" + token + "';" +
                "(function(){var rv=document.getElementById('ReportViewer1');" +
                "if(rv){var t=rv.querySelectorAll('table');for(var i=0;i<t.length;i++){t[i].setAttribute('data-cmis-stale','" + token + "');}}})();" +
                "true;");
            return token;
        }

        /// <summary>
        /// Polls until the page is no longer the one marked with <paramref name="oldToken"/>
        /// (new document, or the report area was re-rendered by an async postback), the document
        /// has finished loading, and <paramref name="jsCondition"/> is true. Throws on timeout.
        /// </summary>
        private async Task WaitForFreshPageAsync(string oldToken, string jsCondition, string what, int timeoutMs = 60000)
        {
            var script =
                "(function(){try{" +
                "if(document.readyState!=='complete')return false;" +
                "if(window.__cmisPage==='" + oldToken + "'){" +
                "var rv=document.getElementById('ReportViewer1');" +
                "if(!rv||!rv.querySelector('table:not([data-cmis-stale])'))return false;}" +
                "return !!(" + jsCondition + ");" +
                "}catch(e){return false;}})()";

            var sw = System.Diagnostics.Stopwatch.StartNew();
            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                try
                {
                    if (await webView.CoreWebView2.ExecuteScriptAsync(script) == "true")
                    {
                        LogMessage($"✓ {what} ready after {sw.ElapsedMilliseconds / 1000.0:F1}s");
                        return;
                    }
                }
                catch (Exception ex)
                {
                    // Script execution can fail transiently while a navigation is in flight.
                    LogMessage($"  (waiting for {what}: {ex.Message})");
                }
                await Task.Delay(400);
            }

            throw new TimeoutException($"Timed out after {timeoutMs / 1000}s waiting for {what}. Current page: {webView.CoreWebView2.Source}");
        }

        private const string JsMisHomeReady = "document.getElementById('TempleteTreeView')!==null";
        private const string JsSwipeReportFormReady =
            "document.getElementById('DayDropDownList8665')!==null && document.getElementById('ViewReportImageButton')!==null && " +
            "(document.getElementById('EmployeeIDDropDownList7322')!==null || document.querySelector('select[id^=\"EmployeeIDDropDownList\"]')!==null)";
        private const string JsAttendanceFormReady =
            "document.querySelector('select[name*=\"Employee\"], select[id*=\"Employee\"], select[id*=\"employee\"], select[name*=\"EMP\"]')!==null";
        private const string JsReportViewerPresent = "document.getElementById('ReportViewer1')!==null";

        private async Task NavigateToMisHomeAsync()
        {
            var token = await MarkCurrentPageAsync();
            LogMessage($"Navigating to MIS home: {_config.MisUrl}");
            webView.CoreWebView2.Navigate(_config.MisUrl);
            await WaitForFreshPageAsync(token, JsMisHomeReady, "MIS home page (report tree)");
        }

        #endregion

        #region Daily swipe-log report (Today / Yesterday)

        /// <summary>
        /// Runs the "Today's and Yesterday's Swipe Log" report for one day and returns its rows.
        /// Throws on any failure; an empty list means the report rendered and had no swipes.
        /// </summary>
        private async Task<List<SwipeLogEntry>> RunDailyReportAsync(bool isToday)
        {
            var label = isToday ? "Today" : "Yesterday";
            var expectedDate = isToday ? DateTime.Today : DateTime.Today.AddDays(-1);
            var employeeId = RequireEmployeeId();

            try
            {
                UpdateStatus($"{label}: opening MIS...", 10);
                await NavigateToMisHomeAsync();

                UpdateStatus($"{label}: opening swipe log report...", 30);
                await OpenSwipeLogReportAsync();

                UpdateStatus($"{label}: generating report...", 50);
                cmbReportType.SelectedIndex = isToday ? 0 : 1; // keep the UI in sync
                await GenerateSwipeLogReportAsync(isToday ? "1" : "0", label);

                UpdateStatus($"{label}: reading swipe rows...", 80);
                var entries = await ExtractSwipeLogAsync(expectedDate, employeeId, label);
                LogMessage($"✅ {label}: {entries.Count} swipe rows for {expectedDate:dd-MMM-yyyy}");
                return entries;
            }
            catch (Exception ex)
            {
                throw new Exception($"{label} swipe log failed: {ex.Message}", ex);
            }
        }

        private string RequireEmployeeId()
        {
            var id = (txtEmployeeId.Text ?? string.Empty).Trim();
            if (string.IsNullOrEmpty(id))
                throw new InvalidOperationException("Employee ID is empty. Enter your Employee ID first.");
            return id;
        }

        private async Task OpenSwipeLogReportAsync()
        {
            var token = await MarkCurrentPageAsync();
            var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                (function() {
                    try {
                        var anchors = document.querySelectorAll('a');
                        for (var i = 0; i < anchors.length; i++) {
                            var text = (anchors[i].textContent || anchors[i].innerText || '').trim().toLowerCase();
                            var id = anchors[i].id || '';
                            if (text === ""today's and yesterday's swipe log"" && id.indexOf('TempleteTreeView') === 0) {
                                anchors[i].click();
                                return JSON.stringify({success: true});
                            }
                        }
                        return JSON.stringify({success: false, error: ""'Today's and Yesterday's Swipe Log' link not found in the report tree""});
                    } catch (ex) {
                        return JSON.stringify({success: false, error: ex.message});
                    }
                })()
            ");
            var click = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(clickResult));
            if (!(bool)click.success)
                throw new Exception((string)click.error);

            LogMessage("✓ Swipe log link clicked");
            await WaitForFreshPageAsync(token, JsSwipeReportFormReady, "swipe log report form");
            btnGenerateReport.IsEnabled = true;
        }

        private async Task GenerateSwipeLogReportAsync(string dayValue, string label)
        {
            await SelectEmployeeInSwipeLogReport();

            var dayResult = await webView.CoreWebView2.ExecuteScriptAsync($@"
                (function() {{
                    try {{
                        var dd = document.getElementById('DayDropDownList8665');
                        if (!dd) return JSON.stringify({{success: false, error: 'Day dropdown not found'}});
                        var target = {JsonConvert.SerializeObject(dayValue)};
                        var exists = false;
                        for (var i = 0; i < dd.options.length; i++) if (dd.options[i].value === target) exists = true;
                        if (!exists) return JSON.stringify({{success: false, error: 'Day option ' + target + ' not in dropdown'}});
                        dd.value = target;
                        dd.dispatchEvent(new Event('change', {{ bubbles: true }}));
                        return JSON.stringify({{success: dd.value === target, finalValue: dd.value,
                            selectedText: dd.options[dd.selectedIndex].text, error: 'Day dropdown did not accept value ' + target}});
                    }} catch (ex) {{
                        return JSON.stringify({{success: false, error: ex.message}});
                    }}
                }})()
            ");
            var day = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(dayResult));
            if (!(bool)day.success)
                throw new Exception($"Could not select '{label}' in Day dropdown: {day.error}");
            LogMessage($"✓ Day set to '{day.selectedText}' (value {day.finalValue})");

            var token = await MarkCurrentPageAsync();
            var genResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                (function() {
                    var b = document.getElementById('ViewReportImageButton');
                    if (!b) return JSON.stringify({success: false, error: 'Generate Report button (ViewReportImageButton) not found'});
                    if (b.disabled || window.getComputedStyle(b).display === 'none')
                        return JSON.stringify({success: false, error: 'Generate Report button is disabled or hidden'});
                    b.click();
                    return JSON.stringify({success: true});
                })()
            ");
            var gen = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(genResult));
            if (!(bool)gen.success)
                throw new Exception((string)gen.error);

            LogMessage("✓ Generate Report clicked - waiting for the report...");
            await WaitForFreshPageAsync(token, JsReportViewerPresent, $"{label} report");
        }

        /// <summary>
        /// Reads swipe rows from the freshly generated report, waiting for asynchronous rendering.
        /// Validates every row belongs to the requested employee and date so a stale or wrong
        /// page can never be counted silently.
        /// </summary>
        private async Task<List<SwipeLogEntry>> ExtractSwipeLogAsync(DateTime expectedDate, string employeeId, string label,
            int timeoutMs = 45000, int emptySettleMs = 8000)
        {
            const string jsRendering =
                "(function(){var w=document.querySelector('[id$=\"AsyncWait\"],[id*=\"AsyncWait_Wait\"]');" +
                "if(!w)return false;var s=window.getComputedStyle(w);return s.display!=='none'&&s.visibility!=='hidden';})()";

            var sw = System.Diagnostics.Stopwatch.StartNew();
            long settledSince = -1;
            string lastError = "no report table yet";

            while (sw.ElapsedMilliseconds < timeoutMs)
            {
                var raw = await webView.CoreWebView2.ExecuteScriptAsync(GetDataExtractionScript());
                var info = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(raw));

                if ((bool)info.success && ((Newtonsoft.Json.Linq.JArray)info.entries).Count > 0)
                {
                    var entries = new List<SwipeLogEntry>();
                    foreach (var entry in (Newtonsoft.Json.Linq.JArray)info.entries)
                    {
                        entries.Add(new SwipeLogEntry
                        {
                            EmployeeId = entry["employeeId"]?.ToString() ?? "",
                            Date = entry["date"]?.ToString() ?? "",
                            MachineName = entry["machineName"]?.ToString() ?? "",
                            Direction = entry["direction"]?.ToString() ?? "",
                            Time = entry["time"]?.ToString() ?? "",
                        });
                    }
                    ValidateSwipeRows(entries, expectedDate, employeeId, label);
                    return entries;
                }

                lastError = (string?)info.error ?? "report table has no swipe rows";

                // No rows yet. Once the report viewer has stopped rendering and stayed that way
                // for a while, accept that the day genuinely has no swipes.
                bool rendering = await webView.CoreWebView2.ExecuteScriptAsync(jsRendering) == "true";
                bool viewerPresent = await webView.CoreWebView2.ExecuteScriptAsync(JsReportViewerPresent) == "true";
                if (viewerPresent && !rendering)
                {
                    if (settledSince < 0) settledSince = sw.ElapsedMilliseconds;
                    if (sw.ElapsedMilliseconds - settledSince >= emptySettleMs)
                    {
                        LogMessage($"ℹ {label}: report finished rendering with no swipe rows ({lastError}) - treating as no swipes for {expectedDate:dd-MMM-yyyy}");
                        return new List<SwipeLogEntry>();
                    }
                }
                else
                {
                    settledSince = -1;
                }

                await Task.Delay(1000);
            }

            throw new TimeoutException($"Report did not show swipe rows within {timeoutMs / 1000}s ({lastError})");
        }

        private void ValidateSwipeRows(List<SwipeLogEntry> entries, DateTime expectedDate, string employeeId, string label)
        {
            var badDate = entries.FirstOrDefault(e => WorkHoursCalculator.ParseDate(e.Date) == null);
            if (badDate != null)
                throw new Exception($"Unrecognised date '{badDate.Date}' in {label} swipe log (row: {badDate.EmployeeId} | {badDate.Date} | {badDate.MachineName}). The report layout may have changed.");

            var wrongDate = entries.Where(e => WorkHoursCalculator.ParseDate(e.Date) != expectedDate.Date)
                                   .Select(e => e.Date).Distinct().ToList();
            if (wrongDate.Count > 0)
                throw new Exception($"{label} report shows date(s) {string.Join(", ", wrongDate)} but {expectedDate:dd-MMM-yyyy} was requested (stale or wrong report page).");

            var wrongEmployee = entries.Where(e => !string.Equals(e.EmployeeId.Trim(), employeeId, StringComparison.OrdinalIgnoreCase))
                                       .Select(e => e.EmployeeId).Distinct().ToList();
            if (wrongEmployee.Count > 0)
                throw new Exception($"{label} report contains employee(s) {string.Join(", ", wrongEmployee)} but {employeeId} was requested.");
        }

        #endregion

        private string GetDataExtractionScript()
        {
            return @"
                    (function() {
                        try {
                            var result = {
                                success: false,
                                entries: [],
                                error: null,
                                debug: {
                                    reportViewerFound: false,
                                    tablesFound: 0,
                                    candidateTables: [],
                                    selectedTable: null,
                                    tableRows: 0,
                                    processedRows: 0
                                }
                            };

                            // Find the ReportViewer1 div
                            var reportViewer = document.querySelector('#ReportViewer1');
                            if (!reportViewer) {
                                result.error = 'ReportViewer1 not found';
                                return JSON.stringify(result);
                            }
                            result.debug.reportViewerFound = true;

                            // Look for all tables within the report viewer
                            var tables = reportViewer.querySelectorAll('table');
                            result.debug.tablesFound = tables.length;

                            // Analyze each table to find the data table
                            var bestTable = null;
                            var maxDataRows = 0;
                            
                            for (var i = 0; i < tables.length; i++) {
                                var table = tables[i];
                                var rows = table.querySelectorAll('tr');
                                var tableInfo = {
                                    index: i,
                                    rows: rows.length,
                                    cells: 0,
                                    hasNumericData: false,
                                    hasEmployeeData: false
                                };
                                
                                if (rows.length > 0) {
                                    var firstRow = rows[0];
                                    var cells = firstRow.querySelectorAll('td, th');
                                    tableInfo.cells = cells.length;
                                    
                                    // Check for potential employee data patterns
                                    for (var j = 0; j < Math.min(3, rows.length); j++) {
                                        var rowCells = rows[j].querySelectorAll('td, th');
                                        for (var k = 0; k < rowCells.length; k++) {
                                            var text = rowCells[k].textContent.trim();
                                            // Look for employee ID patterns (numbers) or time patterns
                                            if (/^\d{3,6}$/.test(text)) {
                                                tableInfo.hasEmployeeData = true;
                                            }
                                            if (/\d{1,2}:\d{2}/.test(text)) {
                                                tableInfo.hasNumericData = true;
                                            }
                                        }
                                    }
                                }
                                
                                result.debug.candidateTables.push(tableInfo);
                                
                                // Select best table based on criteria
                                if (rows.length > 1 && tableInfo.cells >= 3 && 
                                    (tableInfo.hasEmployeeData || tableInfo.hasNumericData)) {
                                    if (rows.length > maxDataRows) {
                                        maxDataRows = rows.length;
                                        bestTable = table;
                                        result.debug.selectedTable = tableInfo;
                                    }
                                }
                            }

                            if (!bestTable) {
                                result.error = 'No suitable data table found. Tables analyzed: ' + tables.length;
                                return JSON.stringify(result);
                            }

                            var rows = bestTable.querySelectorAll('tr');
                            result.debug.tableRows = rows.length;

                            // Extract data from each row
                            for (var i = 0; i < rows.length; i++) {
                                var row = rows[i];
                                var cells = row.querySelectorAll('td, th');
                                
                                if (cells.length >= 2) {
                                    var rowData = [];
                                    for (var j = 0; j < cells.length; j++) {
                                        rowData.push(cells[j].textContent.trim());
                                    }
                                    
                                    // Check if this looks like a data row (not header)
                                    var firstCell = rowData[0];
                                    var isDataRow = /^\d/.test(firstCell) && firstCell.length >= 3;
                                    
                                    if (isDataRow) {
                                        var entry = {
                                            employeeId: rowData[0] || '',
                                            date: rowData[1] || '',
                                            machineName: rowData[2] || '',
                                            direction: rowData[3] || '',
                                            time: rowData[4] || '',
                                            rawData: rowData
                                        };
                                        
                                        result.entries.push(entry);
                                        result.debug.processedRows++;
                                    }
                                }
                            }

                            result.success = true;
                            return JSON.stringify(result);
                        }
                        catch (ex) {
                            return JSON.stringify({
                                success: false,
                                entries: [],
                                error: ex.message,
                                debug: result.debug
                            });
                        }
                    })()
                ";
        }

        private void ShowComparisonResults()
        {
            var comparisonWindow = new ComparisonWindow(this);
            comparisonWindow.LoadComparisonData(_yesterdayData, _todayData);
            comparisonWindow.Show();
        }

        private void SetButtonsEnabled(bool enabled)
        {
            btnStartAutomation.IsEnabled = enabled;
            btnTestPage.IsEnabled = enabled;
            btnExpandTree.IsEnabled = enabled;
            btnClickSwipeLog.IsEnabled = enabled;
            btnGenerateReport.IsEnabled = enabled;
            btnExtractData.IsEnabled = enabled;
            btnReset.IsEnabled = enabled;
            btnMonthlyReport.IsEnabled = enabled;
            if (this.FindName("btnFullReport") is Button frBtnToggle)
                frBtnToggle.IsEnabled = enabled && _monthlyReportCache.Count > 0;
        }

        // Removed old wait methods - now using the same completion detection as manual mode

        private async Task ExecuteExpandTree()
        {
            // Call the SAME function that works in manual mode
            LogMessage("🔄 Calling manual tree expansion function...");
            
            // Simulate the manual button click by calling the same method
            await CallManualExpandTreeFunction();
        }

        private async Task CallManualExpandTreeFunction()
        {
            return;
            // This calls the exact same logic as the manual BtnExpandTree_Click
            /*try
            {
                UpdateStatus("Expanding Leave Management System tree node...", 20);
                LogMessage("Starting tree expansion process...");

                // First, check if the node is already expanded (same as manual)
                var isExpandedResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var node = document.getElementById('TempleteTreeViewn21');
                        if (!node) return JSON.stringify({success: false, error: 'Node not found'});
                        
                        var img = node.querySelector('img');
                        var isExpanded = img && img.src.includes('minus');
                        
                        return JSON.stringify({
                            success: true, 
                            isExpanded: isExpanded,
                            imgSrc: img ? img.src : 'no image',
                            nodeText: node.textContent || 'no text'
                        });
                    })()
                ");

                var expandCheck = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(isExpandedResult));
                LogMessage($"Tree node check: Success={expandCheck.success}, IsExpanded={expandCheck.isExpanded}");

                if (!(bool)expandCheck.success)
                {
                    LogMessage($"ERROR: {expandCheck.error}");
                    throw new Exception($"Tree expansion failed: {expandCheck.error}");
                }

                if ((bool)expandCheck.isExpanded)
                {
                    LogMessage("✓ Leave Management System tree is already expanded");
                    UpdateStatus("Tree already expanded. Looking for swipe log option...", 50);
                    await CheckForSwipeLogOption();
                }
                else
                {
                    LogMessage("🔄 Expanding Leave Management System tree node...");

                    // Click the expand icon to expand the tree (same as manual)
                    var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                        (function() {
                            try {
                                var node = document.getElementById('TempleteTreeViewn21');
                                if (!node) return JSON.stringify({success: false, error: 'Node not found'});
                                
                                // Simulate the click event
                                node.click();
                                
                                return JSON.stringify({success: true, message: 'Click executed'});
                            } catch (ex) {
                                return JSON.stringify({success: false, error: ex.message});
                            }
                        })()
                    ");

                    var clickResponse = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(clickResult));
                    LogMessage($"Click result: Success={clickResponse.success}");

                    if ((bool)clickResponse.success)
                    {
                        LogMessage("✓ Tree node clicked successfully");
                        UpdateStatus("Waiting for tree expansion to complete...", 60);

                        // Wait a moment for the tree to expand (same as manual)
                        await Task.Delay(2000);

                        // Verify expansion occurred and look for swipe log option (same as manual)
                        await CheckForSwipeLogOption();
                    }
                    else
                    {
                        LogMessage($"ERROR clicking tree node: {clickResponse.error}");
                        throw new Exception($"Tree expansion failed: {clickResponse.error}");
                    }
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Tree expansion failed - {ex.Message}");
                throw;
            }*/
        }


        /// <summary>
        /// Selects the employee from txtEmployeeId in the swipe log report's employee dropdown
        /// (EmployeeIDDropDownList*). Options look like "Name - 1476" with value "1476".
        /// </summary>
        private async Task SelectEmployeeInSwipeLogReport()
        {
            var targetEmployeeId = RequireEmployeeId();

            LogMessage($"Selecting employee '{targetEmployeeId}' in swipe log report...");
            var jsTargetId = JsonConvert.SerializeObject(targetEmployeeId);

            var result = await webView.CoreWebView2.ExecuteScriptAsync($@"
                (function() {{
                    try {{
                        var dd = document.getElementById('EmployeeIDDropDownList7322') ||
                                 document.querySelector('select[id^=""EmployeeIDDropDownList""]') ||
                                 document.querySelector('select[title=""EmployeeID""]');
                        if (!dd) {{
                            return JSON.stringify({{success: false, error: 'Employee dropdown not found'}});
                        }}

                        var targetId = {jsTargetId};
                        var previousValue = dd.value;
                        var matchIndex = -1;

                        // 1) Exact value match
                        for (var i = 0; i < dd.options.length; i++) {{
                            if ((dd.options[i].value || '').trim() === targetId) {{ matchIndex = i; break; }}
                        }}
                        // 2) Text ends with '- <id>' (e.g. 'Ramkrishna Pawar - 1476')
                        if (matchIndex < 0) {{
                            for (var j = 0; j < dd.options.length; j++) {{
                                var parts = (dd.options[j].text || '').split('-');
                                if (parts[parts.length - 1].trim() === targetId) {{ matchIndex = j; break; }}
                            }}
                        }}

                        if (matchIndex < 0) {{
                            var available = [];
                            for (var k = 0; k < Math.min(dd.options.length, 15); k++) {{
                                available.push(dd.options[k].value + ' | ' + dd.options[k].text);
                            }}
                            return JSON.stringify({{success: false, error: 'Employee ' + targetId + ' not found in dropdown', available: available}});
                        }}

                        dd.selectedIndex = matchIndex;
                        dd.dispatchEvent(new Event('change', {{ bubbles: true }}));

                        return JSON.stringify({{
                            success: true,
                            previousValue: previousValue,
                            finalValue: dd.value,
                            selectedText: dd.options[dd.selectedIndex].text,
                            selectionChanged: previousValue !== dd.value
                        }});
                    }} catch (ex) {{
                        return JSON.stringify({{success: false, error: ex.message}});
                    }}
                }})()
            ");

            var info = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(result));
            if (!(bool)info.success)
            {
                LogMessage($"ERROR selecting employee: {info.error}");
                if (info.available != null)
                {
                    LogMessage("Available employees:");
                    foreach (var opt in info.available)
                        LogMessage($"  - {opt}");
                }
                throw new Exception($"Failed to select employee: {info.error}");
            }

            LogMessage($"✓ Employee selected: '{info.selectedText}' (Value: {info.finalValue})");
            if ((bool)info.selectionChanged)
            {
                LogMessage($"  Changed from '{info.previousValue}' to '{info.finalValue}'");
                await Task.Delay(1000); // let any change handlers finish
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                if (_logWindow != null)
                {
                    _logWindow.AllowClose = true;
                    _logWindow.Close();
                }
                webView?.Dispose();
            }
            catch (Exception ex)
            {
                // Log but don't prevent closing
                System.Diagnostics.Debug.WriteLine($"Error disposing WebView: {ex.Message}");
            }
        }

        // Manual mode controls removed - keeping manual controls permanently hidden
        /*
        private void ChkManualMode_Checked(object sender, RoutedEventArgs e)
        {
            pnlManualControls.Visibility = Visibility.Visible;
            LogMessage("Manual Mode enabled - showing manual controls");
        }

        private void ChkManualMode_Unchecked(object sender, RoutedEventArgs e)
        {
            pnlManualControls.Visibility = Visibility.Collapsed;
            LogMessage("Manual Mode disabled - hiding manual controls");
        }
        */
        
        private async void ChkShowLogs_Checked(object sender, RoutedEventArgs e)
        {
            if (_logWindow != null)
            {
                _logWindow.Show();
                _config.ShowLogWindow = true;
                await SaveConfigSafelyAsync();
                LogMessage("Log window shown");
            }
        }

        private async void ChkShowLogs_Unchecked(object sender, RoutedEventArgs e)
        {
            if (_logWindow != null)
            {
                _logWindow.Hide();
                _config.ShowLogWindow = false;
                await SaveConfigSafelyAsync();
                LogMessage("Log window hidden");
            }
        }
        
        private async void ChkAutoStartFullReport_Checked(object sender, RoutedEventArgs e)
        {
            _config.AutoStartFullReport = true;
            await SaveConfigSafelyAsync();
            LogMessage("Auto start full report enabled");
        }

        private async void ChkAutoStartFullReport_Unchecked(object sender, RoutedEventArgs e)
        {
            _config.AutoStartFullReport = false;
            await SaveConfigSafelyAsync();
            LogMessage("Auto start full report disabled");
        }
        
        private System.Windows.Threading.DispatcherTimer? _employeeIdSaveTimer;
        private bool _configSaveErrorShown = false;

        private void TxtEmployeeId_TextChanged(object sender, TextChangedEventArgs e)
        {
            // Debounce: save once typing pauses instead of writing the file on every keystroke.
            if (_employeeIdSaveTimer == null)
            {
                _employeeIdSaveTimer = new System.Windows.Threading.DispatcherTimer { Interval = TimeSpan.FromMilliseconds(800) };
                _employeeIdSaveTimer.Tick += async (_, _) =>
                {
                    _employeeIdSaveTimer!.Stop();
                    var id = (txtEmployeeId.Text ?? string.Empty).Trim();
                    if (_config != null && id != _config.EmployeeId)
                    {
                        _config.EmployeeId = id;
                        await SaveConfigSafelyAsync();
                        LogMessage($"Employee ID updated to: {id}");
                    }
                };
            }
            _employeeIdSaveTimer.Stop();
            _employeeIdSaveTimer.Start();
        }

        /// <summary>
        /// Saves settings without ever throwing into a UI event handler. A failure is logged,
        /// shown in the status bar, and reported once per session in a dialog.
        /// </summary>
        private async Task<bool> SaveConfigSafelyAsync()
        {
            try
            {
                await ConfigurationService.SaveConfigurationAsync(_config);
                return true;
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Could not save settings to {ConfigurationService.ConfigFilePath}: {ex.Message}");
                UpdateStatus($"Settings not saved: {ex.Message}", (int)progressBar.Value);
                if (!_configSaveErrorShown)
                {
                    _configSaveErrorShown = true;
                    MessageBox.Show($"Your settings could not be saved:\n\n{ex.Message}\n\nFile: {ConfigurationService.ConfigFilePath}",
                        "Settings Not Saved", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                return false;
            }
        }
        
        private async Task HandleAutoStartWithEmployeeIdValidation()
        {
            try
            {
                // Check if employee ID is empty
                if (string.IsNullOrWhiteSpace(_config.EmployeeId))
                {
                    LogMessage("Auto-start requested but Employee ID is empty. Prompting user...");
                    
                    // Show popup on UI thread and get result
                    var dialogResult = await Dispatcher.InvokeAsync(() =>
                    {
                        var dialog = new EmployeeIdInputDialog()
                        {
                            Owner = this
                        };
                        var result = dialog.ShowDialog();
                        return new { Success = result == true, EmployeeId = result == true ? dialog.EmployeeId : null };
                    });
                    
                    if (dialogResult.Success && !string.IsNullOrWhiteSpace(dialogResult.EmployeeId))
                    {
                        // Update config and UI
                        _config.EmployeeId = dialogResult.EmployeeId.Trim();
                        await SaveConfigSafelyAsync(); // a save failure must not block auto-start
                        
                        // Update UI on main thread
                        Dispatcher.Invoke(() => txtEmployeeId.Text = dialogResult.EmployeeId);
                        
                        LogMessage($"Employee ID set to: {dialogResult.EmployeeId}. Proceeding with auto-start...");
                        
                        // Proceed with auto-start
                        await StartAutoReport();
                    }
                    else
                    {
                        LogMessage("Auto-start cancelled: No valid Employee ID provided.");
                    }
                }
                else
                {
                    LogMessage("Auto-starting full calendar report...");
                    await StartAutoReport();
                }
            }
            catch (Exception ex)
            {
                LogMessage($"Error during auto-start: {ex}");
                UpdateStatus($"Auto-start failed: {ex.Message}", 0);
                MessageBox.Show($"Auto-start failed: {ex.Message}", "Auto-start Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
        }
        
        private async Task StartAutoReport()
        {
            await Task.Delay(1000); // Brief delay after initialization
            BtnFullReport_Click(btnFullReport, new RoutedEventArgs()); // guarded by TryBeginAutomation
        }

        #region Monthly Report Functions

        private async void BtnMonthlyReport_Click(object sender, RoutedEventArgs e)
        {
            if (!TryBeginAutomation("Monthly report")) return;
            try
            {
                LogMessage("📅 Starting Monthly Report extraction...");
                UpdateStatus("Extracting monthly data...", 0);

                var monthlyData = await ExtractMonthlyAttendanceData();
                
                if (monthlyData.Count > 0)
                {
                    // Cache data for calendar
                    UpdateMonthlyReportCache(monthlyData);
                    if (this.FindName("btnFullReport") is Button frBtnMonthly)
                        frBtnMonthly.IsEnabled = true;

                    var monthlyWindow = new MonthlyWindow(txtEmployeeId.Text);
                    monthlyWindow.LoadMonthlyData(monthlyData, txtEmployeeId.Text);
                    monthlyWindow.Show();
                    
                    LogMessage($"✅ Monthly report extracted successfully - {monthlyData.Count} records");
                    UpdateStatus($"Monthly report completed - {monthlyData.Count} records", 100);
                }
                else
                {
                    LogMessage("ℹ Monthly report has no rows for this period");
                    UpdateStatus("Monthly report has no rows for this period", 100);
                    MessageBox.Show("The monthly attendance report was generated but has no rows for this period.", "No Data", 
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Monthly report extraction failed: {ex.Message}");
                UpdateStatus($"Monthly report failed: {ex.Message}", 0);
                MessageBox.Show($"Monthly report extraction failed: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                EndAutomation();
            }
        }

        private async Task<List<MonthlyAttendanceEntry>> ExtractMonthlyAttendanceData()
        {
            var monthlyData = new List<MonthlyAttendanceEntry>();

            try
            {
                // Step 1: Navigate to main page
                LogMessage("1️⃣ Navigating to main page...");
                await NavigateToMisHomeAsync();

                // Step 2: Navigate directly to Attendance Log Report (tree expansion not required)
                LogMessage("2️⃣ Navigating to Attendance Log Report...");
                await NavigateToAttendanceLogReport();

                // Step 3: Select employee
                LogMessage("3️⃣ Selecting employee...");
                await SelectEmployeeInAttendanceReport();

                // Step 4: Set date range (1st of current month to today)
                LogMessage("4️⃣ Setting date range for current month...");
                await SetMonthlyDateRange();

                // Step 5: Generate report
                LogMessage("5️⃣ Generating monthly report...");
                await GenerateAttendanceReport();

                // Step 6: Extract data from report
                LogMessage("6️⃣ Extracting data from monthly report...");
                monthlyData = await ParseMonthlyReportData();

                LogMessage($"✅ Successfully extracted {monthlyData.Count} monthly records");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Error in monthly data extraction: {ex.Message}");
                throw;
            }

            return monthlyData;
        }

        private async Task NavigateToAttendanceLogReport()
        {
            try
            {
                UpdateStatus("Clicking 'Attendance Log Report' link...", 30);
                LogMessage("Starting Attendance Log Report click process...");

                // First verify the Attendance Log Report link is available and get its details (located by visible text)
                var linkVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var anchors = document.querySelectorAll('a');
                        var target = null;
                        for (var i = 0; i < anchors.length; i++) {
                            var text = (anchors[i].textContent || anchors[i].innerText || '').trim();
                            var id = anchors[i].id || '';
                            if (text && text.toLowerCase() === 'attendance log report' && id.startsWith('TempleteTreeView')) {
                                target = anchors[i];
                                break;
                            }
                        }

                        if (target) {
                            return JSON.stringify({
                                found: true,
                                id: target.id || 'no-id',
                                text: target.textContent || target.innerText,
                                href: target.getAttribute('href') || '',
                                hasOnClick: !!target.getAttribute('onclick'),
                                visible: window.getComputedStyle(target).display !== 'none',
                                className: target.className || ''
                            });
                        }

                        return JSON.stringify({found: false, error: 'Attendance Log Report link not found'});
                    })()
                ");

                var linkInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(linkVerificationResult));

                if (!(bool)linkInfo.found)
                {
                    LogMessage("ERROR: Attendance Log Report link not found or disappeared");
                    throw new Exception("Attendance Log Report link not found or disappeared");
                }

                LogMessage($"✓ Attendance Log Report link verified: '{linkInfo.text}'");
                LogMessage($"  ID: {linkInfo.id}, Class: {linkInfo.className}, HasOnClick: {linkInfo.hasOnClick}");

                // Click the Attendance Log Report link (resolved by text match)
                var navToken = await MarkCurrentPageAsync();
                var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var anchors = document.querySelectorAll('a');
                            var target = null;
                            for (var i = 0; i < anchors.length; i++) {
                                var text = (anchors[i].textContent || anchors[i].innerText || '').trim();
                                var id = anchors[i].id || '';
                                if (text && text.toLowerCase() === 'attendance log report' && id.startsWith('TempleteTreeView')) {
                                    target = anchors[i];
                                    break;
                                }
                            }

                            if (!target) {
                                return JSON.stringify({success: false, error: 'Link not found'});
                            }

                            // Simulate the click - this should trigger the __doPostBack
                            target.click();
                            
                            return JSON.stringify({
                                success: true, 
                                message: 'Attendance Log Report link clicked successfully'
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: ex.message});
                        }
                    })()
                ");

                var clickResponse = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(clickResult));

                if ((bool)clickResponse.success)
                {
                    LogMessage("✓ Attendance Log Report link clicked successfully");
                    UpdateStatus("Waiting for Attendance Log Report page to load...", 60);

                    // Wait until the postback has replaced the page with the report form
                    await WaitForFreshPageAsync(navToken, JsAttendanceFormReady, "Attendance Log Report form");

                    LogMessage("✓ Navigation to Attendance Log Report completed");
                }
                else
                {
                    LogMessage($"ERROR clicking Attendance Log Report link: {clickResponse.error}");
                    throw new Exception($"Attendance Log Report click failed: {clickResponse.error}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Attendance Log Report click failed - {ex.Message}");
                throw;
            }
        }

        private async Task SelectEmployeeInAttendanceReport()
        {
            try
            {
                UpdateStatus("Selecting employee in attendance report...", 50);
                LogMessage("Starting employee selection process...");

                var targetEmployeeId = RequireEmployeeId();
                LogMessage($"Looking for employee with ID: {targetEmployeeId}");

                // Step 1: Find and verify the employee dropdown exists
                var dropdownResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            // Try multiple common selectors for employee dropdown
                            var employeeDropdown = document.querySelector('select[name*=""Employee""], select[id*=""Employee""], select[id*=""employee""], select[name*=""EMP""]') ||
                                                 document.getElementById('EmployeeDropDownList') ||
                                                 document.querySelector('select:has(option[value*=""EMP""])');
                            
                            if (!employeeDropdown) {
                                // Debug - show available dropdowns
                                var allSelects = document.querySelectorAll('select');
                                var available = [];
                                for (var i = 0; i < Math.min(allSelects.length, 5); i++) {
                                    available.push({
                                        id: allSelects[i].id || 'no-id',
                                        name: allSelects[i].name || 'no-name',
                                        optionCount: allSelects[i].options.length
                                    });
                                }
                                return JSON.stringify({success: false, error: 'Employee dropdown not found', available: available});
                            }

                            // Get dropdown info and sample options
                            var options = [];
                            for (var i = 0; i < Math.min(employeeDropdown.options.length, 10); i++) {
                                options.push({
                                    value: employeeDropdown.options[i].value,
                                    text: employeeDropdown.options[i].text
                                });
                            }

                            return JSON.stringify({
                                success: true,
                                id: employeeDropdown.id || 'no-id',
                                name: employeeDropdown.name || 'no-name', 
                                optionCount: employeeDropdown.options.length,
                                sampleOptions: options,
                                currentValue: employeeDropdown.value,
                                currentText: employeeDropdown.options[employeeDropdown.selectedIndex].text
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: 'Employee dropdown error'});
                        }
                    })()
                ");

                var dropdownInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(dropdownResult));

                if (!(bool)dropdownInfo.success)
                {
                    LogMessage($"ERROR finding employee dropdown: {dropdownInfo.error}");
                    if (dropdownInfo.available != null)
                    {
                        LogMessage("Available dropdowns:");
                        foreach (var dropdown in dropdownInfo.available)
                        {
                            LogMessage($"  - ID: {dropdown.id}, Name: {dropdown.name}, Options: {dropdown.optionCount}");
                        }
                    }
                    throw new Exception($"Employee dropdown not found: {dropdownInfo.error}");
                }

                LogMessage($"✓ Employee dropdown found: ID='{dropdownInfo.id}', Name='{dropdownInfo.name}', Options={dropdownInfo.optionCount}");
                LogMessage($"  Current selection: '{dropdownInfo.currentText}' (Value: {dropdownInfo.currentValue})");
                LogMessage("Sample options:");
                foreach (var option in dropdownInfo.sampleOptions)
                {
                    LogMessage($"  - Value: '{option.value}', Text: '{option.text}'");
                }

                // Step 2: Select the employee by ID match
                var selectionResult = await webView.CoreWebView2.ExecuteScriptAsync($@"
                    (function() {{
                        try {{
                            var employeeDropdown = document.querySelector('select[name*=""Employee""], select[id*=""Employee""], select[id*=""employee""], select[name*=""EMP""]') ||
                                                 document.getElementById('EmployeeDropDownList') ||
                                                 document.querySelector('select:has(option[value*=""EMP""])');
                            
                            if (!employeeDropdown) {{
                                return JSON.stringify({{success: false, error: 'Employee dropdown disappeared'}});
                            }}

                            var targetId = {JsonConvert.SerializeObject(targetEmployeeId)};
                            var found = false;
                            var selectedOption = null;
                            var previousIndex = employeeDropdown.selectedIndex;
                            
                            // Exact ID match: option value first, then the ID after the last '-' in the text
                            // (e.g. 'Ramkrishna Pawar - 1476'), so '1476' never matches '14760'.
                            var matchIndex = -1;
                            for (var i = 0; i < employeeDropdown.options.length; i++) {{
                                if ((employeeDropdown.options[i].value || '').trim() === targetId) {{ matchIndex = i; break; }}
                            }}
                            if (matchIndex < 0) {{
                                for (var j = 0; j < employeeDropdown.options.length; j++) {{
                                    var parts = (employeeDropdown.options[j].text || '').split('-');
                                    if (parts[parts.length - 1].trim() === targetId) {{ matchIndex = j; break; }}
                                }}
                            }}
                            if (matchIndex >= 0) {{
                                var option = employeeDropdown.options[matchIndex];
                                employeeDropdown.selectedIndex = matchIndex;
                                selectedOption = {{
                                    value: option.value,
                                    text: option.text,
                                    index: matchIndex
                                }};
                                found = true;
                            }}
                            
                            if (found) {{
                                // Trigger change event
                                var changeEvent = new Event('change', {{ bubbles: true }});
                                employeeDropdown.dispatchEvent(changeEvent);
                                
                                return JSON.stringify({{
                                    success: true,
                                    selectedOption: selectedOption,
                                    previousIndex: previousIndex,
                                    selectionChanged: previousIndex !== selectedOption.index,
                                    message: 'Employee selected successfully'
                                }});
                            }} else {{
                                return JSON.stringify({{
                                    success: false, 
                                    error: 'Employee not found with ID: ' + targetId,
                                    searchedId: targetId,
                                    totalOptions: employeeDropdown.options.length
                                }});
                            }}
                        }} catch (ex) {{
                            return JSON.stringify({{success: false, error: 'Employee selection error'}});
                        }}
                    }})()
                ");

                var selectionInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(selectionResult));

                if (!(bool)selectionInfo.success)
                {
                    LogMessage($"ERROR selecting employee: {selectionInfo.error}");
                    throw new Exception($"Employee selection failed: {selectionInfo.error}");
                }

                LogMessage($"✓ Employee selected successfully: '{selectionInfo.selectedOption.text}' (Index: {selectionInfo.selectedOption.index})");
                LogMessage($"  Selection changed: {selectionInfo.selectionChanged} (Previous index: {selectionInfo.previousIndex})");

                // Step 3: Verify the selection stuck by reading it back
                await Task.Delay(500); // Brief pause to let change events complete

                var verificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var employeeDropdown = document.querySelector('select[name*=""Employee""], select[id*=""Employee""], select[id*=""employee""], select[name*=""EMP""]') ||
                                                 document.getElementById('EmployeeDropDownList') ||
                                                 document.querySelector('select:has(option[value*=""EMP""])');
                            
                            if (!employeeDropdown) {
                                return JSON.stringify({success: false, error: 'Employee dropdown disappeared during verification'});
                            }

                            return JSON.stringify({
                                success: true,
                                currentIndex: employeeDropdown.selectedIndex,
                                currentValue: employeeDropdown.value,
                                currentText: employeeDropdown.options[employeeDropdown.selectedIndex].text
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: ex.message});
                        }
                    })()
                ");

                var verificationInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(verificationResult));

                if ((bool)verificationInfo.success)
                {
                    LogMessage($"✓ Employee selection verified: '{verificationInfo.currentText}' (Index: {verificationInfo.currentIndex})");
                    
                    // Verify the selection is exactly our target ID (value, or the ID after the last '-')
                    string currentValue = ((string)verificationInfo.currentValue ?? "").Trim();
                    string currentText = (string)verificationInfo.currentText ?? "";
                    string textId = currentText.Split('-').Last().Trim();
                    if (currentValue == targetEmployeeId || textId == targetEmployeeId)
                    {
                        LogMessage($"✓ Verified selected employee is {targetEmployeeId}");
                    }
                    else
                    {
                        throw new Exception($"Employee dropdown shows '{currentText}' (value {currentValue}) instead of {targetEmployeeId} after selection");
                    }
                }
                else
                {
                    LogMessage($"ERROR verifying employee selection: {verificationInfo.error}");
                    throw new Exception($"Employee selection verification failed: {verificationInfo.error}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Employee selection failed - {ex.Message}");
                throw;
            }
        }

        private async Task SetMonthlyDateRange()
        {
            try
            {
                UpdateStatus("Setting date range for current month...", 60);
                LogMessage("Starting date range configuration...");

                // Set date range to current month (1st to last day)
                var firstOfMonth = new DateTime(DateTime.Now.Year, DateTime.Now.Month, 1);
                var lastOfMonth = firstOfMonth.AddMonths(1).AddDays(-1);
                var startDateString = firstOfMonth.ToString("dd-MMM-yyyy");
                var endDateString = lastOfMonth.ToString("dd-MMM-yyyy");
                LogMessage($"Setting date range: {startDateString} to {endDateString}");

                // Step 1: Find and verify both date inputs exist
                var dateInputResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            // Use specific input name for monthly date range
                            var fromDateInput = document.querySelector('input[name=""DMNDateDateRangeControl4392_FromDateCalender_DTB""]') ||
                                              document.querySelector('input[name*=""FromDateCalender""], input[name*=""FromDate""]') ||
                                              document.querySelector('input[type=""text""][name*=""Date""]');
                            
                            // Also find to date input
                            var toDateInput = document.querySelector('input[name*=""ToDateCalender""], input[name*=""ToDate""]') ||
                                            document.querySelectorAll('input[type=""text""][name*=""Date""]')[1];
                            
                            if (!fromDateInput) {
                                // Debug - show available date-related inputs
                                var allInputs = document.querySelectorAll('input[type=""text""], input[type=""date""]');
                                var available = [];
                                for (var i = 0; i < Math.min(allInputs.length, 8); i++) {
                                    available.push({
                                        id: allInputs[i].id || 'no-id',
                                        name: allInputs[i].name || 'no-name',
                                        placeholder: allInputs[i].placeholder || 'no-placeholder',
                                        title: allInputs[i].title || 'no-title',
                                        value: allInputs[i].value || 'no-value'
                                    });
                                }
                                return JSON.stringify({success: false, error: 'From date input not found', available: available});
                            }

                            return JSON.stringify({
                                success: true,
                                fromDate: {
                                    id: fromDateInput.id || 'no-id',
                                    name: fromDateInput.name || 'no-name',
                                    type: fromDateInput.type,
                                    placeholder: fromDateInput.placeholder || 'no-placeholder',
                                    currentValue: fromDateInput.value || 'empty',
                                    readOnly: fromDateInput.readOnly,
                                    disabled: fromDateInput.disabled
                                },
                                toDate: toDateInput ? {
                                    id: toDateInput.id || 'no-id',
                                    name: toDateInput.name || 'no-name',
                                    type: toDateInput.type,
                                    placeholder: toDateInput.placeholder || 'no-placeholder',
                                    currentValue: toDateInput.value || 'empty',
                                    readOnly: toDateInput.readOnly,
                                    disabled: toDateInput.disabled
                                } : null
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: 'Date input error'});
                        }
                    })()
                ");

                var dateInputInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(dateInputResult));

                if (!(bool)dateInputInfo.success)
                {
                    LogMessage($"ERROR finding date input: {dateInputInfo.error}");
                    if (dateInputInfo.available != null)
                    {
                        LogMessage("Available input fields:");
                        foreach (var input in dateInputInfo.available)
                        {
                            LogMessage($"  - ID: '{input.id}', Name: '{input.name}', Placeholder: '{input.placeholder}', Value: '{input.value}'");
                        }
                    }
                    throw new Exception($"Date input not found: {dateInputInfo.error}");
                }

                LogMessage($"✓ From Date input found: ID='{dateInputInfo.fromDate.id}', Name='{dateInputInfo.fromDate.name}', Type='{dateInputInfo.fromDate.type}'");
                LogMessage($"  Current value: '{dateInputInfo.fromDate.currentValue}', ReadOnly: {dateInputInfo.fromDate.readOnly}, Disabled: {dateInputInfo.fromDate.disabled}");
                
                if (dateInputInfo.toDate != null)
                {
                    LogMessage($"✓ To Date input found: ID='{dateInputInfo.toDate.id}', Name='{dateInputInfo.toDate.name}', Type='{dateInputInfo.toDate.type}'");
                    LogMessage($"  Current value: '{dateInputInfo.toDate.currentValue}', ReadOnly: {dateInputInfo.toDate.readOnly}, Disabled: {dateInputInfo.toDate.disabled}");
                }
                else
                {
                    LogMessage("⚠️ To Date input not found - will only set From Date");
                }

                // Step 2: Set both date values
                var setDateResult = await webView.CoreWebView2.ExecuteScriptAsync($@"
                    (function() {{
                        try {{
                            var fromDateInput = document.querySelector('input[name=""DMNDateDateRangeControl4392_FromDateCalender_DTB""]') ||
                                              document.querySelector('input[name*=""FromDateCalender""], input[name*=""FromDate""]') ||
                                              document.querySelector('input[type=""text""][name*=""Date""]');
                            
                            var toDateInput = document.querySelector('input[name*=""ToDateCalender""], input[name*=""ToDate""]') ||
                                            document.querySelectorAll('input[type=""text""][name*=""Date""]')[1];
                            
                            if (!fromDateInput) {{
                                return JSON.stringify({{success: false, error: 'Date input disappeared'}});
                            }}

                            var fromOldValue = fromDateInput.value;
                            var toOldValue = toDateInput ? toDateInput.value : 'N/A';
                            var startDate = '{startDateString}';
                            var endDate = '{endDateString}';
                            
                            // Set the from date value
                            fromDateInput.value = startDate;
                            
                            // Trigger events for from date
                            var inputEvent = new Event('input', {{ bubbles: true }});
                            fromDateInput.dispatchEvent(inputEvent);
                            
                            var changeEvent = new Event('change', {{ bubbles: true }});
                            fromDateInput.dispatchEvent(changeEvent);
                            
                            // Set the to date if available
                            var toDateSet = false;
                            if (toDateInput) {{
                                toDateInput.value = endDate;
                                toDateInput.dispatchEvent(new Event('input', {{ bubbles: true }}));
                                toDateInput.dispatchEvent(new Event('change', {{ bubbles: true }}));
                                toDateInput.dispatchEvent(new Event('blur', {{ bubbles: true }}));
                                toDateSet = true;
                            }}
                            
                            var blurEvent = new Event('blur', {{ bubbles: true }});
                            fromDateInput.dispatchEvent(blurEvent);

                            return JSON.stringify({{
                                success: true,
                                fromDate: {{
                                    oldValue: fromOldValue,
                                    targetDate: startDate,
                                    newValue: fromDateInput.value,
                                    valueChanged: fromOldValue !== fromDateInput.value
                                }},
                                toDate: toDateSet ? {{
                                    oldValue: toOldValue,
                                    targetDate: endDate,
                                    newValue: toDateInput.value,
                                    valueChanged: toOldValue !== toDateInput.value
                                }} : null,
                                message: 'Date values set successfully'
                            }});
                        }} catch (ex) {{
                            return JSON.stringify({{success: false, error: 'Date setting error'}});
                        }}
                    }})()
                ");

                var setDateInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(setDateResult));

                if (!(bool)setDateInfo.success)
                {
                    LogMessage($"ERROR setting date: {setDateInfo.error}");
                    throw new Exception($"Date setting failed: {setDateInfo.error}");
                }

                LogMessage($"✓ From Date set: '{setDateInfo.fromDate.oldValue}' → '{setDateInfo.fromDate.newValue}'");
                LogMessage($"  Value changed: {setDateInfo.fromDate.valueChanged}, Target: '{setDateInfo.fromDate.targetDate}'");
                
                if (setDateInfo.toDate != null)
                {
                    LogMessage($"✓ To Date set: '{setDateInfo.toDate.oldValue}' → '{setDateInfo.toDate.newValue}'");
                    LogMessage($"  Value changed: {setDateInfo.toDate.valueChanged}, Target: '{setDateInfo.toDate.targetDate}'");
                }

                // Step 3: Verify the date value stuck by reading it back
                await Task.Delay(500); // Brief pause to let change events complete

                var verificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var fromDateInput = document.querySelector('input[name=""DMNDateDateRangeControl4392_FromDateCalender_DTB""]') ||
                                              document.querySelector('input[name*=""FromDateCalender""], input[name*=""FromDate""]') ||
                                              document.querySelector('input[type=""text""][name*=""Date""]');
                            
                            if (!fromDateInput) {
                                return JSON.stringify({success: false, error: 'Date input disappeared during verification'});
                            }

                            return JSON.stringify({
                                success: true,
                                currentValue: fromDateInput.value,
                                isEmpty: fromDateInput.value === '' || fromDateInput.value === null
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: 'Date verification error'});
                        }
                    })()
                ");

                var verificationInfo = JsonConvert.DeserializeObject<dynamic>(
                    UnwrapScriptJson(verificationResult));

                if ((bool)verificationInfo.success)
                {
                    LogMessage($"✓ Date setting verified: Current value = '{verificationInfo.currentValue}'");
                    
                    // Verify it matches our target date
                    if (verificationInfo.currentValue.ToString() == startDateString)
                    {
                        LogMessage($"✓ Verified date matches target: {startDateString}");
                    }
                    else if (!(bool)verificationInfo.isEmpty)
                    {
                        LogMessage($"⚠️ WARNING: Date value '{verificationInfo.currentValue}' does not exactly match target '{startDateString}' but is not empty");
                    }
                    else
                    {
                        LogMessage($"❌ ERROR: Date input is empty after setting");
                        throw new Exception("Date input is empty after setting");
                    }
                }
                else
                {
                    LogMessage($"ERROR verifying date setting: {verificationInfo.error}");
                    throw new Exception($"Date setting verification failed: {verificationInfo.error}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Date range setting failed - {ex.Message}");
                throw;
            }
        }

        private async Task GenerateAttendanceReport()
        {
            UpdateStatus("Generating attendance report...", 70);
            LogMessage("Starting attendance report generation...");

            // Only the report's own View/Generate button: never fall back to an arbitrary submit button.
            var token = await MarkCurrentPageAsync();
            var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                (function() {
                    try {
                        var b = document.getElementById('ViewReportImageButton') ||
                                document.querySelector('input[id*=""ViewReport""], input[name*=""ViewReport""]');
                        if (!b) {
                            var all = document.querySelectorAll('input[type=""submit""], input[type=""image""], input[type=""button""], button');
                            var available = [];
                            for (var i = 0; i < Math.min(all.length, 8); i++)
                                available.push((all[i].id || all[i].name || 'no-id') + ' | ' + (all[i].value || all[i].textContent || '').trim());
                            return JSON.stringify({success: false, error: 'ViewReport (Generate) button not found', available: available});
                        }
                        if (b.disabled) return JSON.stringify({success: false, error: 'Generate button is disabled'});
                        b.click();
                        return JSON.stringify({success: true, id: b.id || b.name});
                    } catch (ex) {
                        return JSON.stringify({success: false, error: ex.message});
                    }
                })()
            ");

            var clickInfo = JsonConvert.DeserializeObject<dynamic>(UnwrapScriptJson(clickResult));
            if (!(bool)clickInfo.success)
            {
                if (clickInfo.available != null)
                    foreach (var btn in clickInfo.available) LogMessage($"  available button: {btn}");
                throw new Exception($"Generate button click failed: {clickInfo.error}");
            }

            LogMessage($"✓ Generate button clicked ({clickInfo.id}) - waiting for the report...");
            UpdateStatus("Waiting for attendance report...", 80);
            await WaitForFreshPageAsync(token, JsReportViewerPresent, "attendance report");
        }

        private async Task<List<MonthlyAttendanceEntry>> ParseMonthlyReportData()
        {
            var monthlyData = new List<MonthlyAttendanceEntry>();
                        LogMessage("▶ Starting single-pass monthly extraction (oReportDiv-first heuristic)...");

                                        var script = @"(function(){
                    var dbg = []; function step(msg){ dbg.push(msg); }
                    try {
                        step('STEP 1: Begin');
                        var rv = document.getElementById('ReportViewer1');
                        if(!rv){ step('ReportViewer1 not found'); return {success:false, error:'ReportViewer1 not found', debug:dbg}; }
                        step('STEP 2: ReportViewer1 located');

                 // STEP 3: Prefer div whose id ends with oReportDiv, then fallback to oReportCell
                 var container = rv.querySelector('div[id$=oReportDiv]');
                 if(container){ step('STEP 3: Found container via id$=oReportDiv: '+ container.id); }
                 if(!container){ container = rv.querySelector('div[id$=oReportCell]'); if(container) step('STEP 3: Fallback container via id$=oReportCell: '+container.id); }
                 // If still not found try broader contains match inside main doc
                 if(!container){ var c2 = rv.querySelector('div[id*=oReportDiv]'); if(c2){ container = c2; step('STEP 3: Contains match oReportDiv: '+container.id);} }
                 if(!container){ var c3 = rv.querySelector('div[id*=oReportCell]'); if(c3){ container = c3; step('STEP 3: Contains match oReportCell: '+container.id);} }

                 // STEP 4: If not found yet, attempt accessible iframes (single pass only)
                 if(!container){
                      var ifr = rv.querySelectorAll('iframe');
                      step('STEP 4: Container not in main doc, scanning '+ifr.length+' iframe(s)');
                      for(var i=0;i<ifr.length && !container;i++){
                         try {
                            var idoc = ifr[i].contentDocument || ifr[i].contentWindow.document;
                            if(!idoc) continue;
                            var cand = idoc.querySelector('div[id$=oReportDiv]') || idoc.querySelector('div[id$=oReportCell]') || idoc.querySelector('div[id*=oReportDiv]') || idoc.querySelector('div[id*=oReportCell]');
                            if(cand){ container = cand; step('STEP 4: Found container inside iframe#'+(ifr[i].id||i)+': '+cand.id); }
                         } catch(e){ step('STEP 4: iframe access blocked: '+ e.message); }
                      }
                 }

                if(!container){ step('FAIL: No report container (oReportDiv / oReportCell) found'); return {success:false, error:'Report container not found', debug:dbg}; }

                 // STEP 5: Collect tables within the chosen container
                 var tables = Array.prototype.slice.call(container.querySelectorAll('table'));
                 step('STEP 5: Tables inside container: '+ tables.length);
                if(!tables.length){ return {success:false, error:'No tables in container', debug:dbg}; }

                 // STEP 6: Score tables similar to daily logic
                 function scoreTable(tb){
                     var rows = tb.rows; if(!rows || rows.length < 2) return 0;
                     var rowCount = rows.length;
                     var numericIds = 0, timeCells = 0, dateCells=0, textCells=0, maxCols=0;
                     var dateRegex = /\b\d{1,2}[\/-][A-Za-z]{3}[\/-]?\d{2,4}\b|\b\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}\b/; /* keep double escapes here because we are inside C# verbatim? Actually we want JS to see \b for word boundary. */
                     var timeRegex = /\b\d{1,2}:\d{2}\b/;
                     for(var r=0; r<Math.min(rows.length, 60); r++){
                        var cells = rows[r].cells; if(!cells) continue; if(cells.length>maxCols) maxCols=cells.length;
                        for(var c=0;c<cells.length;c++){
                          var txt = (cells[c].innerText||'').trim();
                          if(/^\d{3,}$/.test(txt)) numericIds++;
                          else if(timeRegex.test(txt)) timeCells++;
                          else if(dateRegex.test(txt)) dateCells++;
                          else if(txt) textCells++;
                        }
                     }
                     return rowCount*3 + numericIds*2 + timeCells*2 + dateCells*3 + maxCols;
                 }
                 var scored = tables.map(function(t,i){ return {idx:i, score:scoreTable(t), rows:t.rows.length, el:t}; });
                 scored.sort(function(a,b){ return b.score - a.score; });
                 var best = scored[0];
                 step('STEP 6: Best table index '+best.idx+' score='+best.score+' rows='+best.rows);
                if(!best || best.score === 0){ return {success:false, error:'No suitable table (score=0)', debug:dbg}; }

                 var table = best.el;

                 // STEP 7: Detect header row (first 5 rows) looking for employee/date keywords
                 var headerRowIndex = 0; var headerTexts=[];
                 for(var hr=0; hr<Math.min(table.rows.length,5); hr++){
                     var cells = table.rows[hr].cells; if(!cells || cells.length<3) continue;
                     var texts=[]; for(var hc=0; hc<cells.length; hc++){ var t=(cells[hc].innerText||'').replace(/\s+/g,' ').trim().toLowerCase(); texts.push(t); }
                     var joined = texts.join(' | ');
                     if(/employee/.test(joined) && /(id|name)/.test(joined)) { headerRowIndex = hr; headerTexts=texts; step('STEP 7: Header row found at '+hr+' -> '+texts.join(' || ')); break; }
                     if(hr===0){ headerTexts=texts; }
                 }
                if(!headerTexts.length){ return {success:false, error:'Header row not detected', debug:dbg}; }

                 // STEP 8: Build column map (extended)
                 var map={};
                 headerTexts.forEach(function(h,i){
                     var base = h.replace(/\s+/g,' ');
                     if(h.includes('employee') && h.includes('id')) map.empId=i;
                     else if(h.includes('employee') && h.includes('name')) map.name=i;
                     else if(h==='date' || h.includes('date')) map.date=i;
                     else if(h.includes('swipe') && h.includes('count')) map.swipeCount=i;
                     else if(h.includes('in') && h.includes('time')) map.inTime=i;
                     else if(h.includes('out') && h.includes('time')) map.outTime=i;
                     else if(h.includes('total') && h.includes('working') && h.includes('swipe')) map.totalSwipe=i; // Total Working Hours - Swipes
                     else if(h.includes('actual') && h.includes('working') && h.includes('swipe') && !h.includes('wfh') && !h.includes('+')) map.actualSwipe=i; // Actual Working Hours - Swipes (A)
                     else if(h.includes('total') && h.includes('working') && h.includes('wfh')) map.totalWFH=i; // Total Working Hours - WFH (B)
                     else if(h.includes('actual') && h.includes('working') && h.includes('wfh') && !h.includes('+')) map.actualWFH=i; // Actual Working Hours - WFH (B)
                     else if(h.includes('actual') && h.includes('swipe') && h.includes('wfh') && h.includes('+')) map.actualCombined=i; // Combined actual (A)+(B)
                     else if(h.includes('actual') && h.includes('hour') && !map.actualGeneric) map.actualGeneric=i; // fallback generic actual
                     else if(h.includes('total') && h.includes('hour') && !map.totalGeneric) map.totalGeneric=i; // fallback generic total
                     else if(h.includes('first') && h.includes('half') && h.includes('status')) map.firstHalfStatus=i;
                     else if(h.includes('second') && h.includes('half') && h.includes('status')) map.secondHalfStatus=i;
                     else if(h==='status' || (h.includes('status') && !map.status)) map.status=i;
                 });
                 step('STEP 8: Column map '+JSON.stringify(map));
                 if(map.empId==null){
                     for(var c=0; c<headerTexts.length && map.empId==null; c++){
                         var numericHits=0; for(var rr=headerRowIndex+1; rr<Math.min(table.rows.length, headerRowIndex+15); rr++){ var cell=table.rows[rr].cells[c]; if(!cell) continue; var txt=(cell.innerText||'').trim(); if(/^\d{3,}$/.test(txt)) numericHits++; }
                         if(numericHits>=3){ map.empId=c; step('Fallback empId column='+c+' hits='+numericHits); }
                     }
                 }
                 if(map.name==null && map.empId!=null){
                     if(map.empId+1 < headerTexts.length){ map.name=map.empId+1; step('Fallback name column chosen index '+map.name); }
                 }
                 if(map.date==null){
                     var dateRegex2=/\d{1,2}[\/-]\d{1,2}[\/-]\d{2,4}/;
                     for(var c2=0; c2<headerTexts.length && map.date==null; c2++){
                        for(var rr2=headerRowIndex+1; rr2<Math.min(table.rows.length, headerRowIndex+12); rr2++){
                          var cell2=table.rows[rr2].cells[c2]; if(!cell2) continue; var tx=(cell2.innerText||'').trim(); if(dateRegex2.test(tx)){ map.date=c2; step('Fallback date column='+c2+' sample='+tx); break; }
                        }
                     }
                 }
                if(map.empId==null || map.name==null || map.date==null){ return {success:false, error:'Essential columns unresolved', debug:dbg, map:map}; }

                 // STEP 9: Extract rows
                 var results=[]; var totalRows=table.rows.length; step('STEP 9: Total table rows '+totalRows);
                 for(var r=headerRowIndex+1; r<table.rows.length; r++){
                     var row=table.rows[r]; var cells=row.cells; if(!cells || cells.length <= map.name) continue;
                     var idTxt=(cells[map.empId].innerText||'').trim();
                     if(!/^\d{3,}$/.test(idTxt)) continue;
                     var entry={
                         employeeId:idTxt,
                         employeeName:(cells[map.name] && cells[map.name].innerText||'').trim(),
                         date:(cells[map.date] && cells[map.date].innerText||'').trim(),
                         inTime: map.inTime!=null ? (cells[map.inTime].innerText||'').trim() : '',
                         outTime: map.outTime!=null ? (cells[map.outTime].innerText||'').trim() : '',
                         swipeCount: map.swipeCount!=null ? (cells[map.swipeCount].innerText||'').trim() : '',
                         totalSwipeHours: map.totalSwipe!=null ? (cells[map.totalSwipe].innerText||'').trim() : (map.totalGeneric!=null ? (cells[map.totalGeneric].innerText||'').trim():''),
                         actualSwipeHours: map.actualSwipe!=null ? (cells[map.actualSwipe].innerText||'').trim() : (map.actualGeneric!=null ? (cells[map.actualGeneric].innerText||'').trim():''),
                         totalWFHHours: map.totalWFH!=null ? (cells[map.totalWFH].innerText||'').trim() : '',
                         actualWFHHours: map.actualWFH!=null ? (cells[map.actualWFH].innerText||'').trim() : '',
                         actualCombinedHours: map.actualCombined!=null ? (cells[map.actualCombined].innerText||'').trim() : '',
                         totalHours: map.totalSwipe!=null ? (cells[map.totalSwipe].innerText||'').trim() : (map.totalGeneric!=null ? (cells[map.totalGeneric].innerText||'').trim() : ''),
                         actualWorkHours: map.actualCombined!=null ? (cells[map.actualCombined].innerText||'').trim() : (map.actualSwipe!=null ? (cells[map.actualSwipe].innerText||'').trim() : (map.actualGeneric!=null ? (cells[map.actualGeneric].innerText||'').trim():'')),
                         firstHalfStatus: map.firstHalfStatus!=null ? (cells[map.firstHalfStatus].innerText||'').trim() : '',
                         secondHalfStatus: map.secondHalfStatus!=null ? (cells[map.secondHalfStatus].innerText||'').trim() : '',
                         status: map.status!=null ? (cells[map.status].innerText||'').trim() : ''
                     };
                     results.push(entry);
                     if(results.length>=800) break;
                 }
                 step('STEP 10: Extracted data rows '+results.length);
                return {success:true, count:results.length, entries:results, debug:dbg};
              } catch(e){
                step('ERROR: '+e.message);
                return {success:false, error:e.message, debug:dbg};
              }
            })();";

                        // The report renders asynchronously; poll until the table is readable.
                        dynamic? parsed = null;
                        string lastError = "no result";
                        var pollTimer = System.Diagnostics.Stopwatch.StartNew();
                        while (true)
                        {
                            string raw = await webView.CoreWebView2.ExecuteScriptAsync(script);
                            // The script returns an object (not a JSON string), so raw is the object's JSON.
                            parsed = string.IsNullOrWhiteSpace(raw) || raw == "null"
                                ? null
                                : Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(raw);
                            if (parsed?.success == true) break;
                            lastError = parsed?.error != null ? parsed.error.ToString() : "script returned no result";
                            if (pollTimer.ElapsedMilliseconds > 45000)
                                throw new Exception($"Monthly report could not be read after 45s: {lastError}");
                            await Task.Delay(1000);
                        }

                        if (parsed?.debug != null)
                        {
                                LogMessage("=== JS DEBUG (single-pass) ===");
                                foreach (var d in parsed.debug)
                                {
                                        LogMessage($"JS: {d}");
                                }
                                LogMessage("=== END JS DEBUG ===");
                        }

                        if (parsed?.success == true && parsed.entries != null)
                        {
                            foreach (var item in parsed.entries)
                            {
                                // Defensive accessors
                                string GetStr(dynamic obj, string name)
                                {
                                    try { var val = obj?[name]; return val == null ? string.Empty : val.ToString(); } catch { return string.Empty; }
                                }

                                int swipeCountInt = 0;
                                var swipeText = GetStr(item, "swipeCount");
                                if (!string.IsNullOrWhiteSpace(swipeText) && int.TryParse(new string(swipeText.Where(char.IsDigit).ToArray()), out var sc)) swipeCountInt = sc;

                                monthlyData.Add(new MonthlyAttendanceEntry
                                {
                                    EmployeeId = GetStr(item, "employeeId"),
                                    EmployeeName = GetStr(item, "employeeName"),
                                    Date = GetStr(item, "date"),
                                    SwipeCount = swipeCountInt,
                                    InTime = GetStr(item, "inTime"),
                                    OutTime = GetStr(item, "outTime"),
                                    TotalHours = GetStr(item, "totalHours"),
                                    ActualWorkHours = GetStr(item, "actualWorkHours"),
                                    TotalWFHHours = GetStr(item, "totalWFHHours"),
                                    ActualWFHHours = GetStr(item, "actualWFHHours"),
                                    Status = GetStr(item, "status"),
                                    FirstHalfStatus = GetStr(item, "firstHalfStatus"),
                                    SecondHalfStatus = GetStr(item, "secondHalfStatus")
                                });
                            }
                            LogMessage($"✅ Extracted {monthlyData.Count} monthly rows (single-pass, extended columns)");
                        }

                        // Every row must belong to the requested employee; anything else means the
                        // dropdown selection did not take effect.
                        var wantedId = RequireEmployeeId();
                        var otherIds = monthlyData.Select(m => m.EmployeeId.Trim()).Where(id => id != wantedId).Distinct().ToList();
                        if (otherIds.Count > 0)
                            throw new Exception($"Monthly report contains employee(s) {string.Join(", ", otherIds)} but {wantedId} was requested.");

                        if (monthlyData.Count == 0)
                        {
                                LogMessage("ℹ Monthly report table was read but contains no rows for this period");
                        }

                        return monthlyData;
        }

        // ================= FULL REPORT SUPPORT =================
        private List<MonthlyAttendanceEntry> _monthlyReportCache = new();
        private List<WorkHoursCalculation> _todayWorkHoursCalculations = new();
        private List<WorkHoursCalculation> _yesterdayWorkHoursCalculations = new();

        // Call this after monthly extraction succeeds
        private void UpdateMonthlyReportCache(IEnumerable<MonthlyAttendanceEntry> entries)
        {
            _monthlyReportCache = entries?.ToList() ?? new List<MonthlyAttendanceEntry>();
        }

        // Hook points for existing today / yesterday calculation results
        private void UpdateTodayCalculations(IEnumerable<WorkHoursCalculation> entries)
        {
            _todayWorkHoursCalculations = entries?.ToList() ?? new List<WorkHoursCalculation>();
            LogMessage($"[CACHE] Today calculations updated: {_todayWorkHoursCalculations.Count} entries");
            foreach (var calc in _todayWorkHoursCalculations)
            {
                LogMessage($"[CACHE] Today: {calc.Date:yyyy-MM-dd} = {calc.WorkingHoursDisplay}");
            }
        }
        private void UpdateYesterdayCalculations(IEnumerable<WorkHoursCalculation> entries)
        {
            _yesterdayWorkHoursCalculations = entries?.ToList() ?? new List<WorkHoursCalculation>();
            LogMessage($"[CACHE] Yesterday calculations updated: {_yesterdayWorkHoursCalculations.Count} entries");
            foreach (var calc in _yesterdayWorkHoursCalculations)
            {
                LogMessage($"[CACHE] Yesterday: {calc.Date:yyyy-MM-dd} = {calc.WorkingHoursDisplay}");
            }
        }

        private async void BtnFullReport_Click(object sender, RoutedEventArgs e)
        {
            if (!TryBeginAutomation("Full report")) return;
            try
            {
                LogMessage("Full Report: running Monthly -> Today -> Yesterday...");
                var failures = await RunFullCalendarAutomation();

                bool anyData = _monthlyReportCache.Count > 0 || _todayWorkHoursCalculations.Count > 0 || _yesterdayWorkHoursCalculations.Count > 0;
                if (failures.Count > 0)
                {
                    var msg = "Some parts of the full report failed:\n\n- " + string.Join("\n- ", failures) +
                              (anyData ? "\n\nThe calendar shows only the parts that succeeded." : "");
                    UpdateStatus($"Full report finished with {failures.Count} error(s)", 0);
                    MessageBox.Show(msg, "Full Report - Errors", MessageBoxButton.OK, MessageBoxImage.Warning);
                }
                else
                {
                    UpdateStatus("Full report completed", 100);
                }

                if (anyData)
                    ShowFullCalendarWindow();
            }
            catch (Exception ex)
            {
                LogMessage($"Full Report failed: {ex}");
                UpdateStatus($"Full report failed: {ex.Message}", 0);
                MessageBox.Show($"Full report failed: {ex.Message}", "Full Report Error", MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                EndAutomation();
            }
        }

        /// <summary>
        /// Runs Monthly, Today and Yesterday independently so one failure doesn't hide the others.
        /// Caches are cleared first so a failed phase can never show data from a previous run.
        /// Returns the list of failures (empty when everything succeeded).
        /// </summary>
        private async Task<List<string>> RunFullCalendarAutomation()
        {
            var failures = new List<string>();
            _monthlyReportCache = new List<MonthlyAttendanceEntry>();
            _todayWorkHoursCalculations = new List<WorkHoursCalculation>();
            _yesterdayWorkHoursCalculations = new List<WorkHoursCalculation>();

            try
            {
                LogMessage("[CALENDAR] Monthly attendance...");
                var monthly = await ExtractMonthlyAttendanceData();
                UpdateMonthlyReportCache(monthly);
                LogMessage($"[CALENDAR] Monthly rows: {monthly.Count}");
            }
            catch (Exception ex)
            {
                LogMessage($"[CALENDAR] Monthly failed: {ex.Message}");
                failures.Add($"Monthly: {ex.Message}");
            }

            foreach (var isToday in new[] { true, false })
            {
                var label = isToday ? "Today" : "Yesterday";
                try
                {
                    var calc = await DailyHoursCalculationAsync(isToday);
                    if (isToday) UpdateTodayCalculations(calc); else UpdateYesterdayCalculations(calc);
                }
                catch (Exception ex)
                {
                    LogMessage($"[CALENDAR] {label} failed: {ex.Message}");
                    failures.Add(ex.Message);
                }
            }

            return failures;
        }

        private void ShowFullCalendarWindow()
        {
            var model = Services.FullReportBuilder.Build(
                _monthlyReportCache,
                _todayWorkHoursCalculations,
                _yesterdayWorkHoursCalculations);
            var win = new FullReportWindow(model) { Owner = this };
            win.Show();
        }

        /// <summary>
        /// Runs the swipe log for one day and converts it to a calendar entry.
        /// Returns an empty list when the day has no swipes (monthly data then stands); throws on failure.
        /// </summary>
        private async Task<List<WorkHoursCalculation>> DailyHoursCalculationAsync(bool isToday)
        {
            var label = isToday ? "Today" : "Yesterday";
            var day = isToday ? DateTime.Today : DateTime.Today.AddDays(-1);
            var entries = await RunDailyReportAsync(isToday);
            if (entries.Count == 0)
                return new List<WorkHoursCalculation>();

            var result = WorkHoursCalculator.Calculate(entries, DateTime.Now);
            foreach (var e in result.Labeled)
                LogMessage($"[LABEL] {e.Date} {e.Time} | {e.Direction} | {e.MachineName} -> {e.Activity} ({e.Duration})");
            if (result.SkippedRows.Count > 0)
                LogMessage($"⚠ {label}: skipped {result.SkippedRows.Count} unreadable row(s): {string.Join("; ", result.SkippedRows)}");
            LogMessage($"[DAILY] {label} work time: {result.WorkTimeDisplay}" +
                       (result.OngoingTime > TimeSpan.Zero ? $" (incl. {WorkHoursCalculator.Format(result.OngoingTime)} ongoing)" : ""));

            return new List<WorkHoursCalculation>
            {
                new WorkHoursCalculation
                {
                    Date = day,
                    WorkingHoursDisplay = result.WorkTimeDisplay,
                    TotalWorkingHoursDisplay = result.WorkTimeDisplay,
                    Status = result.WorkTime > TimeSpan.Zero ? "Present" : "No Work Time"
                }
            };
        }

        #endregion
    }
}