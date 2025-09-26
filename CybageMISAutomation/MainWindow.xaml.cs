using System.Windows;
using CybageMISAutomation.Models;
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
        private const string MIS_URL = "https://cybagemis.cybage.com/Report%20Builder/RPTN/ReportPage.aspx";
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
            // dataGridResults grid is currently commented out in XAML; attach if reinstated
            if (this.FindName("dataGridResults") is DataGrid dg)
                dg.ItemsSource = _swipeLogData;
            
            // Create and show log window
            _logWindow = new LogWindow();
            _logWindow.Show();
            
            InitializeWebView();
        }

        private async void InitializeWebView()
        {
            try
            {
                UpdateStatus("Initializing WebView2 with Windows Authentication...", 10);

                // Configure WebView2 environment with Windows authentication
                var options = new CoreWebView2EnvironmentOptions();
                options.AdditionalBrowserArguments = "--enable-features=msIntegratedAuth --auth-server-allowlist=*.cybage.com --auth-negotiate-delegate-allowlist=*.cybage.com --disable-web-security --disable-site-isolation-trials";
                options.AllowSingleSignOnUsingOSPrimaryAccount = true;
                
                var userDataFolder = Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData) + "\\CybageReportViewer";
                var env = await CoreWebView2Environment.CreateAsync(null, userDataFolder, options);
                await webView.EnsureCoreWebView2Async(env);

                // Set up event handlers
                webView.CoreWebView2.NavigationStarting += WebView_NavigationStarting;
                webView.CoreWebView2.NavigationCompleted += WebView_NavigationCompleted;
                webView.CoreWebView2.DOMContentLoaded += WebView_DOMContentLoaded;
                webView.CoreWebView2.DocumentTitleChanged += WebView_DocumentTitleChanged;

                _isWebViewInitialized = true;
                UpdateStatus("WebView2 initialized successfully with Windows Authentication.", 20);
                LogMessage("WebView2 initialization complete. Ready for navigation.");

                btnStartAutomation.IsEnabled = true;
                btnStartFullAutomation.IsEnabled = true; // renamed caption only
                btnMonthlyReport.IsEnabled = true;
                if (this.FindName("btnFullReport") is Button frBtnInit)
                    frBtnInit.IsEnabled = true; // Always enabled per new requirement
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
                LogMessage($"Target URL: {MIS_URL}");

                // Navigate to the MIS page
                webView.CoreWebView2.Navigate(MIS_URL);
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
                var cleanJson = pageInfoJson.Trim('"').Replace("\\\"", "\"");
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

                var expandCheck = JsonConvert.DeserializeObject<dynamic>(isExpandedResult.Trim('"').Replace("\\\"", "\""));
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

                    var clickResponse = JsonConvert.DeserializeObject<dynamic>(clickResult.Trim('"').Replace("\\\"", "\""));
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

                var swipeLogInfo = JsonConvert.DeserializeObject<dynamic>(swipeLogCheck.Trim('"').Replace("\\\"", "\""));

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
                var linkVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var swipeLogLink = document.getElementById('TempleteTreeViewt32');
                        if (swipeLogLink) {
                            return JSON.stringify({
                                found: true,
                                text: swipeLogLink.textContent || swipeLogLink.innerText,
                                href: swipeLogLink.getAttribute('href') || '',
                                onclick: swipeLogLink.getAttribute('onclick') || '',
                                visible: window.getComputedStyle(swipeLogLink).display !== 'none'
                            });
                        }
                        return JSON.stringify({found: false, error: 'Link not found'});
                    })()
                ");

                var linkInfo = JsonConvert.DeserializeObject<dynamic>(
                    linkVerificationResult.Trim('"').Replace("\\\"", "\""));

                if (!(bool)linkInfo.found)
                {
                    LogMessage("ERROR: Swipe Log link not found or disappeared");
                    UpdateStatus("Swipe Log link not available", 0);
                    return;
                }

                LogMessage($"✓ Swipe Log link verified: '{linkInfo.text}'");
                LogMessage($"  OnClick: {linkInfo.onclick}");

                // Click the swipe log link
                var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var swipeLogLink = document.getElementById('TempleteTreeViewt32');
                            if (!swipeLogLink) {
                                return JSON.stringify({success: false, error: 'Link not found'});
                            }

                            // Simulate the click - this should trigger the __doPostBack
                            swipeLogLink.click();
                            
                            return JSON.stringify({
                                success: true, 
                                message: 'Swipe log link clicked successfully'
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: ex.message});
                        }
                    })()
                ");

                var clickResponse = JsonConvert.DeserializeObject<dynamic>(
                    clickResult.Trim('"').Replace("\\\"", "\""));

                if ((bool)clickResponse.success)
                {
                    LogMessage("✓ Swipe Log link clicked successfully");
                    UpdateStatus("Waiting for report page to load...", 60);

                    // Wait for the page to navigate/change after the postback
                    await Task.Delay(3000);

                    // Verify we've navigated to the report configuration page
                    await VerifyReportPage();
                }
                else
                {
                    LogMessage($"ERROR clicking swipe log link: {clickResponse.error}");
                    UpdateStatus("Failed to click swipe log link", 0);
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
                    pageVerificationResult.Trim('"').Replace("\\\"", "\""));

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
                    dropdownResult.Trim('"').Replace("\\\"", "\""));

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
                        generateResult.Trim('"').Replace("\\\"", "\""));

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
                    reportVerificationResult.Trim('"').Replace("\\\"", "\""));

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
                    extractResult.Trim('"').Replace("\\\"", "\""));

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
                webView.CoreWebView2.Navigate("https://cybagemis.cybage.com/Report%20Builder/RPTN/ReportPage.aspx");
                
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
            if (_isFullAutomationRunning)
            {
                LogMessage("❌ Full automation is already running. Please wait for completion.");
                return;
            }

            await StartFullAutomation();
        }

        public async Task StartFullAutomation()
        {
            try
            {
                _isFullAutomationRunning = true;
                _todayData.Clear();
                _yesterdayData.Clear();

                // Disable all buttons during automation
                SetButtonsEnabled(false);
                btnStartFullAutomation.IsEnabled = false;

                LogMessage("🚀 Starting FULL AUTOMATION workflow...");
                LogMessage("  This will extract both Today and Yesterday data automatically");
                UpdateStatus("Full automation started...", 0);

                // Phase 1: Extract Today's Data
                LogMessage("📅 PHASE 1: Extracting Today's Data");
                await ExecuteTodayAutomation();

                if (_todayData.Count == 0)
                {
                    LogMessage("⚠ Warning: No data extracted for Today. Continuing with Yesterday...");
                }

                // Wait before starting yesterday automation
                await Task.Delay(2000);

                // Phase 2: Extract Yesterday's Data  
                LogMessage("📅 PHASE 2: Extracting Yesterday's Data");
                await ExecuteYesterdayAutomation();

                if (_yesterdayData.Count == 0)
                {
                    LogMessage("⚠ Warning: No data extracted for Yesterday.");
                }

                // Phase 3: Show Comparison Results
                LogMessage("📊 PHASE 3: Displaying Results");
                ShowComparisonResults();

                LogMessage($"✅ FULL AUTOMATION COMPLETED!");
                LogMessage($"  Today: {_todayData.Count} records");
                LogMessage($"  Yesterday: {_yesterdayData.Count} records");
                UpdateStatus($"Full automation completed - Today: {_todayData.Count}, Yesterday: {_yesterdayData.Count}", 100);
            }
            catch (Exception ex)
            {
                LogMessage($"❌ FULL AUTOMATION FAILED: {ex.Message}");
                UpdateStatus("Full automation failed", 0);
                MessageBox.Show($"Full automation failed: {ex.Message}", "Automation Error",
                                MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                _isFullAutomationRunning = false;
                SetButtonsEnabled(true);
                btnStartFullAutomation.IsEnabled = true;
            }
        }

        private async Task ExecuteTodayAutomation()
        {
            try
            {
                // Step 1: Navigate to main page
                LogMessage("1️⃣ Navigating to main page...");
                webView.CoreWebView2.Navigate(MIS_URL);
                await WaitForPageLoad();

                // Step 2: Expand tree
                LogMessage("2️⃣ Expanding Leave Management tree...");
                await ExecuteExpandTree();
                await Task.Delay(500); // Brief pause

                // Step 3: Click swipe log link
                LogMessage("3️⃣ Clicking swipe log link...");
                await ExecuteClickSwipeLog();

                // Step 4: Select Today
                LogMessage("4️⃣ Selecting Today from dropdown...");
                cmbReportType.SelectedIndex = 0; // Today
                await Task.Delay(500);

                // Step 5: Generate report
                LogMessage("5️⃣ Generating Today's report...");
                await ExecuteGenerateReport();

                // Step 6: Extract data
                LogMessage("6️⃣ Extracting Today's data...");
                _todayData = await ExecuteDataExtraction();
                LogMessage($"  ✅ Today's data extracted: {_todayData.Count} records");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Error in Today automation: {ex.Message}");
                throw;
            }
        }

        private async Task ExecuteYesterdayAutomation()
        {
            try
            {
                // Step 7: Reset to main page
                LogMessage("7️⃣ Resetting to main page...");
                webView.CoreWebView2.Navigate(MIS_URL);
                await WaitForPageLoad();

                // Step 8: Expand tree
                LogMessage("8️⃣ Expanding Leave Management tree...");
                await ExecuteExpandTree();
                await Task.Delay(500); // Brief pause

                // Step 9: Click swipe log link
                LogMessage("9️⃣ Clicking swipe log link...");
                await ExecuteClickSwipeLog();

                // Step 10: Select Yesterday
                LogMessage("🔟 Selecting Yesterday from dropdown...");
                cmbReportType.SelectedIndex = 1; // Yesterday
                await Task.Delay(500);

                // Step 11: Generate report
                LogMessage("1️⃣1️⃣ Generating Yesterday's report...");
                await ExecuteGenerateReport();

                // Step 12: Extract data
                LogMessage("1️⃣2️⃣ Extracting Yesterday's data...");
                _yesterdayData = await ExecuteDataExtraction();
                LogMessage($"  ✅ Yesterday's data extracted: {_yesterdayData.Count} records");
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Error in Yesterday automation: {ex.Message}");
                throw;
            }
        }

        private async Task<List<SwipeLogEntry>> ExecuteDataExtraction()
        {
            // Reuse the existing data extraction logic but return the data instead of showing window
            var extractResult = await webView.CoreWebView2.ExecuteScriptAsync(GetDataExtractionScript());
            var extractInfo = JsonConvert.DeserializeObject<dynamic>(
                extractResult.Trim('"').Replace("\\\"", "\""));

            var entries = new List<SwipeLogEntry>();
            
            if ((bool)extractInfo.success)
            {
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
            }
            
            return entries;
        }

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

        private async Task WaitForPageLoad()
        {
            // Wait for navigation to complete with proper detection
            var maxWaitTime = 15000; // 15 seconds
            var checkInterval = 500; // 0.5 seconds
            var elapsed = 0;

            while (elapsed < maxWaitTime)
            {
                try
                {
                    var readyState = await webView.CoreWebView2.ExecuteScriptAsync("document.readyState");
                    if (readyState.Trim('"') == "complete")
                    {
                        LogMessage("Page loading completed successfully");
                        await Task.Delay(1000); // Additional wait for dynamic content
                        return;
                    }
                }
                catch
                {
                    // Continue waiting if script execution fails
                }

                await Task.Delay(checkInterval);
                elapsed += checkInterval;
            }

            LogMessage("⚠ Page load timeout - continuing anyway");
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
            // This calls the exact same logic as the manual BtnExpandTree_Click
            try
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

                var expandCheck = JsonConvert.DeserializeObject<dynamic>(isExpandedResult.Trim('"').Replace("\\\"", "\""));
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

                    var clickResponse = JsonConvert.DeserializeObject<dynamic>(clickResult.Trim('"').Replace("\\\"", "\""));
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
            }
        }

        private async Task ExecuteClickSwipeLog()
        {
            // Call the SAME function that works in manual mode
            LogMessage("🔗 Calling manual swipe log click function...");
            await CallManualSwipeLogClickFunction();
        }

        private async Task CallManualSwipeLogClickFunction()
        {
            try
            {
                UpdateStatus("Clicking 'Today's and Yesterday's Swipe Log' link...", 30);
                LogMessage("Starting swipe log link click process...");

                // First verify the swipe log link is still available and get its details (same as manual)
                var linkVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var swipeLogLink = document.getElementById('TempleteTreeViewt32');
                        if (swipeLogLink) {
                            return JSON.stringify({
                                found: true,
                                text: swipeLogLink.textContent || swipeLogLink.innerText,
                                href: swipeLogLink.getAttribute('href') || '',
                                onclick: swipeLogLink.getAttribute('onclick') || '',
                                visible: window.getComputedStyle(swipeLogLink).display !== 'none'
                            });
                        }
                        return JSON.stringify({found: false, error: 'Link not found'});
                    })()
                ");

                var linkInfo = JsonConvert.DeserializeObject<dynamic>(
                    linkVerificationResult.Trim('"').Replace("\\\"", "\""));

                if (!(bool)linkInfo.found)
                {
                    LogMessage("ERROR: Swipe Log link not found or disappeared");
                    throw new Exception("Swipe Log link not found or disappeared");
                }

                LogMessage($"✓ Swipe Log link verified: '{linkInfo.text}'");
                LogMessage($"  OnClick: {linkInfo.onclick}");

                // Click the swipe log link (same as manual)
                var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var swipeLogLink = document.getElementById('TempleteTreeViewt32');
                            if (!swipeLogLink) {
                                return JSON.stringify({success: false, error: 'Link not found'});
                            }

                            // Simulate the click - this should trigger the __doPostBack
                            swipeLogLink.click();
                            
                            return JSON.stringify({
                                success: true, 
                                message: 'Swipe log link clicked successfully'
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: ex.message});
                        }
                    })()
                ");

                var clickResponse = JsonConvert.DeserializeObject<dynamic>(
                    clickResult.Trim('"').Replace("\\\"", "\""));

                if ((bool)clickResponse.success)
                {
                    LogMessage("✓ Swipe Log link clicked successfully");
                    UpdateStatus("Waiting for report page to load...", 60);

                    // Wait for the page to navigate/change after the postback (same as manual)
                    await Task.Delay(3000);

                    // Verify we've navigated to the report configuration page (same as manual)
                    await VerifyReportPage();
                }
                else
                {
                    LogMessage($"ERROR clicking swipe log link: {clickResponse.error}");
                    throw new Exception($"Swipe log click failed: {clickResponse.error}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Swipe log click failed - {ex.Message}");
                throw;
            }
        }

        private async Task ExecuteGenerateReport()
        {
            // Call the SAME function that works in manual mode  
            LogMessage("📊 Calling manual report generation function...");
            await CallManualGenerateReportFunction();
        }

        private async Task CallManualGenerateReportFunction()
        {
            try
            {
                UpdateStatus("Setting up report parameters and generating report...", 40);
                LogMessage("Starting report generation process...");

                // Get the selected report type from UI (same as manual)
                var selectedItem = cmbReportType.SelectedItem as ComboBoxItem;
                var reportTypeValue = selectedItem?.Tag?.ToString() ?? "1"; // Default to Today
                var reportTypeName = selectedItem?.Content?.ToString() ?? "Today";

                LogMessage($"Selected report type: {reportTypeName} (Value: {reportTypeValue})");

                // First, set the day dropdown selection (same as manual)
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

                var dropdownInfo = JsonConvert.DeserializeObject<dynamic>(
                    dropdownResult.Trim('"').Replace("\\\"", "\""));

                if (!(bool)dropdownInfo.success)
                {
                    LogMessage($"ERROR setting dropdown: {dropdownInfo.error}");
                    throw new Exception($"Failed to set dropdown: {dropdownInfo.error}");
                }

                LogMessage($"✓ Dropdown set successfully to '{dropdownInfo.selectedText}' (Value: {dropdownInfo.finalValue})");
                if ((bool)dropdownInfo.selectionChanged)
                {
                    LogMessage($"  Changed from '{dropdownInfo.previousValue}' to '{dropdownInfo.finalValue}'");
                    await Task.Delay(1000); // Brief wait for any UI updates
                }

                // Now click the Generate Report button (same as manual)
                UpdateStatus("Dropdown set. Clicking Generate Report button...", 70);
                await Task.Delay(1000); // Wait for any postback or JavaScript to complete

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
                    generateResult.Trim('"').Replace("\\\"", "\""));

                if ((bool)generateResponse.success)
                {
                    LogMessage("✓ Generate Report button clicked successfully");
                    UpdateStatus("Report generation started. Waiting for results...", 85);

                    // Wait for the report to generate and load (same as manual)
                    await Task.Delay(5000);

                    // Verify the report has loaded (same as manual)
                    await VerifyReportGenerated();
                }
                else
                {
                    LogMessage($"ERROR clicking Generate Report button: {generateResponse.error}");
                    throw new Exception($"Failed to click Generate Report button: {generateResponse.error}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Report generation failed - {ex.Message}");
                throw;
            }
        }

        private void Window_Closing(object sender, System.ComponentModel.CancelEventArgs e)
        {
            try
            {
                _logWindow?.Close();
                webView?.Dispose();
            }
            catch (Exception ex)
            {
                // Log but don't prevent closing
                System.Diagnostics.Debug.WriteLine($"Error disposing WebView: {ex.Message}");
            }
        }

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

        #region Monthly Report Functions

        private async void BtnMonthlyReport_Click(object sender, RoutedEventArgs e)
        {
            try
            {
                SetButtonsEnabled(false);
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
                    LogMessage("⚠ No monthly data found");
                    MessageBox.Show("No monthly data found for the specified employee.", "No Data", 
                                   MessageBoxButton.OK, MessageBoxImage.Information);
                }
            }
            catch (Exception ex)
            {
                LogMessage($"❌ Monthly report extraction failed: {ex.Message}");
                UpdateStatus("Monthly report failed", 0);
                MessageBox.Show($"Monthly report extraction failed: {ex.Message}", "Error",
                               MessageBoxButton.OK, MessageBoxImage.Error);
            }
            finally
            {
                SetButtonsEnabled(true);
            }
        }

        private async Task<List<MonthlyAttendanceEntry>> ExtractMonthlyAttendanceData()
        {
            var monthlyData = new List<MonthlyAttendanceEntry>();

            try
            {
                // Step 1: Navigate to main page
                LogMessage("1️⃣ Navigating to main page...");
                webView.CoreWebView2.Navigate(MIS_URL);
                await WaitForPageLoad();

                // Step 2: Expand Leave Management tree (same as Today/Yesterday)
                LogMessage("2️⃣ Expanding Leave Management tree...");
                await ExecuteExpandTree();
                await Task.Delay(500); // Brief pause

                // Step 3: Navigate to Attendance Log Report
                LogMessage("3️⃣ Navigating to Attendance Log Report...");
                await NavigateToAttendanceLogReport();

                // Step 4: Select employee
                LogMessage("4️⃣ Selecting employee...");
                await SelectEmployeeInAttendanceReport();

                // Step 5: Set date range (1st of current month to today)
                LogMessage("5️⃣ Setting date range for current month...");
                await SetMonthlyDateRange();

                // Step 6: Generate report
                LogMessage("6️⃣ Generating monthly report...");
                await GenerateAttendanceReport();

                // Step 7: Extract data from report
                LogMessage("7️⃣ Extracting data from monthly report...");
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

                // First verify the Attendance Log Report link is available and get its details (same pattern as swipe log)
                var linkVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        var attendanceLogLink = document.getElementById('TempleteTreeViewt22');
                        if (attendanceLogLink) {
                            return JSON.stringify({
                                found: true,
                                text: attendanceLogLink.textContent || attendanceLogLink.innerText,
                                href: attendanceLogLink.getAttribute('href') || '',
                                hasOnClick: !!attendanceLogLink.getAttribute('onclick'),
                                visible: window.getComputedStyle(attendanceLogLink).display !== 'none'
                            });
                        }
                        return JSON.stringify({found: false, error: 'Attendance Log Report link not found'});
                    })()
                ");

                var linkInfo = JsonConvert.DeserializeObject<dynamic>(
                    linkVerificationResult.Trim('"').Replace("\\\"", "\""));

                if (!(bool)linkInfo.found)
                {
                    LogMessage("ERROR: Attendance Log Report link not found or disappeared");
                    throw new Exception("Attendance Log Report link not found or disappeared");
                }

                LogMessage($"✓ Attendance Log Report link verified: '{linkInfo.text}'");
                LogMessage($"  HasOnClick: {linkInfo.hasOnClick}");

                // Click the Attendance Log Report link (same pattern as swipe log)
                var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var attendanceLogLink = document.getElementById('TempleteTreeViewt22');
                            if (!attendanceLogLink) {
                                return JSON.stringify({success: false, error: 'Link not found'});
                            }

                            // Simulate the click - this should trigger the __doPostBack
                            attendanceLogLink.click();
                            
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
                    clickResult.Trim('"').Replace("\\\"", "\""));

                if ((bool)clickResponse.success)
                {
                    LogMessage("✓ Attendance Log Report link clicked successfully");
                    UpdateStatus("Waiting for Attendance Log Report page to load...", 60);

                    // Wait for the page to navigate/change after the postback (same as swipe log)
                    await Task.Delay(3000);

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

                var targetEmployeeId = txtEmployeeId.Text;
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
                    dropdownResult.Trim('"').Replace("\\\"", "\""));

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

                            var targetId = '{targetEmployeeId}';
                            var found = false;
                            var selectedOption = null;
                            var previousIndex = employeeDropdown.selectedIndex;
                            
                            // Search through all options for ID match (as substring)
                            for (var i = 0; i < employeeDropdown.options.length; i++) {{
                                var option = employeeDropdown.options[i];
                                if (option.text.includes(targetId) || option.value.includes(targetId)) {{
                                    employeeDropdown.selectedIndex = i;
                                    selectedOption = {{
                                        value: option.value,
                                        text: option.text,
                                        index: i
                                    }};
                                    found = true;
                                    break;
                                }}
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
                    selectionResult.Trim('"').Replace("\\\"", "\""));

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
                    verificationResult.Trim('"').Replace("\\\"", "\""));

                if ((bool)verificationInfo.success)
                {
                    LogMessage($"✓ Employee selection verified: '{verificationInfo.currentText}' (Index: {verificationInfo.currentIndex})");
                    
                    // Verify it contains our target ID
                    if (verificationInfo.currentText.ToString().Contains(targetEmployeeId))
                    {
                        LogMessage($"✓ Verified selection contains target employee ID: {targetEmployeeId}");
                    }
                    else
                    {
                        LogMessage($"⚠️ WARNING: Selected employee '{verificationInfo.currentText}' does not contain ID '{targetEmployeeId}'");
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
                    dateInputResult.Trim('"').Replace("\\\"", "\""));

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
                    setDateResult.Trim('"').Replace("\\\"", "\""));

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
                    verificationResult.Trim('"').Replace("\\\"", "\""));

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
            try
            {
                UpdateStatus("Generating attendance report...", 70);
                LogMessage("Starting report generation process...");

                // Step 1: Find and verify the generate button exists
                var buttonResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            // Try multiple common selectors for generate button
                            var generateBtn = document.querySelector('input[name*=""ViewReport""], input[title*=""Generate""], input[value*=""Generate""]') ||
                                            document.querySelector('input[type=""submit""], input[type=""button""][onclick*=""Report""]') ||
                                            document.querySelector('button[onclick*=""Report""], a[onclick*=""Report""]') ||
                                            document.querySelector('input[value*=""View""], input[value*=""Show""]');
                            
                            if (!generateBtn) {
                                // Debug - show available buttons and inputs
                                var allButtons = document.querySelectorAll('input[type=""submit""], input[type=""button""], button');
                                var available = [];
                                for (var i = 0; i < Math.min(allButtons.length, 8); i++) {
                                    available.push({
                                        type: allButtons[i].type || allButtons[i].tagName,
                                        value: allButtons[i].value || 'no-value',
                                        text: allButtons[i].textContent || allButtons[i].innerText || 'no-text',
                                        name: allButtons[i].name || 'no-name',
                                        id: allButtons[i].id || 'no-id',
                                        hasOnClick: !!allButtons[i].getAttribute('onclick')
                                    });
                                }
                                return JSON.stringify({success: false, error: 'Generate button not found', available: available});
                            }

                            return JSON.stringify({
                                success: true,
                                type: generateBtn.type || generateBtn.tagName,
                                value: generateBtn.value || 'no-value',
                                text: generateBtn.textContent || generateBtn.innerText || 'no-text',
                                name: generateBtn.name || 'no-name',
                                id: generateBtn.id || 'no-id',
                                hasOnClick: !!generateBtn.getAttribute('onclick'),
                                disabled: generateBtn.disabled
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: 'Generate button find error'});
                        }
                    })()
                ");

                var buttonInfo = JsonConvert.DeserializeObject<dynamic>(
                    buttonResult.Trim('"').Replace("\\\"", "\""));

                if (!(bool)buttonInfo.success)
                {
                    LogMessage($"ERROR finding generate button: {buttonInfo.error}");
                    if (buttonInfo.available != null)
                    {
                        LogMessage("Available buttons:");
                        foreach (var button in buttonInfo.available)
                        {
                            LogMessage($"  - Type: '{button.type}', Value: '{button.value}', Text: '{button.text}', ID: '{button.id}', HasOnClick: {button.hasOnClick}");
                        }
                    }
                    throw new Exception($"Generate button not found: {buttonInfo.error}");
                }

                LogMessage($"✓ Generate button found: {buttonInfo.type} - '{buttonInfo.value}' (ID: '{buttonInfo.id}')");
                    LogMessage($"  Text: '{buttonInfo.text}', Disabled: {buttonInfo.disabled}, HasOnClick: {buttonInfo.hasOnClick}");                // Step 2: Click the generate button
                var clickResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            var generateBtn = document.querySelector('input[name*=""ViewReport""], input[title*=""Generate""], input[value*=""Generate""]') ||
                                            document.querySelector('input[type=""submit""], input[type=""button""][onclick*=""Report""]') ||
                                            document.querySelector('button[onclick*=""Report""], a[onclick*=""Report""]') ||
                                            document.querySelector('input[value*=""View""], input[value*=""Show""]');
                            
                            if (!generateBtn) {
                                return JSON.stringify({success: false, error: 'Generate button disappeared'});
                            }

                            if (generateBtn.disabled) {
                                return JSON.stringify({success: false, error: 'Generate button is disabled'});
                            }

                            // Click the generate button
                            generateBtn.click();

                            return JSON.stringify({
                                success: true,
                                buttonClicked: generateBtn.value || generateBtn.textContent || generateBtn.innerText,
                                message: 'Generate button clicked successfully'
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: 'Generate button click error'});
                        }
                    })()
                ");

                var clickInfo = JsonConvert.DeserializeObject<dynamic>(
                    clickResult.Trim('"').Replace("\\\"", "\""));

                if (!(bool)clickInfo.success)
                {
                    LogMessage($"ERROR clicking generate button: {clickInfo.error}");
                    throw new Exception($"Generate button click failed: {clickInfo.error}");
                }

                LogMessage($"✓ Generate button clicked successfully: '{clickInfo.buttonClicked}'");
                UpdateStatus("Waiting for report to generate...", 80);
                
                // Step 3: Wait for the report to load and verify it appeared
                await Task.Delay(5000); // Wait for report generation

                var reportVerificationResult = await webView.CoreWebView2.ExecuteScriptAsync(@"
                    (function() {
                        try {
                            // Look for common report indicators
                            var reportTable = document.querySelector('#ReportViewer1 table, table[id*=""report""], .report table, table.ReportTable') ||
                                            document.querySelector('table:has(td:contains(""Employee"")), table:has(th:contains(""Employee""))') ||
                                            document.querySelector('div[id*=""Report""], div[class*=""report""]');
                            
                            var reportContent = document.querySelector('#ReportViewer1, div[id*=""Report""], iframe[src*=""Report""]');
                            
                            // Check for error messages
                            var errorMsg = document.querySelector('.error, .Error, div:contains(""No Data""), div:contains(""Error"")');
                            
                            // Check page title or headers
                            var pageTitle = document.title || '';
                            var headers = document.querySelectorAll('h1, h2, h3');
                            var headerTexts = [];
                            for (var i = 0; i < Math.min(headers.length, 3); i++) {
                                headerTexts.push(headers[i].textContent || headers[i].innerText);
                            }

                            return JSON.stringify({
                                success: true,
                                reportTableFound: !!reportTable,
                                reportContentFound: !!reportContent,
                                errorMessageFound: !!errorMsg,
                                pageTitle: pageTitle,
                                headerTexts: headerTexts,
                                tableCount: document.querySelectorAll('table').length,
                                reportIndicators: {
                                    reportViewer: !!document.getElementById('ReportViewer1'),
                                    reportDiv: !!document.querySelector('div[id*=""Report""]'),
                                    reportClass: !!document.querySelector('.report, .Report')
                                }
                            });
                        } catch (ex) {
                            return JSON.stringify({success: false, error: 'Report verification error'});
                        }
                    })()
                ");

                var verificationInfo = JsonConvert.DeserializeObject<dynamic>(
                    reportVerificationResult.Trim('"').Replace("\\\"", "\""));

                if ((bool)verificationInfo.success)
                {
                    LogMessage($"✓ Report generation verification completed:");
                    LogMessage($"  Page title: '{verificationInfo.pageTitle}'");
                    LogMessage($"  Report table found: {verificationInfo.reportTableFound}");
                    LogMessage($"  Report content found: {verificationInfo.reportContentFound}");
                    LogMessage($"  Error message found: {verificationInfo.errorMessageFound}");
                    LogMessage($"  Total tables on page: {verificationInfo.tableCount}");
                    
                    if (verificationInfo.headerTexts != null && ((Newtonsoft.Json.Linq.JArray)verificationInfo.headerTexts).Count > 0)
                    {
                        LogMessage("  Page headers:");
                        foreach (var header in verificationInfo.headerTexts)
                        {
                            LogMessage($"    - {header}");
                        }
                    }

                    if ((bool)verificationInfo.errorMessageFound)
                    {
                        LogMessage("⚠️ WARNING: Error message detected on page - report may not have generated properly");
                    }
                    else if ((bool)verificationInfo.reportTableFound || (bool)verificationInfo.reportContentFound)
                    {
                        LogMessage("✓ Report appears to have generated successfully");
                    }
                    else
                    {
                        LogMessage("⚠️ WARNING: No clear report content found - may need to wait longer or check different selectors");
                    }
                }
                else
                {
                    LogMessage($"ERROR verifying report generation: {verificationInfo.error}");
                    // Don't throw here as this is just verification - the report might still have generated
                }

                LogMessage("✓ Report generation process completed");
            }
            catch (Exception ex)
            {
                LogMessage($"ERROR: Report generation failed - {ex.Message}");
                throw;
            }
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

                        string raw = await webView.CoreWebView2.ExecuteScriptAsync(script);
                        if (string.IsNullOrWhiteSpace(raw) || raw == "null")
                        {
                                LogMessage("⚠ Extraction script returned null/empty (single-pass)");
                                return monthlyData; // empty
                        }
                        // ExecuteScriptAsync returns JSON text of the returned JS value directly (we returned an object not a string)
                        dynamic? parsed = null;
                        try { parsed = Newtonsoft.Json.JsonConvert.DeserializeObject<dynamic>(raw); }
                        catch (Exception ex)
                        {
                            LogMessage($"⚠ JSON parse failure (single-pass direct object): {ex.Message}");
                            LogMessage($"RAW: {raw.Substring(0, Math.Min(raw.Length, 400))}");
                            return monthlyData;
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
                        else
                        {
                                var err = parsed?.error != null ? parsed.error.ToString() : "Unknown failure";
                                LogMessage($"⚠ Monthly extraction failed (single-pass) - {err}");
                        }

                        if (monthlyData.Count == 0)
                        {
                                LogMessage("⚠ No monthly data extracted (single-pass heuristic)");
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
            try
            {
                btnFullReport.IsEnabled = false;
                LogMessage("Full Report button triggered: running complete automation chain...");
                await RunFullCalendarAutomation();
                ShowFullCalendarWindow();
            }
            catch (Exception ex)
            {
                LogMessage($"Full Report failed: {ex.Message}");
            }
            finally
            {
                btnFullReport.IsEnabled = true;
            }
        }

        private async Task RunFullCalendarAutomation(bool showWindowOnCompletion = false)
        {
            // Sequence: Monthly -> Today -> Yesterday -> Calendar
            try
            {
                LogMessage("[CALENDAR AUTO] Starting monthly extraction...");
                var monthly = await ExtractMonthlyAttendanceData();
                UpdateMonthlyReportCache(monthly);
                LogMessage($"[CALENDAR AUTO] Monthly cached: {monthly.Count} records");

                LogMessage("[CALENDAR AUTO] Extracting Today swipe log (lightweight) ...");
                var todayCalc = await QuickDailyHoursExtraction(isToday:true);
                UpdateTodayCalculations(todayCalc);
                LogMessage($"[CALENDAR AUTO] Today calc entries: {todayCalc.Count}");

                LogMessage("[CALENDAR AUTO] Extracting Yesterday swipe log (lightweight) ...");
                var yestCalc = await QuickDailyHoursExtraction(isToday:false);
                UpdateYesterdayCalculations(yestCalc);
                LogMessage($"[CALENDAR AUTO] Yesterday calc entries: {yestCalc.Count}");

                if (showWindowOnCompletion)
                    ShowFullCalendarWindow();
            }
            catch (Exception ex)
            {
                LogMessage($"[CALENDAR AUTO] Failed: {ex.Message}");
            }
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

        private async Task<List<WorkHoursCalculation>> QuickDailyHoursExtraction(bool isToday)
        {
            var dayType = isToday ? "Today" : "Yesterday";
            LogMessage($"[DAILY] Starting {dayType} extraction...");
            var list = new List<WorkHoursCalculation>();
            try
            {
                LogMessage($"[DAILY] Navigating to MIS for {dayType}...");
                webView.CoreWebView2.Navigate(MIS_URL);
                await WaitForPageLoad();
                
                LogMessage($"[DAILY] Expanding tree for {dayType}...");
                await ExecuteExpandTree();
                
                LogMessage($"[DAILY] Clicking swipe log for {dayType}...");
                await ExecuteClickSwipeLog();
                
                LogMessage($"[DAILY] Setting report type for {dayType} (index: {(isToday ? 0 : 1)})...");
                cmbReportType.SelectedIndex = isToday ? 0 : 1;
                
                LogMessage($"[DAILY] Generating report for {dayType}...");
                await ExecuteGenerateReport();
                
                LogMessage($"[DAILY] Extracting data for {dayType}...");
                var entries = await ExecuteDataExtraction();
                LogMessage($"[DAILY] Raw entries extracted for {dayType}: {entries.Count}");
                
                if (entries.Any())
                {
                    // Use the same working logic as ComparisonWindow
                    var labeledEntries = LabelSwipeEntries(entries);
                    var workHours = CalculateWorkHoursFromLabels(labeledEntries);
                    
                    LogMessage($"[DAILY] Labeled entries for {dayType}: {labeledEntries.Count}");
                    LogMessage($"[DAILY] Calculated work hours for {dayType}: {workHours}");
                    
                    var grouped = entries.GroupBy(e => e.Date);
                    foreach (var g in grouped)
                    {
                        var targetDate = DateTime.TryParse(g.Key, out var dt) ? dt : (isToday ? DateTime.Today : DateTime.Today.AddDays(-1));
                        
                        list.Add(new WorkHoursCalculation
                        {
                            Date = targetDate,
                            WorkingHoursDisplay = workHours,
                            TotalWorkingHoursDisplay = workHours,
                            Status = workHours != "0:00" && workHours != "Error" ? "Present" : "No Data"
                        });
                        
                        LogMessage($"[DAILY] Added calculation for {dayType}: Date={targetDate:yyyy-MM-dd}, Hours={workHours}");
                    }
                }
                else
                {
                    LogMessage($"[DAILY] No entries found for {dayType}");
                }
            }
            catch (Exception ex)
            {
                LogMessage($"[DAILY] {dayType} extraction failed: {ex.Message}");
                LogMessage($"[DAILY] Stack trace: {ex.StackTrace}");
            }
            
            LogMessage($"[DAILY] {dayType} extraction completed with {list.Count} calculation entries");
            return list;
        }

        private static TimeSpan CalcHoursSpan(IEnumerable<string> times)
        {
            var parsed = new List<TimeSpan>();
            foreach (var t in times)
            {
                // Handle both "HH:mm:ss" and "HH:mm:ss AM/PM" formats
                if (TimeSpan.TryParse(t, out var ts))
                {
                    parsed.Add(ts);
                }
                else if (DateTime.TryParse(t, out var dt))
                {
                    // If it's in AM/PM format, parse as DateTime and extract TimeOfDay
                    parsed.Add(dt.TimeOfDay);
                }
            }
            
            if (parsed.Count < 2) return TimeSpan.Zero;
            var start = parsed.Min(); 
            var end = parsed.Max();
            
            // Handle case where work spans midnight (end < start)
            if (end < start)
            {
                end = end.Add(TimeSpan.FromDays(1));
            }
            
            return end - start;
        }

        private static string FormatSpan(TimeSpan span)
        {
            if (span <= TimeSpan.Zero) return string.Empty;
            return $"{(int)span.TotalHours:00}:{span.Minutes:00}";
        }

        // Copy of working logic from ComparisonWindow for chunk-based calculation
        private string IdentifyGateType(string machineName)
        {
            if (string.IsNullOrEmpty(machineName))
                return "Unknown";

            // Based on CT2 building structure and VBA gate classification
            // Main gates - building/campus entry points (like Basement, Parking)
            var mainGatePatterns = new[] { "Basement", "Parking", "Tripod", "Ground", "Reception", "Main Gate", "Security", "Entry Gate" };
            
            // Work gates - actual office/work floors (4th Floor, 5th Floor, etc.)
            var workGatePatterns = new[] { "4th Floor", "5th Floor", "6th Floor", "7th Floor", "8th Floor", "9th Floor", "Floor Le", "Floor Ri", "Building.*Floor", "Office Floor" };
            
            // Play gates - recreation areas (cafeteria, lobby, etc.)
            var playGatePatterns = new[] { "Cafeteria", "Recreation", "Lobby", "Canteen", "Food Court", "Gaming", "Break" };

            // Check in order of specificity
            foreach (var pattern in workGatePatterns)
            {
                if (machineName.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return "WorkGate";
            }

            foreach (var pattern in playGatePatterns)
            {
                if (machineName.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return "PlayGate";
            }

            foreach (var pattern in mainGatePatterns)
            {
                if (machineName.Contains(pattern, StringComparison.OrdinalIgnoreCase))
                    return "MainGate";
            }

            return "Unknown";
        }

        private DateTime ParseTime(string timeStr)
        {
            if (string.IsNullOrEmpty(timeStr))
                return DateTime.MinValue;

            // Try to parse various time formats
            var formats = new[]
            {
                "hh:mm:ss tt",  // 12-hour format with AM/PM
                "h:mm:ss tt",   // 12-hour format with AM/PM (single digit hour)
                "HH:mm:ss",
                "H:mm:ss", 
                "HH:mm",
                "H:mm",
                "dd/MM/yyyy HH:mm:ss",
                "dd-MM-yyyy HH:mm:ss",
                "MM/dd/yyyy HH:mm:ss",
                "yyyy-MM-dd HH:mm:ss"
            };

            foreach (var format in formats)
            {
                if (DateTime.TryParseExact(timeStr, format, null, System.Globalization.DateTimeStyles.None, out DateTime result))
                {
                    // If only time was parsed, assume today's date
                    if (result.Date == DateTime.MinValue.Date)
                    {
                        result = DateTime.Today.Add(result.TimeOfDay);
                    }
                    return result;
                }
            }

            // Fallback to general parsing
            if (DateTime.TryParse(timeStr, out DateTime generalResult))
            {
                return generalResult;
            }

            return DateTime.MinValue;
        }

        private List<SwipeLogEntry> LabelSwipeEntries(List<SwipeLogEntry> entries)
        {
            if (entries == null || entries.Count == 0)
                return entries ?? new List<SwipeLogEntry>();

            var labeledEntries = new List<SwipeLogEntry>();
            
            // Sort entries by time 
            var sortedEntries = entries.OrderBy(e => ParseTime(e.Time)).ToList();
            
            LogMessage($"[LABEL] Labeling {entries.Count} swipe entries and calculating durations");
            
            for (int i = 0; i < sortedEntries.Count; i++)
            {
                var entry = sortedEntries[i];
                var gateType = IdentifyGateType(entry.MachineName);
                var direction = entry.Direction;
                var activity = "";
                var duration = "";
                
                // Apply VBA labeling logic - based on your clarification:
                // WORK = inside work area (office floors)
                // PLAY = outside work area (basement, parking, reception, etc.)
                if (gateType == "MainGate" && (entry.MachineName.Contains("Basement") || entry.MachineName.Contains("Parking")))
                {
                    // Real exit gates - basement/parking areas are PLAY (outside work area)
                    activity = "PLAY"; // Outside work area
                }
                else if (gateType == "WorkGate" || entry.MachineName.Contains("Floor"))
                {
                    // Floor gates - WORK area (inside work area)
                    activity = "WORK";
                }
                else if (gateType == "PlayGate")
                {
                    // Recreation areas - PLAY (outside work area)
                    activity = "PLAY";
                }
                else if (gateType == "MainGate")
                {
                    // Other main gates (reception, lobby) - PLAY (outside work area)
                    activity = "PLAY";
                }
                else
                {
                    // Unknown areas - default to PLAY (outside work area)
                    activity = "PLAY";
                }
                
                // Calculate duration from this swipe to next swipe
                if (i < sortedEntries.Count - 1)
                {
                    var currentTime = ParseTime(entry.Time);
                    var nextTime = ParseTime(sortedEntries[i + 1].Time);
                    
                    if (nextTime > currentTime)
                    {
                        var timeDiff = nextTime - currentTime;
                        var totalMinutes = (int)timeDiff.TotalMinutes;
                        var hours = totalMinutes / 60;
                        var minutes = totalMinutes % 60;
                        duration = $"{hours}:{minutes:D2}";
                    }
                    else
                    {
                        duration = "0:00";
                    }
                }
                else
                {
                    duration = "-"; // Last entry has no next swipe
                }
                
                // Create labeled entry with duration
                var labeledEntry = new SwipeLogEntry
                {
                    EmployeeId = entry.EmployeeId,
                    Date = entry.Date,
                    MachineName = entry.MachineName,
                    Direction = entry.Direction,
                    Time = entry.Time,
                    Activity = activity,
                    Duration = duration
                };
                
                LogMessage($"[LABEL] {entry.Time} | {direction} | {entry.MachineName} | {gateType} → Activity: {activity} | Duration: {duration}");
                labeledEntries.Add(labeledEntry);
            }
            
            return labeledEntries;
        }

        private string CalculateWorkHoursFromLabels(List<SwipeLogEntry> labeledEntries)
        {
            if (labeledEntries == null || labeledEntries.Count == 0)
                return "0:00";

            try
            {
                var totalWorkMinutes = 0.0;
                
                LogMessage($"[CALC] Calculating work time from {labeledEntries.Count} labeled entries");
                
                // Calculate time chunks between consecutive swipes
                for (int i = 0; i < labeledEntries.Count - 1; i++)
                {
                    var currentEntry = labeledEntries[i];
                    var nextEntry = labeledEntries[i + 1];
                    
                    // The time chunk between current and next swipe has the activity of current EXIT
                    if (currentEntry.Activity == "WORK")
                    {
                        var startTime = ParseTime(currentEntry.Time);
                        var endTime = ParseTime(nextEntry.Time);
                        
                        if (endTime > startTime)
                        {
                            var chunkMinutes = (endTime - startTime).TotalMinutes;
                            totalWorkMinutes += chunkMinutes;
                            
                            LogMessage($"[CALC] WORK chunk: {startTime:HH:mm:ss} to {endTime:HH:mm:ss} = {chunkMinutes:F0} minutes");
                        }
                    }
                }
                
                // Handle ongoing work for today (if last entry indicates work)
                if (labeledEntries.Count > 0)
                {
                    var lastEntry = labeledEntries.Last();
                    if (lastEntry.Activity == "WORK")
                    {
                        var lastTime = ParseTime(lastEntry.Time);
                        var now = DateTime.Now;
                        
                        // Only add ongoing time if it's the same day
                        if (now.Date == lastTime.Date && now > lastTime)
                        {
                            var ongoingMinutes = (now - lastTime).TotalMinutes;
                            totalWorkMinutes += ongoingMinutes;
                            
                            LogMessage($"[CALC] ONGOING work: {lastTime:HH:mm:ss} to {now:HH:mm:ss} = {ongoingMinutes:F0} minutes");
                        }
                    }
                }
                
                var hours = (int)(totalWorkMinutes / 60);
                var minutes = (int)(totalWorkMinutes % 60);
                
                LogMessage($"[CALC] TOTAL WORK TIME: {totalWorkMinutes:F0} minutes = {hours}:{minutes:D2}");
                
                return $"{hours}:{minutes:D2}";
            }
            catch (Exception ex)
            {
                LogMessage($"[CALC] Error calculating work hours: {ex.Message}");
                return "Error";
            }
        }

        #endregion
    }
}