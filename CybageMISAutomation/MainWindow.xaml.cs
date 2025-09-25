using System.Windows;
using System.Windows.Controls;
using Microsoft.Web.WebView2.Core;
using Microsoft.Web.WebView2.Wpf;
using CybageMISAutomation.Models;
using System.Collections.ObjectModel;
using Newtonsoft.Json;
using System.Globalization;

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
        
        // Work hours calculation
        private WorkHoursCalculation _workHoursCalculator = new WorkHoursCalculation();
        private List<SwipeEntry> _todaySwipeEntries = new List<SwipeEntry>();
        private List<SwipeEntry> _yesterdaySwipeEntries = new List<SwipeEntry>();

        public MainWindow()
        {
            InitializeComponent();
            dataGridResults.ItemsSource = _swipeLogData;
            
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
                btnFullAutomation.IsEnabled = true;
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
                                            employeeName: rowData[1] || '',
                                            date: rowData[2] || '',
                                            inTime: rowData[3] || '',
                                            outTime: rowData[4] || '',
                                            duration: rowData[5] || '',
                                            status: rowData[6] || '',
                                            location: rowData[7] || '',
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
                            EmployeeName = entry["employeeName"]?.ToString() ?? "",
                            Date = entry["date"]?.ToString() ?? "",
                            InTime = entry["inTime"]?.ToString() ?? "",
                            OutTime = entry["outTime"]?.ToString() ?? "",
                            Duration = entry["duration"]?.ToString() ?? "",
                            Status = entry["status"]?.ToString() ?? "",
                            Location = entry["location"]?.ToString() ?? ""
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
                btnFullAutomation.IsEnabled = false;

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
                btnFullAutomation.IsEnabled = true;
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
            // Use enhanced extraction for Today's and Yesterday's Swipe Log data
            var extractResult = await webView.CoreWebView2.ExecuteScriptAsync(GetSwipeLogExtractionScript());
            var extractInfo = JsonConvert.DeserializeObject<dynamic>(
                extractResult.Trim('"').Replace("\\\"", "\""));

            var entries = new List<SwipeLogEntry>();
            var swipeEntries = new List<SwipeEntry>();
            
            if ((bool)extractInfo.success)
            {
                foreach (var entry in (Newtonsoft.Json.Linq.JArray)extractInfo.entries)
                {
                    var swipeLogEntry = new SwipeLogEntry
                    {
                        EmployeeId = entry["employeeId"]?.ToString() ?? "",
                        Date = entry["date"]?.ToString() ?? "",
                        Gate = entry["gate"]?.ToString() ?? "",
                        Direction = entry["direction"]?.ToString() ?? "",
                        SwipeTime = entry["swipeTime"]?.ToString() ?? "",
                        // Keep legacy fields for compatibility
                        EmployeeName = entry["employeeName"]?.ToString() ?? "",
                        InTime = entry["inTime"]?.ToString() ?? "",
                        OutTime = entry["outTime"]?.ToString() ?? "",
                        Duration = entry["duration"]?.ToString() ?? "",
                        Status = entry["status"]?.ToString() ?? "",
                        Location = entry["location"]?.ToString() ?? ""
                    };
                    
                    entries.Add(swipeLogEntry);
                    
                    // Also create SwipeEntry for work hours calculation
                    if (!string.IsNullOrEmpty(swipeLogEntry.SwipeTime) && 
                        DateTime.TryParse($"{swipeLogEntry.Date} {swipeLogEntry.SwipeTime}", out DateTime swipeDateTime))
                    {
                        swipeEntries.Add(new SwipeEntry
                        {
                            EmployeeId = swipeLogEntry.EmployeeId,
                            SwipeDateTime = swipeDateTime,
                            Gate = swipeLogEntry.Gate,
                            Direction = swipeLogEntry.Direction,
                            TimeString = swipeLogEntry.SwipeTime,
                            DateString = swipeLogEntry.Date
                        });
                    }
                }
                
                // Calculate work hours if we have swipe entries
                if (swipeEntries.Count > 0)
                {
                    var sessions = _workHoursCalculator.CalculateWorkSessions(swipeEntries);
                    var workHours = _workHoursCalculator.GetTotalWorkHours(sessions);
                    var playHours = _workHoursCalculator.GetTotalPlayHours(sessions);
                    
                    // Add calculated hours to the first entry for display
                    if (entries.Count > 0)
                    {
                        entries[0].CalculatedWorkHours = _workHoursCalculator.FormatDuration(workHours);
                        entries[0].CalculatedPlayHours = _workHoursCalculator.FormatDuration(playHours);
                    }
                    
                    LogMessage($"📊 Work Hours Calculated - Work: {_workHoursCalculator.FormatDuration(workHours)}, Play: {_workHoursCalculator.FormatDuration(playHours)}");
                }
            }
            
            return entries;
        }

        private string GetDataExtractionScript()
        {
            // Legacy script for backward compatibility
            return GetSwipeLogExtractionScript();
        }
        
        private string GetSwipeLogExtractionScript()
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
                                    processedRows: 0,
                                    reportType: 'unknown'
                                }
                            };

                            // Find the ReportViewer1 div
                            var reportViewer = document.querySelector('#ReportViewer1');
                            if (!reportViewer) {
                                result.error = 'ReportViewer1 not found';
                                return JSON.stringify(result);
                            }
                            result.debug.reportViewerFound = true;

                            // Check for swipe log report indicators
                            var pageContent = document.body.textContent || document.body.innerText;
                            var isSwipeLogReport = pageContent.includes('Swipe Log') || 
                                                 pageContent.includes('Entry') || 
                                                 pageContent.includes('Exit') ||
                                                 pageContent.includes('Gate');
                            
                            result.debug.reportType = isSwipeLogReport ? 'swipe_log' : 'attendance_log';

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
                                    hasEmployeeData: false,
                                    hasSwipeData: false
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
                                            // Look for employee ID patterns (numbers)
                                            if (/^\d{3,6}$/.test(text)) {
                                                tableInfo.hasEmployeeData = true;
                                            }
                                            // Look for time patterns
                                            if (/\d{1,2}:\d{2}/.test(text)) {
                                                tableInfo.hasNumericData = true;
                                            }
                                            // Look for swipe-specific patterns
                                            if (/Entry|Exit|Gate/i.test(text)) {
                                                tableInfo.hasSwipeData = true;
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

                            // Extract data from each row based on report type
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
                                        var entry;
                                        
                                        // Parse based on report type and column count
                                        if (isSwipeLogReport && rowData.length >= 5) {
                                            // Today's and Yesterday's Swipe Log format:
                                            // Employee ID, Date, Gate, Direction, Swipe Time
                                            entry = {
                                                employeeId: rowData[0] || '',
                                                employeeName: '',
                                                date: rowData[1] || '',
                                                gate: rowData[2] || '',
                                                direction: rowData[3] || '',
                                                swipeTime: rowData[4] || '',
                                                inTime: '',
                                                outTime: '',
                                                duration: '',
                                                status: '',
                                                location: rowData[2] || '', // Gate as location
                                                rawData: rowData
                                            };
                                        } else {
                                            // Attendance Log Report format (legacy):
                                            // Employee ID, Employee Name, Date, In Time, Out Time, Duration, Status, Location
                                            entry = {
                                                employeeId: rowData[0] || '',
                                                employeeName: rowData[1] || '',
                                                date: rowData[2] || '',
                                                gate: '',
                                                direction: '',
                                                swipeTime: '',
                                                inTime: rowData[3] || '',
                                                outTime: rowData[4] || '',
                                                duration: rowData[5] || '',
                                                status: rowData[6] || '',
                                                location: rowData[7] || '',
                                                rawData: rowData
                                            };
                                        }
                                        
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

        private void ChkManualMode_Changed(object sender, RoutedEventArgs e)
        {
            if (chkManualMode.IsChecked == true)
            {
                pnlManualControls.Visibility = Visibility.Visible;
            }
            else
            {
                pnlManualControls.Visibility = Visibility.Collapsed;
            }
        }
    }
}