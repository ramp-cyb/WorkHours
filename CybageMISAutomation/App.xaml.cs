using System.Windows;
using System.IO;
using System;
using System.Diagnostics;
using System.Reflection;

namespace CybageMISAutomation
{
    /// <summary>
    /// Interaction logic for App.xaml
    /// </summary>
    public partial class App : Application
    {
        private static readonly string LogPath = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
            "CybageMISAutomation",
            "startup.log");

        protected override void OnStartup(StartupEventArgs e)
        {
            try
            {
                // Create log directory
                Directory.CreateDirectory(Path.GetDirectoryName(LogPath));
                
                // Log startup attempt
                LogMessage($"=== Application Startup {DateTime.Now} ===");
                LogMessage($"OS: {Environment.OSVersion}");
                LogMessage($".NET Version: {Environment.Version}");
                LogMessage($"Working Directory: {Environment.CurrentDirectory}");
                LogMessage($"Executable Path: {Assembly.GetExecutingAssembly().Location}");
                LogMessage($"Command Line: {string.Join(" ", e.Args)}");
                
                // Set up global exception handling
                AppDomain.CurrentDomain.UnhandledException += OnUnhandledException;
                DispatcherUnhandledException += OnDispatcherUnhandledException;
                
                LogMessage("Calling base.OnStartup...");
                base.OnStartup(e);
                LogMessage("Base startup completed successfully");
            }
            catch (Exception ex)
            {
                LogMessage($"FATAL ERROR in OnStartup: {ex}");
                ShowErrorDialog($"Startup Error: {ex.Message}\n\nCheck log: {LogPath}");
                Shutdown(1);
            }
        }

        private void OnUnhandledException(object sender, UnhandledExceptionEventArgs e)
        {
            LogMessage($"UNHANDLED EXCEPTION: {e.ExceptionObject}");
            if (e.IsTerminating)
            {
                ShowErrorDialog($"Fatal Error: {e.ExceptionObject}\n\nCheck log: {LogPath}");
            }
        }

        private void OnDispatcherUnhandledException(object sender, System.Windows.Threading.DispatcherUnhandledExceptionEventArgs e)
        {
            LogMessage($"DISPATCHER EXCEPTION: {e.Exception}");
            ShowErrorDialog($"UI Error: {e.Exception.Message}\n\nCheck log: {LogPath}");
            e.Handled = true; // Prevent crash
        }

        private static void LogMessage(string message)
        {
            try
            {
                File.AppendAllText(LogPath, $"[{DateTime.Now:yyyy-MM-dd HH:mm:ss.fff}] {message}\n");
            }
            catch
            {
                // Ignore logging errors
            }
        }

        private static void ShowErrorDialog(string message)
        {
            try
            {
                MessageBox.Show(message, "Cybage MIS Automation Error", 
                    MessageBoxButton.OK, MessageBoxImage.Error);
            }
            catch
            {
                // If MessageBox fails, try console
                Console.WriteLine(message);
            }
        }
    }
}