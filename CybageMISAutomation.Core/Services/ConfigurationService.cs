using System.Text.Json;
using CybageMISAutomation.Models;

namespace CybageMISAutomation.Services
{
    /// <summary>
    /// Loads and saves user settings.
    ///
    /// Settings live in %LocalAppData%\CybageMISAutomation\config.json because the app is installed
    /// per-machine under Program Files, which standard users cannot write to. The config.json shipped
    /// next to the executable is only used as the initial defaults on first run.
    ///
    /// A config file that cannot be parsed is never silently overwritten: it is backed up and the
    /// problem is exposed through <see cref="LastLoadError"/> so the UI can tell the user.
    /// </summary>
    public static class ConfigurationService
    {
        private const string ConfigFileName = "config.json";

        public static string ConfigDirectory { get; } = Path.Combine(
            Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData), "CybageMISAutomation");

        public static string ConfigFilePath { get; } = Path.Combine(ConfigDirectory, ConfigFileName);

        /// <summary>Defaults shipped with the application (read-only after install).</summary>
        public static string DefaultConfigFilePath { get; } = Path.Combine(AppContext.BaseDirectory, ConfigFileName);

        /// <summary>Set when the last load hit a problem (e.g. corrupt file); null when it loaded cleanly.</summary>
        public static string? LastLoadError { get; private set; }

        private static readonly SemaphoreSlim _saveLock = new(1, 1);

        private static readonly JsonSerializerOptions _serializerOptions = new()
        {
            PropertyNameCaseInsensitive = true,
            WriteIndented = true,
            PropertyNamingPolicy = JsonNamingPolicy.CamelCase
        };

        public static AppConfig CurrentConfig { get; private set; } = new();

        public static async Task<AppConfig> LoadConfigurationAsync()
        {
            LastLoadError = null;

            if (File.Exists(ConfigFilePath))
            {
                var loaded = await TryReadAsync(ConfigFilePath);
                if (loaded.Config != null)
                {
                    CurrentConfig = loaded.Config;
                    return CurrentConfig;
                }

                // Corrupt user config: keep a copy so nothing is lost, then fall back to defaults.
                var backup = ConfigFilePath + ".corrupt-" + DateTime.Now.ToString("yyyyMMdd-HHmmss");
                try { File.Copy(ConfigFilePath, backup, overwrite: true); }
                catch (Exception copyEx) { backup = $"(backup failed: {copyEx.Message})"; }
                LastLoadError = $"Settings file could not be read ({loaded.Error}). Defaults are being used; the old file was saved as {backup}.";
            }

            // First run (or corrupt user file): start from the shipped defaults when available.
            var defaults = File.Exists(DefaultConfigFilePath) ? await TryReadAsync(DefaultConfigFilePath) : (null, null);
            if (defaults.Error != null)
                LastLoadError ??= $"Default settings file {DefaultConfigFilePath} could not be read ({defaults.Error}).";

            CurrentConfig = defaults.Config ?? new AppConfig();
            try
            {
                await SaveConfigurationAsync(CurrentConfig);
            }
            catch (Exception ex)
            {
                LastLoadError ??= $"Settings could not be saved to {ConfigFilePath}: {ex.Message}";
            }
            return CurrentConfig;
        }

        private static async Task<(AppConfig? Config, string? Error)> TryReadAsync(string path)
        {
            try
            {
                var json = await File.ReadAllTextAsync(path);
                var config = JsonSerializer.Deserialize<AppConfig>(json, _serializerOptions);
                return config == null ? (null, "file is empty") : (config, null);
            }
            catch (Exception ex)
            {
                return (null, ex.Message);
            }
        }

        /// <summary>Writes atomically (temp file + replace). Throws on failure; callers decide how to report it.</summary>
        public static async Task SaveConfigurationAsync(AppConfig config)
        {
            await _saveLock.WaitAsync();
            try
            {
                Directory.CreateDirectory(ConfigDirectory);
                var json = JsonSerializer.Serialize(config, _serializerOptions);
                var temp = ConfigFilePath + ".tmp";
                await File.WriteAllTextAsync(temp, json);
                File.Move(temp, ConfigFilePath, overwrite: true);
                CurrentConfig = config;
            }
            finally
            {
                _saveLock.Release();
            }
        }

        public static AppConfig GetCurrentConfig() => CurrentConfig;

        public static async Task UpdateSettingAsync<T>(string propertyName, T value)
        {
            var property = typeof(AppConfig).GetProperty(propertyName)
                ?? throw new ArgumentException($"Unknown setting '{propertyName}'", nameof(propertyName));
            if (!property.CanWrite)
                throw new ArgumentException($"Setting '{propertyName}' is read-only", nameof(propertyName));
            property.SetValue(CurrentConfig, value);
            await SaveConfigurationAsync(CurrentConfig);
        }

        /// <summary>Kept for API compatibility; configuration is no longer cached between loads.</summary>
        public static void InvalidateCache() { }
    }
}
