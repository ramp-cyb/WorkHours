using System.Text.Json;
using CybageMISAutomation.Models;

namespace CybageMISAutomation.Services
{
    public static class ConfigurationService
    {
        private const string ConfigFileName = "config.json";
        private static readonly string ConfigFilePath = Path.Combine(AppContext.BaseDirectory, ConfigFileName);

        public static AppConfig CurrentConfig { get; private set; } = new();

        public static async Task<AppConfig> LoadConfigurationAsync()
        {
            try
            {
                if (File.Exists(ConfigFilePath))
                {
                    var jsonContent = await File.ReadAllTextAsync(ConfigFilePath);
                    var config = JsonSerializer.Deserialize<AppConfig>(jsonContent, new JsonSerializerOptions
                    {
                        PropertyNameCaseInsensitive = true,
                        WriteIndented = true
                    });

                    CurrentConfig = config ?? new AppConfig();
                    return CurrentConfig;
                }

                CurrentConfig = new AppConfig();
                await SaveConfigurationAsync(CurrentConfig);
                return CurrentConfig;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to load configuration: {ex.Message}. Using defaults.");
                CurrentConfig = new AppConfig();
                return CurrentConfig;
            }
        }

        public static async Task SaveConfigurationAsync(AppConfig config)
        {
            try
            {
                var jsonOptions = new JsonSerializerOptions
                {
                    WriteIndented = true,
                    PropertyNamingPolicy = JsonNamingPolicy.CamelCase
                };

                var jsonContent = JsonSerializer.Serialize(config, jsonOptions);
                await File.WriteAllTextAsync(ConfigFilePath, jsonContent);
                CurrentConfig = config;
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to save configuration: {ex.Message}");
                throw;
            }
        }

        public static AppConfig GetCurrentConfig() => CurrentConfig;

        public static async Task UpdateSettingAsync<T>(string propertyName, T value)
        {
            try
            {
                var property = typeof(AppConfig).GetProperty(propertyName);
                if (property != null && property.CanWrite)
                {
                    property.SetValue(CurrentConfig, value);
                    await SaveConfigurationAsync(CurrentConfig);
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine($"Failed to update setting {propertyName}: {ex.Message}");
                throw;
            }
        }
    }
}
