using System;
using System.IO;
using System.Text.Json;
using System.Threading.Tasks;
using CybageMISAutomationLinux.Models;

namespace CybageMISAutomationLinux.Services
{
    public class ConfigurationService
    {
        private const string CONFIG_FILE_NAME = "config.json";
        private static readonly string ConfigFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, CONFIG_FILE_NAME);
        
        public static AppConfig CurrentConfig { get; private set; } = new AppConfig();
        
        /// <summary>
        /// Load configuration from config.json file. Creates default if file doesn't exist.
        /// </summary>
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
                else
                {
                    // Create default config file
                    CurrentConfig = new AppConfig();
                    await SaveConfigurationAsync(CurrentConfig);
                    return CurrentConfig;
                }
            }
            catch (Exception ex)
            {
                // If loading fails, return default config and log error
                Console.WriteLine($"Failed to load configuration: {ex.Message}. Using defaults.");
                CurrentConfig = new AppConfig();
                return CurrentConfig;
            }
        }
        
        /// <summary>
        /// Save configuration to config.json file
        /// </summary>
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
        
        /// <summary>
        /// Get the current configuration instance
        /// </summary>
        public static AppConfig GetCurrentConfig() => CurrentConfig;
        
        /// <summary>
        /// Update a specific configuration setting and save
        /// </summary>
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