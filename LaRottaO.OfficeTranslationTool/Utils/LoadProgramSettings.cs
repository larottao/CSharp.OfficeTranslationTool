using LaRottaO.OfficeTranslationTool.Models;
using System.Text.Json;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Utils
{
    internal static class LoadProgramSettings
    {
        private static ProgramSettings programSettings;

        public static (Boolean success, String errorReason) load()
        {
            String settingsFilePath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "Settings.json");

            if (!File.Exists(settingsFilePath))
            {
                var defaultSettings = new ProgramSettings
                {
                    DeepLUrl = "https://insert-the-deepl-api-url.com",
                    DeepLAuthKey = "insert-the-deepl-auth-key"
                };

                var defaultJson = JsonSerializer.Serialize(defaultSettings);
                File.WriteAllText(settingsFilePath, defaultJson);

                programSettings = defaultSettings;
                deepLUrl = defaultSettings.DeepLUrl;
                deepLAuthKey = defaultSettings.DeepLAuthKey;

                return (true, "Default settings file created.");
            }

            try
            {
                var json = File.ReadAllText(settingsFilePath);
                programSettings = JsonSerializer.Deserialize<ProgramSettings>(json);

                if (programSettings == null)
                {
                    return (false, $"Unable to load program settings. The .json file structure is not valid.");
                }

                deepLUrl = programSettings.DeepLUrl;
                deepLAuthKey = programSettings.DeepLAuthKey;

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to load program settings {ex.ToString()}");
            }
        }
    }
}