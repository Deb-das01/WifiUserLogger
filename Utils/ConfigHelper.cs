using ClosedXML;

namespace WifiUserLogger.Utils
{
    public static class ConfigHelper
    {
        // Get the value of the specified key from Config.json
        public static string configFilePath = Path.Combine(Directory.GetCurrentDirectory(), "Config.json");
        public static string GetConfigValue(string key)
        {
            // Read the JSON file and parse it to retrieve the value for the specified key
            if (!File.Exists(configFilePath))
            {
                throw new FileNotFoundException($"Config file not found at path: {configFilePath}");
            }

            string jsonContent = File.ReadAllText(configFilePath);
            var jsonObject = System.Text.Json.JsonDocument.Parse(jsonContent);
            if (jsonObject != null)
            {
                if (jsonObject.RootElement.TryGetProperty(key, out var value))
                {
                    return value.GetString() ?? string.Empty;
                }
                else
                {
                    throw new KeyNotFoundException($"Key '{key}' not found in config file.");
                }
            }
            else
            {
                throw new Exception("Failed to parse config file.");
            }
        }
    }
}
