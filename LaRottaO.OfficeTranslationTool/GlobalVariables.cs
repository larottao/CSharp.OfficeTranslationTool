namespace LaRottaO.OfficeTranslationTool
{
    internal class GlobalVariables
    {
        public static String currentOfficeDocPath { get; set; } = "";
        public static String selectedSourceLanguage { get; set; } = "";
        public static String selectedTargetLanguage { get; set; } = "";
        public static Boolean replaceInProgress { get; set; } = false;
        public static String deepLUrl { get; set; } = "https://api-free.deepl.com/v2/translate";

        //TODO REMOVE AND LOAD FROM FILE
        public static String deepLAuthKey { get; set; } = "e9c2c043-2be4-4465-94b0-cdaa26941cab:fx";

        public static Dictionary<string, string> AVAILABLE_LANGUAGES { get; } = new Dictionary<string, string>
        {
            { "Bulgarian", "bg" },
            { "Chinese Simplified", "zh-CN" },
            { "Chinese Traditional", "zh-TW" },
            { "English", "en" },
            { "Finnish", "fi" },
            { "French", "fr" },
            { "German", "de" },
            { "Hindi", "hi" },
            { "Irish", "ga" },
            { "Norwegian", "no" },
            { "Polish", "pl" },
            { "Spanish", "es" },
            { "Swedish", "sv" },
            { "Romanian", "ro" },
        };
    }
}