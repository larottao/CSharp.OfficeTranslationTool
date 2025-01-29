namespace LaRottaO.OfficeTranslationTool
{
    internal class GlobalVariables
    {
        public static String currentOfficeDocPath { get; set; } = "";
        public static String currentOfficeDocExtension { get; set; } = "";
        public static String selectedSourceLanguage { get; set; } = "";

        public static TRANSLATION_METHOD selectedTranslationMethod { get; set; } = TRANSLATION_METHOD.DEEP_L_API;
        public static String selectedTargetLanguage { get; set; } = "";
        public static Boolean replaceInProgress { get; set; } = false;
        public static string? jsonDictionaryPath { get; set; } = "";
        public static String deepLUrl { get; set; } = "https://api-free.deepl.com/v2/translate";

        public static string googleTranslateURL { get; set; } = "https://translate.google.com/?sl=SOURCELANG&tl=DESTINATIONLANG&op=translate";
        public static String googleTranslateInputCssSelector { get; set; } = "[aria-label='Source text']";
        public static String googleTranslateCopyButtonCssSelector { get; set; } = "[aria-label='Copy translation']";
        public static string googleTranslateSeleniumProfileName { get; set; } = "Automatizacion";

        //Just an example, made up key
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

        public enum TRANSLATION_METHOD
        { DEEP_L_API, GOOGLE_TRANS_WEB }

        public static Dictionary<string, TRANSLATION_METHOD> AVAILABLE_TRANSLATION_METHODS { get; } = new Dictionary<string, TRANSLATION_METHOD>
        {
            { "Using DeepL API", TRANSLATION_METHOD.DEEP_L_API },
            { "Using Google Translate Web", TRANSLATION_METHOD.GOOGLE_TRANS_WEB }
        };
    }
}