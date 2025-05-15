using LaRottaO.OfficeTranslationTool.Models;
using static LaRottaO.OfficeTranslationTool.GlobalConstants;

namespace LaRottaO.OfficeTranslationTool
{
    internal class GlobalVariables
    {
        public static String currentOfficeDocPath { get; set; } = "";
        public static String currentOfficeDocExtension { get; set; } = "";
        public static TransLang? selectedSourceLanguage { get; set; } = null;
        public static TransLang? selectedTargetLanguage { get; set; } = null;
        public static TRANSLATION_METHOD selectedTranslationMethod { get; set; } = TRANSLATION_METHOD.DEEP_L_API;
  
        public static Boolean replaceInProgress { get; set; } = false;
        public static string? jsonDictionaryPath { get; set; } = "";
        public static String deepLUrl { get; set; } = "https://api-free.deepl.com/v2/translate";
        public static string googleTranslateURL { get; set; } = "https://translate.google.com/?sl=SOURCELANG&tl=DESTINATIONLANG&op=translate";
        public static String googleTranslateInputCssSelector { get; set; } = "[aria-label='Source text']";
        public static String googleTranslateCopyButtonCssSelector { get; set; } = "[aria-label='Copy translation']";
        public static string googleTranslateSeleniumProfileName { get; set; } = "Automatizacion";
               
        //Just an example, made up key
        public static String deepLAuthKey { get; set; } = "e9c2c043-2be4-4465-94b0-cdaa26941cab:fx";

 
    }
}