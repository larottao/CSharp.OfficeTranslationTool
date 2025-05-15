using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using LaRottaO.OfficeTranslationTool.Models;

namespace LaRottaO.OfficeTranslationTool
{
    public static class GlobalConstants
    {
        public enum ElementType
        { TABLE, SHAPE, PARAGRAPH }

        public static readonly Dictionary<string, TransLang> AVAILABLE_TRANS_LANGS = new Dictionary<string, TransLang>
        {
            { "Bulgarian", new TransLang("bg", Microsoft.Office.Core.MsoLanguageID.msoLanguageIDBulgarian)},
            { "Chinese Simplified", new TransLang("zh-CN", Microsoft.Office.Core.MsoLanguageID.msoLanguageIDSimplifiedChinese)},
            { "Chinese Traditional", new TransLang("zh-TW", Microsoft.Office.Core.MsoLanguageID.msoLanguageIDTraditionalChinese)},
            { "English", new TransLang("en", Microsoft.Office.Core.MsoLanguageID.msoLanguageIDEnglishUS)},
            { "Finnish", new TransLang("fi", Microsoft.Office.Core.MsoLanguageID.msoLanguageIDFinnish)},
            { "French", new TransLang("fr",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDFrench)},
            { "German", new TransLang("de",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDGerman)},
            { "Hindi", new TransLang("hi",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDHindi)},
            { "Irish", new TransLang("ga",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDGaelicIreland)},
            { "Norwegian", new TransLang("no",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDNorwegianBokmol)},
            { "Polish", new TransLang("pl",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDPolish)},
            { "Spanish", new TransLang("es",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDMexicanSpanish)},
            { "Swedish", new TransLang("sv",Microsoft.Office.Core.MsoLanguageID.msoLanguageIDSwedish)},
            { "Romanian", new TransLang("ro", Microsoft.Office.Core.MsoLanguageID.msoLanguageIDRomanian)},
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