using System.ComponentModel;

namespace LaRottaO.OfficeTranslationTool.Models
{
    public class SavedTranslation
    {
        [Browsable(false)]
        public String sourceLanguage { get; set; }

        [Browsable(false)]
        public String targetLanguage { get; set; }

        public string term { get; set; }
        public string translation { get; set; }

        [Browsable(false)]
        public Boolean isAPartialText { get; set; }
    }
}