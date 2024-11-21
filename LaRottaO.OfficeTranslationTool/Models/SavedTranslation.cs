namespace LaRottaO.OfficeTranslationTool.Models
{
    public class SavedTranslation
    {
        public String sourceLanguage { get; set; }
        public String targetLanguage { get; set; }
        public string term { get; set; }
        public string translation { get; set; }
        public Boolean isAPartialText { get; set; }
    }
}