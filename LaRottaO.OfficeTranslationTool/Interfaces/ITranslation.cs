namespace LaRottaO.OfficeTranslationTool.Interfaces
{
    internal interface ITranslation
    {
        (bool success, string errorReason, string translatedText) translate(string term);
    }
}