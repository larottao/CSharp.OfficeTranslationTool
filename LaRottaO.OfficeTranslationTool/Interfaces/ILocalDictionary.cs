using LaRottaO.OfficeTranslationTool.Models;

namespace LaRottaO.OfficeTranslationTool.Interfaces
{
    internal interface ILocalDictionary
    {
        (bool success, string errorReason) initializeLocalDictionary();

        (bool success, string errorReason) addOrUpdateLocalDictionary(string term, String translation, bool isPartial);

        (bool success, string errorReason, bool termExists, string termTranslation) getTermFromLocalDictionary(string term);

        (bool success, string errorReason, List<ElementToBeTranslated> replacedExpressions) replacePartialExpressions(List<ElementToBeTranslated> elementsTobeExamined);

        (bool success, string errorReason, List<SavedTranslation> partialExpressions) getPartialExpressionList();

        (bool success, string errorReason) deleteFromLocalDictionary(string term, String translation, bool isPartial);
    }
}