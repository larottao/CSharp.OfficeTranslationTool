using LaRottaO.OfficeTranslationTool.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace LaRottaO.OfficeTranslationTool.Interfaces
{
    //ETBT 'Element to be translated' is a more generic way for calling Shapes, Tables, Paragrpaphs, etc.
    internal interface IProcessOfficeFile
    {
        (bool success, string errorReason) launchOfficeProgramInstance();

        (bool success, string errorReason) openOfficeFile();

        bool isOfficeProgramOpen();

        bool isOfficeFileOpen();

        (bool success, string errorReason) extractETBTsFromFile();

        (bool success, string errorReason) overwriteETBTsStoredInMemory(List<ElementToBeTranslated> elementToBeTranslated);

        (bool success, string errorReason, List<ElementToBeTranslated> shapes) getETBTsStoredInMemory();

        (bool success, string errorReason, ElementToBeTranslated shape) getETBTFromMemoryAtIndex(int index);

        (bool success, string errorReason, Shape? shape) navigateToETBTOnFile(ElementToBeTranslated elementToBeTranslated);

        (bool success, string errorReason) replaceETBTText(ElementToBeTranslated elementToBeTranslated, Boolean useOriginalText, Boolean useTranslatedText, Boolean shrinkIfNecessary);

        (bool success, string errorReason) saveChangesOnFile();

        (bool success, string errorReason) closeCurrentlyOpenFile(Boolean saveChangesBeforeClosing);

        (bool success, string errorReason) closeOfficeProgramInstance();
    }
}