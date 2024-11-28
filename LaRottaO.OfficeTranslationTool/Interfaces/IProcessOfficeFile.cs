using LaRottaO.OfficeTranslationTool.Models;
using Microsoft.Office.Interop.PowerPoint;

namespace LaRottaO.OfficeTranslationTool.Interfaces
{
    internal interface IProcessOfficeFile
    {
        (bool success, string errorReason) launchOfficeProgramInstance();

        (bool success, string errorReason) openOfficeFile();

        bool isOfficeProgramOpen();

        bool isOfficeFileOpen();

        (bool success, string errorReason) extractShapesFromFile();

        (bool success, string errorReason) overwriteShapesStoredInMemory(List<PptShape> shapes);

        (bool success, string errorReason, List<PptShape> shapes) getShapesStoredInMemory();

        (bool success, string errorReason, PptShape shape) getShapeFromMemoryAtIndex(int index);

        (bool success, string errorReason, Shape? shape) navigateToShapeOnFile(PptShape shape);

        (bool success, string errorReason) replaceShapeText(PptShape pptShape, Boolean useOriginalText, Boolean useTranslatedText, Boolean shrinkIfNecessary);

        (bool success, string errorReason) saveChangesOnFile();

        (bool success, string errorReason) closeCurrentlyOpenFile(Boolean saveChangesBeforeClosing);

        (bool success, string errorReason) closeOfficeProgramInstance();
    }
}