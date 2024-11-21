using LaRottaO.OfficeTranslationTool.Models;

namespace LaRottaO.OfficeTranslationTool.Interfaces
{
    internal interface IProcessOfficeFile
    {
        (bool success, string errorReason) launchOfficeProgramInstance();

        (bool success, string errorReason) openOfficeFile();

        bool isOfficeProgramOpen();

        bool isOfficeFileOpen();

        (bool success, string errorReason) extractShapesFromFile();

        (bool success, string errorReason) overwriteShapesStoredInMemory(List<ShapeElement> shapes);

        (bool success, string errorReason, List<ShapeElement> shapes) getShapesStoredInMemory();

        (bool success, string errorReason, ShapeElement shape) getShapeFromMemoryAtIndex(int index);

        (bool success, string errorReason) navigateToShapeOnFile(ShapeElement shapeElement);

        (bool success, string errorReason) replaceShapeText(ShapeElement shapeElement);

        (bool success, string errorReason) saveChangesOnFile();

        (bool success, string errorReason) closeCurrentlyOpenFile(Boolean saveChangesBeforeClosing);

        (bool success, string errorReason) closeOfficeProgramInstance();
    }
}