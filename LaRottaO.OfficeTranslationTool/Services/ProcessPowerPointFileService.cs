using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Runtime.InteropServices;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace LaRottaO.OfficeTranslationTool.Services
{
    internal class ProcessPowerPointFileService : IProcessOfficeFile
    {
        private Application pptApp;
        private Presentation pptPresentation;
        private List<ShapeElement> shapesInPresentation;

        public (bool success, string errorReason) closeCurrentlyOpenFile(bool saveChangesBeforeClosing)
        {
            try
            {
                if (saveChangesBeforeClosing)
                {
                    pptPresentation.Save();
                }

                pptPresentation.Close();

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, "Unable to close file. " + ex.ToString());
            }
        }

        public (bool success, string errorReason) extractShapesFromFile()
        {
            shapesInPresentation = new List<ShapeElement>();

            int indexOnPresentationCounter = 0;

            foreach (Slide slide in pptPresentation.Slides)
            {
                foreach (Microsoft.Office.Interop.PowerPoint.Shape shape in slide.Shapes)
                {
                    try
                    {
                        if (shape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                        {
                            shape.Ungroup();
                        }
                    }
                    catch (Exception ex)
                    {
                        Debug.WriteLine($"DEBUG: Unable to ungroup the shape with Id: {slide.SlideID} on Slide: {slide.SlideNumber}. Reason: {ex.ToString()}");
                    }
                }

                int indexOnSlideCounter = 0;

                foreach (Shape shape in slide.Shapes)
                {
                    if (shape.HasTextFrame == MsoTriState.msoTrue)
                    {
                        if (shape.TextFrame.HasText == MsoTriState.msoTrue)
                        {
                            var textRange = shape.TextFrame.TextRange;

                            ShapeElement newElement = new ShapeElement();

                            newElement.indexOnPresentation = indexOnPresentationCounter;
                            newElement.indexOnSlide = indexOnSlideCounter;
                            newElement.slideNumber = slide.SlideNumber;
                            newElement.info = $"Slide {slide.SlideNumber} Item {shape.Id}";
                            newElement.originalText = textRange.Text.ToString();

                            shapesInPresentation.Add(newElement);
                        }
                    }
                    else if (shape.HasTable == MsoTriState.msoTrue)
                    {
                        Table table = shape.Table;

                        for (int col = 1; col <= table.Columns.Count; col++)
                        {
                            for (int row = 1; row <= table.Rows.Count; row++)
                            {
                                Cell cell = table.Cell(row, col);

                                Microsoft.Office.Interop.PowerPoint.Shape cellShape = cell.Shape;

                                if (cellShape.HasTextFrame == MsoTriState.msoTrue && cellShape.TextFrame.HasText == MsoTriState.msoTrue)
                                {
                                    var textRange = cellShape.TextFrame.TextRange;

                                    ShapeElement newElement = new ShapeElement();

                                    newElement.belongsToATable = true;
                                    newElement.parentTableRow = row;
                                    newElement.parentTableColumn = col;

                                    newElement.indexOnPresentation = indexOnPresentationCounter;
                                    newElement.indexOnSlide = indexOnSlideCounter;
                                    newElement.slideNumber = slide.SlideNumber;
                                    newElement.info = $"Slide {slide.SlideNumber} Item {shape.Id}";
                                    newElement.originalText = textRange.Text.ToString();

                                    shapesInPresentation.Add(newElement);
                                }
                            }
                        }
                    }

                    indexOnSlideCounter++;
                } //End foreach (Shape shape in slide.Shapes)

                indexOnPresentationCounter++;
            } //End foreach (Slide slide in pptPresentation.Slides)

            return (true, "");
        }

        public (bool success, string errorReason, List<ShapeElement> shapes) getShapesStoredInMemory()
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || shapesInPresentation == null || shapesInPresentation.Count == 0)
            {
                return (false, "No shapes to show", new List<ShapeElement>());
            }

            return (true, "", shapesInPresentation);
        }

        public bool isOfficeProgramOpen()
        {
            return (pptApp != null);
        }

        public (bool success, string errorReason) launchOfficeProgramInstance()
        {
            try
            {
                pptApp = new Application();
                Debug.WriteLine("Powerpoint launched OK");
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to open PowerPoint application. {ex.ToString()}");
            }
        }

        public (bool success, string errorReason) navigateToShapeOnFile(ShapeElement shapeElement)
        {
            try
            {
                Slide slide = pptPresentation.Slides[shapeElement.slideNumber];
                Shape shape = slide.Shapes[shapeElement.indexOnSlide + 1];

                slide.Select();
                shape.Select();

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        public (bool success, string errorReason) openOfficeFile()
        {
            try
            {
                pptPresentation = pptApp.Presentations.Open(currentOfficeDocPath);
                Debug.WriteLine(currentOfficeDocPath + " loaded. Total slides on pptx: " + pptPresentation.Slides.Count);

                if (pptPresentation.Slides.Count == 0)
                {
                    return (false, "Presentation has zero slides.");
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, "Unable to open presentation. " + ex.ToString());
            }
        }

        public (bool success, string errorReason) replaceShapeText(ShapeElement shapeElement, Boolean useOriginalText, Boolean useTranslatedText, Boolean shrinkIfNecessary)
        {
            try
            {
                Slide slide = pptPresentation.Slides[shapeElement.slideNumber];
                Shape shape = slide.Shapes[shapeElement.indexOnSlide + 1];

                slide.Select();
                shape.Select();

                if (useTranslatedText)
                {
                    shape.TextFrame.TextRange.Text = shapeElement.newText;
                }

                if (useOriginalText)
                {
                    shape.TextFrame.TextRange.Text = shapeElement.originalText;
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        public (bool success, string errorReason) saveChangesOnFile()
        {
            throw new NotImplementedException();
        }

        public (bool success, string errorReason) closeOfficeProgramInstance()
        {
            try
            {
                pptApp.Quit();
                Marshal.ReleaseComObject(pptApp);
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to close power point. {ex.ToString()}");
            }
        }

        public bool isOfficeFileOpen()
        {
            return (pptPresentation != null);
        }

        public (bool success, string errorReason, ShapeElement shape) getShapeFromMemoryAtIndex(int index)
        {
            //TODO return empty shape

            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || shapesInPresentation == null || shapesInPresentation.Count == 0)
            {
                return (false, "No shapes to show", null);
            }

            return (true, "", shapesInPresentation[index]);
        }

        public (bool success, string errorReason) overwriteShapesStoredInMemory(List<ShapeElement> shapes)
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen())
            {
                return (false, "No shapes to overwrite");
            }

            shapesInPresentation = shapes;

            return (true, "");
        }
    }
}