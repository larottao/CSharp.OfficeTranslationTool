﻿using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using LaRottaO.OfficeTranslationTool.Utils;
using Microsoft.Office.Core;
using Microsoft.Office.Interop.PowerPoint;
using System.Diagnostics;
using System.Runtime.InteropServices;
using System.Text;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;
using Application = Microsoft.Office.Interop.PowerPoint.Application;
using Shape = Microsoft.Office.Interop.PowerPoint.Shape;

namespace LaRottaO.OfficeTranslationTool.Services
{
    internal class ProcessPowerPointUsingInterop : IProcessOfficeFile
    {
        private static Application pptApp;
        private static Presentation pptPresentation;
        private static List<ElementToBeTranslated> shapesInPresentation;

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

        public (bool success, string errorReason) extractETBTsFromFile()
        {
            try
            {
                shapesInPresentation = new List<ElementToBeTranslated>();
                int indexOnPresentationCounter = 0;

                foreach (Slide slide in pptPresentation.Slides)
                {
                    // Process each shape on the slide
                    int indexOnSlideCounter = 0;

                    foreach (Shape shape in slide.Shapes)
                    {
                        ProcessShapeRecursively(shape, slide.SlideNumber, indexOnPresentationCounter, indexOnSlideCounter);
                        indexOnSlideCounter++;
                    }

                    indexOnPresentationCounter++;
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        private void ProcessShapeRecursively(Shape shape, int slideNumber, int indexOnPresentation, int indexOnSlide)
        {
            try
            {
                // Handle grouped shapes recursively
                if (shape.Type == MsoShapeType.msoGroup)
                {
                    foreach (Shape groupedShape in shape.GroupItems)
                    {
                        ProcessShapeRecursively(groupedShape, slideNumber, indexOnPresentation, indexOnSlide);
                    }
                }
                else if (shape.HasTextFrame == MsoTriState.msoTrue && shape.TextFrame.HasText == MsoTriState.msoTrue)
                {
                    // Process standard shapes with text
                    var textRange = shape.TextFrame.TextRange;

                    ElementToBeTranslated newElement = new ElementToBeTranslated
                    {
                        internalId = shape.Id,
                        type = GlobalConstants.ElementType.SHAPE,
                        indexOnPresentation = indexOnPresentation,
                        slideNumber = slideNumber,
                        info = $"Slide {slideNumber} Shape {shape.Id}",
                        originalText = textRange.Text
                    };

                    shapesInPresentation.Add(newElement);
                }
                else if (shape.HasTable == MsoTriState.msoTrue)
                {
                    // Process tables and their cells
                    Table table = shape.Table;

                    for (int row = 1; row <= table.Rows.Count; row++)
                    {
                        for (int col = 1; col <= table.Columns.Count; col++)
                        {
                            Cell cell = table.Cell(row, col);
                            Shape cellShape = cell.Shape;

                            if (cellShape.HasTextFrame == MsoTriState.msoTrue && cellShape.TextFrame.HasText == MsoTriState.msoTrue)
                            {
                                var textRange = cellShape.TextFrame.TextRange;

                                ElementToBeTranslated newElement = new ElementToBeTranslated
                                {
                                    internalId = shape.Id, // Table ID for grouping
                                    type = GlobalConstants.ElementType.TABLE,
                                    parentTableRow = row,
                                    parentTableColumn = col,
                                    indexOnPresentation = indexOnPresentation,
                                    slideNumber = slideNumber,
                                    info = $"Slide {slideNumber} Table {shape.Id} Row {row} Col {col}",
                                    originalText = textRange.Text
                                };

                                shapesInPresentation.Add(newElement);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Debug.WriteLine($"Error processing shape on Slide {slideNumber}, Shape ID {shape.Id}: {ex.Message}");
            }
        }

        public (bool success, string errorReason, ElementToBeTranslated shape) getETBTFromMemoryAtIndex(int index)
        {
            //TODO return empty shape

            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || shapesInPresentation == null || shapesInPresentation.Count == 0)
            {
                return (false, "No shapes to show", null);
            }

            return (true, "", shapesInPresentation[index]);
        }

        public (bool success, string errorReason, List<ElementToBeTranslated> shapes) getETBTsStoredInMemory()
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || shapesInPresentation == null || shapesInPresentation.Count == 0)
            {
                return (false, "No shapes to show", new List<ElementToBeTranslated>());
            }

            return (true, "", shapesInPresentation);
        }

        public bool isOfficeFileOpen()
        {
            return (pptPresentation != null);
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

        public (bool success, string errorReason, Shape? shape) navigateToETBTOnFile(ElementToBeTranslated shapeElement)
        {
            //TODO This seems like a waste of resources but ensures there are no errors that break the code

            Slide slide;

            try
            {
                slide = pptPresentation.Slides[shapeElement.slideNumber];
                slide.Select();
            }
            catch (Exception ex)
            {
                return (false, $"Unable to select slide before selecting shape: {ex.ToString()}", null);
            }

            foreach (Shape possiblyGroupedShape in slide.Shapes)
            {
                try
                {
                    if (possiblyGroupedShape.Type == Microsoft.Office.Core.MsoShapeType.msoGroup)
                    {
                        possiblyGroupedShape.Ungroup();
                    }
                }
                catch (Exception ex)
                {
                    return (false, $"Unable to ungroup shapes before selecting: {ex.ToString()}", null);
                }
            }

            if (slide.Shapes == null || slide.Shapes.Count == 0)
            {
                return (false, $"Unable to select shape, slide contains no shapes", null);
            }

            foreach (Shape ungroupedShape in slide.Shapes)
            {
                if (shapeElement.internalId.Equals(ungroupedShape.Id))
                {
                    try
                    {
                        ungroupedShape.Select();
                        return (true, "", ungroupedShape);
                    }
                    catch (Exception ex)
                    {
                        return (false, $"Unable to select shape: {ex.ToString()}", null);
                    }
                }
            }

            return (false, $"Unable to select shape. It was not found on the slide", null);
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

        public (bool success, string errorReason) overwriteETBTsStoredInMemory(List<ElementToBeTranslated> shapes)
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen())
            {
                return (false, "No shapes to overwrite");
            }

            shapesInPresentation = shapes;

            return (true, "");
        }

        public (bool success, string errorReason) replaceETBTText(ElementToBeTranslated pptShape, Boolean useOriginalText, Boolean useTranslatedText)
        {
            try
            {
                var navResult = navigateToETBTOnFile(pptShape);

                //*****************************************************************************
                //If it fails it cannot continue, because the shape to replace would be null.
                //******************************************************************************

                if (!navResult.success)
                {
                    return (false, navResult.errorReason);
                }

                pptShape.originalPptxShape = navResult.shape;

                var savePropsResult = saveOriginalShapeProperties(pptShape);

                if (!savePropsResult.success)
                {
                    return (false, savePropsResult.errorReason);
                }

                var replaceResult = replaceShapeInnerText(pptShape, useOriginalText, useTranslatedText);

                if (!replaceResult.success)
                {
                    return (false, replaceResult.errorReason);
                }

                var resizeResult = restoreOriginalPropertiesIfChanged(pptShape);

                if (!resizeResult.success)
                {
                    return (false, resizeResult.errorReason);
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        private static (bool success, string errorReason, OriginalShapeProperties props) saveOriginalShapeProperties(ElementToBeTranslated pptShape)
        {
            OriginalShapeProperties props = new OriginalShapeProperties();

            try
            {
                //Pending to Delete
                return (true, "", props);
            }
            catch (Exception ex)
            {
                return (false, $"Unable to save shape properties before text change. {ex.ToString()}", props);
            }
        }

        private static (bool success, string errorReason) restoreOriginalPropertiesIfChanged(ElementToBeTranslated pptShape)
        {
            try
            {
                if (pptShape.type.Equals(GlobalConstants.ElementType.TABLE))
                {
                    return (true, "");
                }

                float originalFontSize = pptShape.originalPptxShape.TextFrame.TextRange.Font.Size;

                Microsoft.Office.Interop.PowerPoint.TextFrame textFrame = pptShape.originalPptxShape.TextFrame;

                textFrame.AutoSize = PpAutoSize.ppAutoSizeNone;
                textFrame.MarginLeft = 0;
                textFrame.MarginRight = 0;
                textFrame.MarginTop = 0;
                textFrame.MarginBottom = 0;

                // Initial font size
                float fontSize = textFrame.TextRange.Font.Size;
                float minFontSize = originalFontSize - 2.0f; // Define a minimum font size to avoid overly shrinking text
                const float scaleFactor = 0.95f; // Reduce font size by 5% in each iteration

                while (!IsTextInsideShape(pptShape.originalPptxShape) && fontSize > minFontSize)
                {
                    fontSize *= scaleFactor;
                    textFrame.TextRange.Font.Size = fontSize;
                }

                if (fontSize <= minFontSize)
                {
                    Console.WriteLine("Text could not be resized to fit within the shape.");
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to resize text: {ex.ToString()}");
            }
        }

        private static bool IsTextInsideShape(Shape shape)
        {
            var textFrame = shape.TextFrame;
            var textRange = textFrame.TextRange;

            // Get text dimensions
            float textWidth = textRange.BoundWidth;
            float textHeight = textRange.BoundHeight;

            // Get shape dimensions
            float shapeWidth = shape.Width;
            float shapeHeight = shape.Height;

            // Check shape type
            switch (shape.AutoShapeType)
            {
                case MsoAutoShapeType.msoShapeRectangle:
                    // Text must fit within the rectangle
                    return textWidth <= shapeWidth && textHeight <= shapeHeight;

                case MsoAutoShapeType.msoShapeOval:
                    // For a circle/ellipse, ensure the text fits within the bounding ellipse
                    float radiusX = shapeWidth / 2;
                    float radiusY = shapeHeight / 2;
                    return (textWidth / 2 <= radiusX) && (textHeight / 2 <= radiusY);

                case MsoAutoShapeType.msoShapeIsoscelesTriangle:
                    // For a triangle, the text must fit within the triangular bounds
                    // Approximation: Ensure the text's bounding box fits within the triangle's height and width
                    return textWidth <= shapeWidth && textHeight <= shapeHeight;

                default:
                    // Default behavior for unsupported shapes (treat as rectangle)
                    return textWidth <= shapeWidth && textHeight <= shapeHeight;
            }
        }

        private static (bool success, string errorReason) replaceShapeInnerText(ElementToBeTranslated shape, Boolean useOriginalText, Boolean useTranslatedText)
        {
            try
            {
                // Determine the replacement text
                string replacementText = useTranslatedText ? shape.newText : shape.originalText;

                // Handle based on shape type
                switch (shape.type)
                {
                    case GlobalConstants.ElementType.SHAPE:
                        foreach (Slide slide in pptPresentation.Slides)
                        {
                            if (slide.SlideNumber != shape.slideNumber) continue;

                            foreach (Shape pptShape in slide.Shapes)
                            {
                                if (pptShape.Id == shape.internalId && pptShape.HasTextFrame == MsoTriState.msoTrue && pptShape.TextFrame.HasText == MsoTriState.msoTrue)
                                {
                                    // Replace text while preserving formatting
                                    var textRange = pptShape.TextFrame.TextRange;
                                    string fullText = textRange.Text;
                                    int startIndex = fullText.IndexOf(shape.originalText, StringComparison.Ordinal);

                                    if (startIndex >= 0)
                                    {
                                        textRange.Replace(shape.originalText, replacementText);
                                        return (true, "");
                                    }
                                    else
                                    {
                                        return (false, $"Text not found in shape on Slide {shape.slideNumber}, Shape ID {shape.internalId}");
                                    }
                                }
                            }
                        }
                        break;

                    case GlobalConstants.ElementType.TABLE:
                        foreach (Slide slide in pptPresentation.Slides)
                        {
                            if (slide.SlideNumber != shape.slideNumber) continue;

                            foreach (Shape pptShape in slide.Shapes)
                            {
                                if (pptShape.Id == shape.internalId && pptShape.HasTable == MsoTriState.msoTrue)
                                {
                                    Table table = pptShape.Table;

                                    if (shape.parentTableRow <= table.Rows.Count && shape.parentTableColumn <= table.Columns.Count)
                                    {
                                        Cell cell = table.Cell(shape.parentTableRow, shape.parentTableColumn);
                                        Shape cellShape = cell.Shape;

                                        if (cellShape.HasTextFrame == MsoTriState.msoTrue && cellShape.TextFrame.HasText == MsoTriState.msoTrue)
                                        {
                                            var textRange = cellShape.TextFrame.TextRange;
                                            string fullText = textRange.Text;
                                            int startIndex = fullText.IndexOf(shape.originalText, StringComparison.Ordinal);

                                            if (startIndex >= 0)
                                            {
                                                textRange.Replace(shape.originalText, replacementText);
                                                return (true, "");
                                            }
                                            else
                                            {
                                                return (false, $"Text not found in table cell on Slide {shape.slideNumber}, Table {shape.internalId}, Row {shape.parentTableRow}, Column {shape.parentTableColumn}");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        return (false, $"Row or column out of bounds for table on Slide {shape.slideNumber}, Table {shape.internalId}");
                                    }
                                }
                            }
                        }
                        break;

                    default:
                        return (false, $"Unsupported shape type: {shape.type}");
                }

                return (false, $"Shape not found on Slide {shape.slideNumber}, ID {shape.internalId}");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to replace shape inner text: {ex.ToString()}");
            }
        }

        public (bool success, string errorReason) saveChangesOnFile()
        {
            //TODO
            throw new NotImplementedException();
        }

        public (bool success, string errorReason) replaceAllETBTsText(List<ElementToBeTranslated> elementsToBeTranslated, bool useOriginalText, bool useTranslatedText)
        {
            throw new NotImplementedException();
        }
    }
}

/*
public (bool success, string errorReason) extractShapesFromFileOriginalImp()
{
    try
    {
        shapesInPresentation = new List<PptShape>();

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
                    continue;
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

                        PptShape newElement = new PptShape();

                        newElement.internalId = shape.Id;

                        newElement.belongsToATable = false;
                        newElement.indexOnPresentation = indexOnPresentationCounter;
                        newElement.indexOnSlide = indexOnSlideCounter;
                        newElement.slideNumber = slide.SlideNumber;
                        newElement.info = $"Slide {slide.SlideNumber} Text {shape.Id}";
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

                                PptShape newElement = new PptShape();

                                newElement.internalId = shape.Id;

                                newElement.belongsToATable = true;
                                newElement.parentTableRow = row;
                                newElement.parentTableColumn = col;

                                newElement.indexOnPresentation = indexOnPresentationCounter;
                                newElement.indexOnSlide = indexOnSlideCounter;
                                newElement.slideNumber = slide.SlideNumber;
                                newElement.info = $"Slide {slide.SlideNumber} Table {shape.Id} {row},{col}";
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
    catch (Exception ex)
    {
        return (false, ex.ToString());
    }
}
private static (bool success, string errorReason) replaceShapeInnerTextOriginalImp(PptShape shape, Boolean useOriginalText, Boolean useTranslatedText)
{
    try
    {
        if (shape.belongsToATable)
        {
            Shape tableShape = shape.originalShape;
            Table table = tableShape.Table;

            Cell cell = table.Cell(shape.parentTableRow, shape.parentTableColumn);

            Shape cellShape = cell.Shape;

            if (useTranslatedText)
            {
                cellShape.TextFrame.TextRange.Text = shape.newText;
            }

            if (useOriginalText)
            {
                cellShape.TextFrame.TextRange.Text = shape.originalText;
            }
        }
        else
        {
            if (useTranslatedText)
            {
                shape.originalShape.TextFrame.TextRange.Text = shape.newText;
            }

            if (useOriginalText)
            {
                shape.originalShape.TextFrame.TextRange.Text = shape.originalText;
            }
        }

        return (true, "");
    }
    catch (Exception ex)
    {
        return (false, $"Unable to replace shape inner text {ex.ToString()}");
    }
}

*/