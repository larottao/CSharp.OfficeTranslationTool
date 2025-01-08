using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Word.Application;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;
using System.Reflection.Metadata;
using Document = Microsoft.Office.Interop.Word.Document;
using Range = Microsoft.Office.Interop.Word.Range;
using Microsoft.Office.Core;
using Shape = Microsoft.Office.Interop.Word.Shape;
using Microsoft.Office.Interop.PowerPoint;
using Table = Microsoft.Office.Interop.Word.Table;
using Cell = Microsoft.Office.Interop.Word.Cell;
using Row = Microsoft.Office.Interop.Word.Row;

namespace LaRottaO.OfficeTranslationTool.Services
{
    public class ProcessWordUsingInterop : IProcessOfficeFile
    {
        private Application wordApp;
        private Document wordDocument;
        private List<ElementToBeTranslated> elementsInDocument;

        public (bool success, string errorReason) closeCurrentlyOpenFile(bool saveChangesBeforeClosing)
        {
            try
            {
                if (saveChangesBeforeClosing)
                {
                    wordDocument.Save();
                }

                if (wordDocument != null)
                {
                    wordDocument.Close();
                }

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
                wordApp.Quit();
                Marshal.ReleaseComObject(wordApp);
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to close Word. {ex.ToString()}");
            }
        }

        public (bool success, string errorReason) extractETBTsFromFile()
        {
            try
            {
                elementsInDocument = new List<ElementToBeTranslated>();

                // Iterate through sections
                foreach (Section section in wordDocument.Sections)
                {
                    // Iterate through paragraphs
                    foreach (Paragraph paragraph in section.Range.Paragraphs)
                    {
                        string text = paragraph.Range.Text.Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            elementsInDocument.Add(new ElementToBeTranslated
                            {
                                internalId = paragraph.ParaID,
                                originalText = paragraph.Range.Text.TrimEnd('\r', '\a').Trim(),
                                type = GlobalConstants.ElementType.PARAGRAPH
                            });
                        }
                    }

                    // Iterate through shapes (e.g., text boxes)
                    foreach (Shape shape in section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                    {
                        if (shape.TextFrame.HasText == 1)
                        {
                            string shapeText = shape.TextFrame.TextRange.Text.Trim();
                            if (!string.IsNullOrEmpty(shapeText))
                            {
                                elementsInDocument.Add(new ElementToBeTranslated
                                {
                                    internalId = shape.ID,
                                    originalText = shapeText,
                                    type = GlobalConstants.ElementType.SHAPE
                                });
                                Console.WriteLine("Shape: " + shapeText);
                            }
                        }
                    }
                }

                // Iterate through tables
                foreach (Table table in wordDocument.Tables)
                {
                    foreach (Row row in table.Rows)
                    {
                        foreach (Cell cell in row.Cells)
                        {
                            string cellText = cell.Range.Text.Trim();
                            if (!string.IsNullOrEmpty(cellText))
                            {
                                elementsInDocument.Add(new ElementToBeTranslated
                                {
                                    internalId = cell.ID,
                                    originalText = cellText,
                                    type = GlobalConstants.ElementType.TABLE
                                });
                            }
                        }
                    }
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        public (bool success, string errorReason, List<ElementToBeTranslated> shapes) getETBTsStoredInMemory()
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || elementsInDocument == null || elementsInDocument.Count == 0)
            {
                return (false, "No shapes to show", new List<ElementToBeTranslated>());
            }

            return (true, "", elementsInDocument);
        }

        public (bool success, string errorReason, ElementToBeTranslated shape) getETBTFromMemoryAtIndex(int index)
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || elementsInDocument == null || elementsInDocument.Count == 0)
            {
                return (false, "No shapes to show", default(ElementToBeTranslated));
            }

            return (true, "", elementsInDocument[index]);
        }

        public bool isOfficeFileOpen()
        {
            return (wordDocument != null);
        }

        public bool isOfficeProgramOpen()
        {
            return (wordApp != null);
        }

        public (bool success, string errorReason) launchOfficeProgramInstance()
        {
            try
            {
                wordApp = new Application();
                wordApp.Visible = true;
                Debug.WriteLine("Word launched OK");
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to open Word application. {ex.ToString()}");
            }
        }

        public (bool success, string errorReason) openOfficeFile()
        {
            try
            {
                wordDocument = wordApp.Documents.Open(currentOfficeDocPath);
                Debug.WriteLine(currentOfficeDocPath + " loaded.");
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, "Unable to open document. " + ex.ToString());
            }
        }

        public (bool success, string errorReason) overwriteETBTsStoredInMemory(List<ElementToBeTranslated> shapes)
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen())
            {
                return (false, "No shapes to overwrite");
            }

            elementsInDocument = shapes;
            return (true, "");
        }

        public (bool success, string errorReason) replaceETBTText(ElementToBeTranslated shape, bool useOriginalText, bool useTranslatedText, bool shrinkIfNecessary)
        {
            try
            {
                switch (shape.type)
                {
                    case GlobalConstants.ElementType.PARAGRAPH:
                        foreach (Section section in wordDocument.Sections)
                        {
                            foreach (Paragraph paragraph in section.Range.Paragraphs)
                            {
                                if (paragraph.ParaID == shape.internalId)
                                {
                                    string originalText = paragraph.Range.Text.TrimEnd('\r', '\a').Trim();

                                    if (originalText.Contains(shape.originalText))
                                    {
                                        // Preserve leading and trailing whitespace or special characters
                                        string fullText = paragraph.Range.Text;
                                        string leadingText = fullText.Substring(0, fullText.IndexOf(shape.originalText));
                                        string trailingText = fullText.Substring(fullText.IndexOf(shape.originalText) + shape.originalText.Length);

                                        // Replace the text
                                        paragraph.Range.Text = leadingText + shape.newText + trailingText;

                                        return (true, string.Empty);
                                    }
                                    else
                                    {
                                        return (false, $"Text mismatch for paragraph with ID {shape.internalId}.");
                                    }
                                }
                            }
                        }
                        return (false, $"Paragraph with ID {shape.internalId} not found.");

                    case GlobalConstants.ElementType.SHAPE:
                        foreach (Section section in wordDocument.Sections)
                        {
                            foreach (Shape wordShape in section.Headers[WdHeaderFooterIndex.wdHeaderFooterPrimary].Shapes)
                            {
                                if (wordShape.ID == shape.internalId && wordShape.TextFrame.HasText == -1)
                                {
                                    string fullText = wordShape.TextFrame.TextRange.Text;
                                    string leadingText = fullText.Substring(0, fullText.IndexOf(shape.originalText));
                                    string trailingText = fullText.Substring(fullText.IndexOf(shape.originalText) + shape.originalText.Length);

                                    wordShape.TextFrame.TextRange.Text = leadingText + shape.newText + trailingText;
                                    return (true, string.Empty);
                                }
                            }
                        }
                        return (false, $"Shape with ID {shape.internalId} not found.");

                    case GlobalConstants.ElementType.TABLE:
                        foreach (Table table in wordDocument.Tables)
                        {
                            if (table.ID == shape.internalId)
                            {
                                foreach (Row row in table.Rows)
                                {
                                    foreach (Cell cell in row.Cells)
                                    {
                                        if (cell.Range.Text.Contains(shape.originalText))
                                        {
                                            string fullText = cell.Range.Text;
                                            string leadingText = fullText.Substring(0, fullText.IndexOf(shape.originalText));
                                            string trailingText = fullText.Substring(fullText.IndexOf(shape.originalText) + shape.originalText.Length);

                                            cell.Range.Text = leadingText + shape.newText + trailingText;
                                            return (true, string.Empty);
                                        }
                                    }
                                }
                                return (false, $"Text not found in table with ID {shape.internalId}.");
                            }
                        }
                        return (false, $"Table with ID {shape.internalId} not found.");

                    default:
                        return (false, "Unsupported shape type.");
                }
            }
            catch (Exception ex)
            {
                return (false, $"Error replacing shape text: {ex.Message}");
            }
        }

        public (bool success, string errorReason) saveChangesOnFile()
        {
            try
            {
                wordDocument.Save();
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to save document: {ex.ToString()}");
            }
        }

        public (bool success, string errorReason, Microsoft.Office.Interop.PowerPoint.Shape? shape) navigateToETBTOnFile(ElementToBeTranslated shape)
        {
            return (false, "", null);
        }

        public (bool success, string errorReason) replaceETBTText(ElementToBeTranslated elementToBeTranslated, bool useOriginalText, bool useTranslatedText)
        {
            throw new NotImplementedException();
        }

        public (bool success, string errorReason) replaceAllETBTsText(List<ElementToBeTranslated> elementsToBeTranslated, bool useOriginalText, bool useTranslatedText)
        {
            throw new NotImplementedException();
        }
    }
}