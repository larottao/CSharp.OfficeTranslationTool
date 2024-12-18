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
        private List<PptShape> elementsInDocument;

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

        public (bool success, string errorReason) extractShapesFromFile()
        {
            try
            {
                elementsInDocument = new List<PptShape>();

                // Iterate through sections
                foreach (Section section in wordDocument.Sections)
                {
                    // Iterate through paragraphs
                    foreach (Paragraph paragraph in section.Range.Paragraphs)
                    {
                        string text = paragraph.Range.Text.Trim();
                        if (!string.IsNullOrEmpty(text))
                        {
                            elementsInDocument.Add(new PptShape
                            {
                                section = section.Index,
                                internalId = paragraph.ParaID,
                                originalText = text,
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
                                elementsInDocument.Add(new PptShape
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
                                elementsInDocument.Add(new PptShape
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

        public (bool success, string errorReason, List<PptShape> shapes) getShapesStoredInMemory()
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || elementsInDocument == null || elementsInDocument.Count == 0)
            {
                return (false, "No shapes to show", new List<PptShape>());
            }

            return (true, "", elementsInDocument);
        }

        public (bool success, string errorReason, PptShape shape) getShapeFromMemoryAtIndex(int index)
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen() || elementsInDocument == null || elementsInDocument.Count == 0)
            {
                return (false, "No shapes to show", default(PptShape));
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

        public (bool success, string errorReason) overwriteShapesStoredInMemory(List<PptShape> shapes)
        {
            if (!isOfficeProgramOpen() || !isOfficeFileOpen())
            {
                return (false, "No shapes to overwrite");
            }

            elementsInDocument = shapes;
            return (true, "");
        }

        public (bool success, string errorReason) replaceShapeText(PptShape shape, Boolean useOriginalText, Boolean useTranslatedText, Boolean shrinkIfNecessary)
        {
            try
            {
                if (shape.type == GlobalConstants.ElementType.PARAGRAPH)
                {
                    foreach (Section theSection in wordDocument.Sections)
                    {
                        Debug.WriteLine(theSection.Index);

                        foreach (Paragraph theParagraph in theSection.Range.Paragraphs)
                        {
                            if (theParagraph.ParaID == shape.internalId)
                            {
                                Debug.WriteLine($"found:{theParagraph.ParaID}");
                                theParagraph.Range.Text = shape.newText;
                                break;
                            }
                        }
                    }
                }

                return (true, $"");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to replace shape text: {ex.ToString()}");
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

        public (bool success, string errorReason, Microsoft.Office.Interop.PowerPoint.Shape? shape) navigateToShapeOnFile(PptShape shape)
        {
            return (false, "", null);
        }
    }
}