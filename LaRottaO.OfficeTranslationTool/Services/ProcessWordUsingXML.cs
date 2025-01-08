using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using Microsoft.Office.Interop.Word;
using System.Diagnostics;
using System.IO.Compression;
using System.Runtime.InteropServices;
using System;
using System.Collections.Generic;
using System.IO;

using System.IO.Compression;

using System.Linq;
using System.Text;
using System.Xml.Linq;

using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Services
{
    public class ProcessWordUsingXML : IProcessOfficeFile
    {
        private List<ElementToBeTranslated> elementsInDocument;

        public (bool success, string errorReason) closeCurrentlyOpenFile(bool saveChangesBeforeClosing)
        {
            throw new NotImplementedException();
        }

        public (bool success, string errorReason) closeOfficeProgramInstance()
        {
            throw new NotImplementedException();
        }

        public (bool success, string errorReason) extractETBTsFromFile()
        {
            try
            {
                elementsInDocument = new List<ElementToBeTranslated>();

                using (var archive = ZipFile.OpenRead(GlobalVariables.currentOfficeDocPath))
                {
                    var documentEntry = archive.GetEntry("word/document.xml");
                    if (documentEntry != null)
                    {
                        using (var stream = documentEntry.Open())
                        {
                            XDocument documentXml = XDocument.Load(stream);
                            ExtractParagraphsFromXml(documentXml, GlobalConstants.ElementType.PARAGRAPH);
                        }
                    }

                    var headers = archive.Entries.Where(e => e.FullName.StartsWith("word/header") && e.FullName.EndsWith(".xml"));
                    foreach (var headerEntry in headers)
                    {
                        using (var stream = headerEntry.Open())
                        {
                            XDocument headerXml = XDocument.Load(stream);
                            ExtractShapesFromXml(headerXml, GlobalConstants.ElementType.SHAPE);
                        }
                    }

                    var tables = archive.Entries.Where(e => e.FullName.StartsWith("word/table") && e.FullName.EndsWith(".xml"));
                    foreach (var tableEntry in tables)
                    {
                        using (var stream = tableEntry.Open())
                        {
                            XDocument tableXml = XDocument.Load(stream);
                            ExtractTablesFromXml(tableXml, GlobalConstants.ElementType.TABLE);
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

        private void ExtractParagraphsFromXml(XDocument xml, GlobalConstants.ElementType elementType)
        {
            var paragraphs = xml.Descendants().Where(e => e.Name.LocalName == "p");
            foreach (var paragraph in paragraphs)
            {
                var text = string.Concat(paragraph.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
                if (!string.IsNullOrEmpty(text))
                {
                    elementsInDocument.Add(new ElementToBeTranslated
                    {
                        internalId = Guid.NewGuid().ToString(), // Generate a unique ID for simplicity
                        originalText = text,
                        type = elementType
                    });
                }
            }
        }

        private void ExtractShapesFromXml(XDocument xml, GlobalConstants.ElementType elementType)
        {
            var shapes = xml.Descendants().Where(e => e.Name.LocalName == "t"); // Assuming text inside shape is similar
            foreach (var shape in shapes)
            {
                var text = shape.Value;
                if (!string.IsNullOrEmpty(text))
                {
                    elementsInDocument.Add(new ElementToBeTranslated
                    {
                        internalId = Guid.NewGuid().ToString(),
                        originalText = text,
                        type = elementType
                    });
                }
            }
        }

        private void ExtractTablesFromXml(XDocument xml, GlobalConstants.ElementType elementType)
        {
            var cells = xml.Descendants().Where(e => e.Name.LocalName == "tc");
            foreach (var cell in cells)
            {
                var text = string.Concat(cell.Descendants().Where(e => e.Name.LocalName == "t").Select(e => e.Value));
                if (!string.IsNullOrEmpty(text))
                {
                    elementsInDocument.Add(new ElementToBeTranslated
                    {
                        internalId = Guid.NewGuid().ToString(),
                        originalText = text,
                        type = elementType
                    });
                }
            }
        }

        public (bool success, string errorReason, ElementToBeTranslated shape) getETBTFromMemoryAtIndex(int index)
        {
            return (true, "", elementsInDocument[index]);
        }

        public (bool success, string errorReason, List<ElementToBeTranslated> shapes) getETBTsStoredInMemory()
        {
            return (true, "", elementsInDocument);
        }

        public bool isOfficeFileOpen()
        {
            throw new NotImplementedException();
        }

        public bool isOfficeProgramOpen()
        {
            return false;
        }

        public (bool success, string errorReason) launchOfficeProgramInstance()
        {
            return (true, "");
        }

        public (bool success, string errorReason, Microsoft.Office.Interop.PowerPoint.Shape? shape) navigateToETBTOnFile(ElementToBeTranslated elementToBeTranslated)
        {
            return (true, "", null);
        }

        public (bool success, string errorReason) openOfficeFile()
        {
            return (true, "");
        }

        public (bool success, string errorReason) overwriteETBTsStoredInMemory(List<ElementToBeTranslated> elementToBeTranslated)
        {
            elementsInDocument = elementToBeTranslated;

            return (true, "");
        }

        public (bool success, string errorReason) saveChangesOnFile()
        {
            throw new NotImplementedException();
        }

        public (bool success, string errorReason) replaceETBTText(ElementToBeTranslated elementToBeTranslated, bool useOriginalText, bool useTranslatedText)
        {
            try
            {
                using (var archive = ZipFile.Open(GlobalVariables.currentOfficeDocPath, ZipArchiveMode.Update))
                {
                    var documentEntry = archive.GetEntry("word/document.xml");
                    if (documentEntry != null)
                    {
                        if (UpdateXml(documentEntry, elementToBeTranslated))
                            return (true, "");
                    }

                    var headers = archive.Entries.Where(e => e.FullName.StartsWith("word/header") && e.FullName.EndsWith(".xml"));
                    foreach (var headerEntry in headers)
                    {
                        if (UpdateXml(headerEntry, elementToBeTranslated))
                            return (true, "");
                    }

                    var tables = archive.Entries.Where(e => e.FullName.StartsWith("word/table") && e.FullName.EndsWith(".xml"));
                    foreach (var tableEntry in tables)
                    {
                        if (UpdateXml(tableEntry, elementToBeTranslated))
                            return (true, "");
                    }
                }

                return (false, "Element not found in document.");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        private bool UpdateXml(ZipArchiveEntry entry, ElementToBeTranslated elementToBeTranslated)
        {
            using (var stream = entry.Open())
            {
                XDocument xml = XDocument.Load(stream);

                var xmlElements = xml.Descendants().Where(e => e.Name.LocalName == "t" && e.Value == elementToBeTranslated.originalText);
                foreach (var xmlElement in xmlElements)
                {
                    xmlElement.Value = elementToBeTranslated.newText;
                    stream.SetLength(0); // Clear the stream
                    using (var writer = new StreamWriter(stream, Encoding.UTF8))
                    {
                        xml.Save(writer);
                    }
                    return true; // Exit after first match to prevent redundant writes
                }
            }
            return false; // No match found
        }

        public (bool success, string errorReason) replaceAllETBTsText(List<ElementToBeTranslated> elementsToBeTranslated, bool useOriginalText, bool useTranslatedText)
        {
            try
            {
                using (var archive = ZipFile.Open(GlobalVariables.currentOfficeDocPath, ZipArchiveMode.Update))
                {
                    var documentEntry = archive.GetEntry("word/document.xml");
                    if (documentEntry != null)
                    {
                        UpdateWholeXml(documentEntry, elementsToBeTranslated);
                    }

                    var headers = archive.Entries.Where(e => e.FullName.StartsWith("word/header") && e.FullName.EndsWith(".xml"));
                    foreach (var headerEntry in headers)
                    {
                        UpdateWholeXml(headerEntry, elementsToBeTranslated);
                    }

                    var tables = archive.Entries.Where(e => e.FullName.StartsWith("word/table") && e.FullName.EndsWith(".xml"));
                    foreach (var tableEntry in tables)
                    {
                        UpdateWholeXml(tableEntry, elementsToBeTranslated);
                    }
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        private void UpdateWholeXml(ZipArchiveEntry entry, List<ElementToBeTranslated> modifiedElements)
        {
            using (var stream = entry.Open())
            {
                XDocument xml = XDocument.Load(stream);

                foreach (var element in modifiedElements)
                {
                    var xmlElements = xml.Descendants().Where(e => e.Name.LocalName == "t" && e.Value == element.originalText);
                    foreach (var xmlElement in xmlElements)
                    {
                        xmlElement.Value = element.newText;
                    }
                }

                stream.SetLength(0); // Clear the stream
                using (var writer = new StreamWriter(stream, Encoding.UTF8))
                {
                    xml.Save(writer);
                }
            }
        }
    }
}