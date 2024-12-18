using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using LaRottaO.OfficeTranslationTool.Services;
using LaRottaO.OfficeTranslationTool.Utils;
using LaRottaO.OfficeTranslationTool.Utils.Utils;
using System.Diagnostics;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool
{
    internal class FormLogic(MainForm mainForm)
    {
        private IProcessOfficeFile _iProcessOfficeFile;

        private readonly ILocalDictionary _iDictionary = new JsonDictionaryService();

        private readonly ITranslation _itranslation = new TranslateUsingDeepLService();

        private readonly MainForm _mainForm = mainForm;

        public void test()
        {
            Debug.WriteLine(_iProcessOfficeFile.getShapesStoredInMemory().shapes);
        }

        public async Task launchSelectFileDialog()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Office & Json Files|*.*;*.*",
                Title = "Select a translation project or Office File to translate"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            if (!String.IsNullOrEmpty(currentOfficeDocPath))
            {
                UIHelpers.offerToSaveDocumentBeforeExiting(_iProcessOfficeFile);
            }

            await openOfficeFile(openFileDialog.FileName);
        }

        public async Task openOfficeFile(String fileName)
        {
            await Task.Run(async () =>
            {
                string extension = Path.GetExtension(fileName);

                if (!File.Exists(fileName))
                {
                    UIHelpers.showErrorMessage($"The associated file {fileName} was not found.");
                    return;
                }

                switch (extension.ToLower())
                {
                    case ".json":

                        //**************************************************************
                        //If this program was used before, there should be an Office
                        //file named file.extension.json next to it,
                        //so let' get rid of the .json part and try to open the file.
                        //**************************************************************

                        String associatedFileName = fileName.Replace(".json", "");

                        await openOfficeFile(associatedFileName);

                        return;

                    case ".docx":
                    case ".doc":

                        _iProcessOfficeFile = new ProcessWordUsingInterop();

                        break;

                    case ".pptx":
                    case ".ppt":

                        _iProcessOfficeFile = new ProcessPowerPointUsingInterop();

                        break;

                    case ".xlsx":
                    case ".xls":
                        //_iProcessOfficeFile = new ProcessExcelFileService();
                        break;

                    default:
                        UIHelpers.showErrorMessage("Unsupported file format");
                        return;
                }

                currentOfficeDocPath = fileName;
                var launchAppResult = _iProcessOfficeFile.launchOfficeProgramInstance();

                if (!launchAppResult.success)
                {
                    UIHelpers.showErrorMessage(launchAppResult.errorReason);
                    return;
                }

                var openFileResult = _iProcessOfficeFile.openOfficeFile();

                if (!openFileResult.success)
                {
                    UIHelpers.showErrorMessage(openFileResult.errorReason);
                    return;
                }

                //**********************************************
                //Flow depends if there's already a JSON
                //***********************************************

                mainForm.panelLoading.InvokeFromAnotherThread(() =>
                {
                    mainForm.panelLoading.Visible = true;
                });

                if (File.Exists(currentOfficeDocPath + ".json"))
                {
                    await loadShapesFromJson();
                }
                else
                {
                    await loadShapesFromOfficeFile();
                }

                mainForm.panelLoading.InvokeFromAnotherThread(() =>
                {
                    mainForm.panelLoading.Visible = false;
                });
            });
        }

        private async Task loadShapesFromJson()
        {
            await Task.Run(async () =>
            {
                Debug.WriteLine("Project .json found, loading shapes from it");

                var loadResult = LoadOfficeDocumentFromJson.load();

                if (!loadResult.success)
                {
                    UIHelpers.showErrorMessage($"{loadResult.errorReason} - creating a new project");
                    await loadShapesFromOfficeFile();
                }

                _iProcessOfficeFile.overwriteShapesStoredInMemory(loadResult.shapes);

                _mainForm.mainDataGridView.InvokeFromAnotherThread(() =>
                {
                    _mainForm.mainDataGridView.DataSource = _iProcessOfficeFile.getShapesStoredInMemory().shapes;
                });
            });
        }

        private async Task loadShapesFromOfficeFile()
        {
            await Task.Run(() =>
            {
                Debug.WriteLine("No project .json found, loading shapes from Office file...");

                var extractionResult = _iProcessOfficeFile.extractShapesFromFile();

                if (!extractionResult.success)
                {
                    UIHelpers.showErrorMessage(extractionResult.errorReason);
                    return;
                }

                _mainForm.mainDataGridView.InvokeFromAnotherThread(() =>
                {
                    _mainForm.mainDataGridView.DataSource = _iProcessOfficeFile.getShapesStoredInMemory().shapes;
                });

                var saveResult = SaveOfficeDocumentAsJson.save(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

                if (!saveResult.success)
                {
                    UIHelpers.showErrorMessage(saveResult.errorReason);
                }
            });
        }

        public void closeOfficeFile()
        {
            if (!String.IsNullOrEmpty(currentOfficeDocPath))
            {
                UIHelpers.offerToSaveDocumentBeforeExiting(_iProcessOfficeFile);
            }
        }

        public (Boolean success, String errorReason) saveNewTranslationTypedByUserOnMainDgv(int row, int col, String newValue)
        {
            var resultGetChangedValue = _iProcessOfficeFile.getShapeFromMemoryAtIndex(row);

            if (!resultGetChangedValue.success)
            {
                return (false, resultGetChangedValue.errorReason);
            }

            var addChangedValueToDic = addTranslationToDictionary(resultGetChangedValue.shape.originalText, resultGetChangedValue.shape.newText, false);

            if (!addChangedValueToDic.success)
            {
                return (false, addChangedValueToDic.errorReason);
            }

            var saveDocumentAsJson = SaveOfficeDocumentAsJson.save(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

            if (!saveDocumentAsJson.success)
            {
                return (false, saveDocumentAsJson.errorReason);
            }

            return (true, resultGetChangedValue.errorReason);
        }

        public (Boolean success, String errorReason) userClickedMainDataGridRow(int row, int col)
        {
            if (replaceInProgress)
            {
                return (false, "Replace in progress");
            }

            var resultGetChangedValue = _iProcessOfficeFile.getShapeFromMemoryAtIndex(row);

            if (!resultGetChangedValue.success)
            {
                return (false, resultGetChangedValue.errorReason);
            }

            var selectShapeOnFile = _iProcessOfficeFile.navigateToShapeOnFile(resultGetChangedValue.shape);

            if (!selectShapeOnFile.success)
            {
                return (false, selectShapeOnFile.errorReason);
            }

            return (true, "");
        }

        public (Boolean success, String errorReason) userClickedDgvPartialExpressionsRow(int row, int col)
        {
            SavedTranslation partialExpression = _iDictionary.getPartialExpressionList().partialExpressions[row];

            _mainForm.textBoxNewPartialExpTerm.Text = partialExpression.term;
            _mainForm.textBoxNewPartialExpTermTrans.Text = partialExpression.translation;

            return (true, "");
        }

        public (Boolean success, String errorReason) deleteEntryFromPartialExpressionDic(string term, String translation)
        {
            var resultDelete = _iDictionary.deleteFromLocalDictionary(term, translation, true);

            if (!resultDelete.success)
            {
                UIHelpers.showErrorMessage(resultDelete.errorReason);
                return (false, resultDelete.errorReason);
            }

            populatePartialExpressionsDgv(_mainForm.dataGridViewPartialExpressions);

            _mainForm.textBoxNewPartialExpTerm.Text = "";

            _mainForm.textBoxNewPartialExpTermTrans.Text = "";

            return resultDelete;
        }

        public void setDictionaryLanguage(String sourceLanguage, String targetLanguage)
        {
            if (!String.IsNullOrEmpty(sourceLanguage))
            {
                selectedSourceLanguage = AVAILABLE_LANGUAGES[sourceLanguage].ToUpper();
            }

            if (!String.IsNullOrEmpty(targetLanguage))
            {
                selectedTargetLanguage = AVAILABLE_LANGUAGES[targetLanguage].ToUpper();
            }

            if (areBothSourceAndDestintionLanguagesSet())
            {
                Debug.WriteLine($"Selected Source Language: {selectedSourceLanguage} Selected Target Language: {selectedTargetLanguage} ");

                _iDictionary.initializeLocalDictionary();
            }
        }

        public Boolean areBothSourceAndDestintionLanguagesSet()
        {
            return (!String.IsNullOrEmpty(selectedSourceLanguage)) && (!String.IsNullOrEmpty(selectedTargetLanguage));
        }

        public (Boolean success, String errorReason) addTranslationToDictionary(String term, String translation, Boolean isPartial)
        {
            var addResult = _iDictionary.addOrUpdateLocalDictionary(term, translation, isPartial);

            if (!addResult.success)
            {
                UIHelpers.showErrorMessage(addResult.errorReason);
                return addResult;
            }

            if (isPartial)
            {
                populatePartialExpressionsDgv(_mainForm.dataGridViewPartialExpressions);
                _mainForm.textBoxNewPartialExpTerm.Text = "";
                _mainForm.textBoxNewPartialExpTermTrans.Text = "";
            }

            return addResult;
        }

        public async Task<(Boolean success, String errorReason)> translateAllShapeElements(String sourceLanguage, String targetLanguage)
        {
            await Task.Run(() =>
            {
                _iDictionary.initializeLocalDictionary();

                //Iterate on all items

                foreach (PptShape shapeUnderTranslation in _iProcessOfficeFile.getShapesStoredInMemory().shapes)
                {
                    //Check if the string is not a number, blank or pure symbols

                    if (string.IsNullOrEmpty(shapeUnderTranslation.originalText))
                    {
                        continue;
                    }

                    if (!shapeUnderTranslation.originalText.Any(char.IsLetter))
                    {
                        continue;
                    }

                    UIHelpers.setCursorOnDataGridRowThreadSafe(_mainForm.mainDataGridView, shapeUnderTranslation.indexOnPresentation, true);

                    //Check if the term exists in local dictionary as a complete word

                    var localDicResult = _iDictionary.getTermFromLocalDictionary(shapeUnderTranslation.originalText);

                    if (!localDicResult.success)
                    {
                        //Local dictionary failed, fatal error, abort process
                        return (false, $"Unable to load local dictionary. {localDicResult.errorReason}");
                    }

                    if (localDicResult.success && localDicResult.termExists)
                    {
                        Debug.WriteLine($"{shapeUnderTranslation.originalText} found on local dictionary!");
                        shapeUnderTranslation.newText = localDicResult.termTranslation;

                        _mainForm.mainDataGridView.InvokeFromAnotherThread(() =>
                        {
                            _mainForm.mainDataGridView.Refresh();
                        });

                        continue;
                    }

                    Debug.WriteLine($"{shapeUnderTranslation.originalText} not found on local dictionary, using API");

                    var apiResult = _itranslation.translate(shapeUnderTranslation.originalText);

                    if (!apiResult.success)
                    {
                        //API failed, continue with the next one
                        UIHelpers.showErrorMessage($"Unable to get translation from API. {apiResult.errorReason}");
                        continue;
                    }

                    Debug.WriteLine($"{shapeUnderTranslation.originalText} found on API");

                    shapeUnderTranslation.newText = apiResult.translatedText;
                    _iDictionary.addOrUpdateLocalDictionary(shapeUnderTranslation.originalText, apiResult.translatedText, false);

                    _mainForm.mainDataGridView.InvokeFromAnotherThread(() =>
                    {
                        _mainForm.mainDataGridView.Refresh();
                    });
                }

                //Checks for partial text and replaces

                var partialReplaceResult = _iDictionary.replacePartialExpressions(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

                if (!partialReplaceResult.success)
                {
                    UIHelpers.showErrorMessage($"Unable to replace partial words.{partialReplaceResult.errorReason}");
                }

                _iProcessOfficeFile.overwriteShapesStoredInMemory(partialReplaceResult.replacedExpressions);

                _mainForm.mainDataGridView.InvokeFromAnotherThread(() =>
                {
                    _mainForm.mainDataGridView.Refresh();
                });

                SaveOfficeDocumentAsJson.save(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

                return (true, "");
            });

            return (true, "");
        }

        public async Task<(Boolean success, String errorReason)> applyChangesOnOfficeFile(Boolean useOriginalText, Boolean useTranslatedText)
        {
            await Task.Run(() =>
            {
                ///////////////////

                _iDictionary.initializeLocalDictionary();

                //Iterate on all items

                foreach (PptShape shapeUnderTranslation in _iProcessOfficeFile.getShapesStoredInMemory().shapes)
                {
                    //Check if the string is not a number, blank or pure symbols

                    if (string.IsNullOrEmpty(shapeUnderTranslation.originalText))
                    {
                        continue;
                    }

                    if (!shapeUnderTranslation.originalText.Any(char.IsLetter))
                    {
                        continue;
                    }

                    UIHelpers.setCursorOnDataGridRowThreadSafe(_mainForm.mainDataGridView, shapeUnderTranslation.indexOnPresentation, true);

                    var replaceResult = _iProcessOfficeFile.replaceShapeText(shapeUnderTranslation, useOriginalText, useTranslatedText, true);

                    /*
                    if (!replaceResult.success)
                    {
                        DialogResult dialogResult = UIHelpers.showYesNoQuestion($"Failed to replace shape text. {replaceResult.errorReason} Do you want to continue?");
                        if (dialogResult == DialogResult.No)
                        {
                            return (false, replaceResult.errorReason);
                        }
                    }
                    */
                }

                return (true, "");

                ///////////////////
            });

            return (true, "");
        }

        public void populatePartialExpressionsDgv(DataGridView dataGridView)
        {
            dataGridView.DataSource = _iDictionary.getPartialExpressionList().partialExpressions;

            dataGridView.Refresh();
        }
    }
}