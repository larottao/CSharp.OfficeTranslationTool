using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using LaRottaO.OfficeTranslationTool.Services;
using LaRottaO.OfficeTranslationTool.Utils;
using LaRottaO.OfficeTranslationTool.Utils.Utils;
using System.Diagnostics;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool
{
    internal class FormLogic
    {
        private IProcessOfficeFile _iProcessOfficeFile;

        private ILocalDictionary _iDictionary = new JsonDictionaryService();

        private ITranslation _itranslation = new TranslateUsingDeepLService();

        private DataGridView _dataGridView;

        public FormLogic(DataGridView dataGridView)
        {
            _dataGridView = dataGridView;
        }

        public void test()
        {
            Debug.WriteLine(_iProcessOfficeFile.getShapesStoredInMemory().shapes);
        }

        public void openOfficeFile()
        {
            OpenFileDialog openFileDialog = new OpenFileDialog
            {
                Filter = "Office Files|*.*;*.*",
                Title = "Select an Office File to translate"
            };

            if (openFileDialog.ShowDialog() != DialogResult.OK)
            {
                return;
            }

            if (!String.IsNullOrEmpty(currentOfficeDocPath))
            {
                UIHelpers.offerToSaveDocumentBeforeExiting(_iProcessOfficeFile);
            }

            string extension = Path.GetExtension(openFileDialog.FileName);

            //TODO: Add Word support

            switch (extension.ToLower())
            {
                case ".pptx":
                case ".ppt":
                    _iProcessOfficeFile = new ProcessPowerPointFileService();
                    break;

                case ".xlsx":
                case ".xls":
                    //_iProcessOfficeFile = new ProcessExcelFileService();
                    break;

                default:
                    UIHelpers.showErrorMessage("Unsupported file format");
                    return;
            }

            currentOfficeDocPath = openFileDialog.FileName;
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

            if (File.Exists(currentOfficeDocPath + ".json"))
            {
                loadShapesFromJson();
            }
            else
            {
                loadShapesFromOfficeFile();
            }
        }

        private void loadShapesFromJson()
        {
            Debug.WriteLine("Project .json found, loading shapes from it");

            var loadResult = LoadOfficeDocumentFromJson.load();

            if (!loadResult.success)
            {
                UIHelpers.showErrorMessage($"{loadResult.errorReason} - creating a new project");
                loadShapesFromOfficeFile();
            }

            _iProcessOfficeFile.overwriteShapesStoredInMemory(loadResult.shapes);

            _dataGridView.DataSource = _iProcessOfficeFile.getShapesStoredInMemory().shapes;
        }

        private void loadShapesFromOfficeFile()
        {
            Debug.WriteLine("No project .json found, loading shapes from Office file...");

            var extractionResult = _iProcessOfficeFile.extractShapesFromFile();

            if (!extractionResult.success)
            {
                UIHelpers.showErrorMessage(extractionResult.errorReason);
                return;
            }

            _dataGridView.DataSource = _iProcessOfficeFile.getShapesStoredInMemory().shapes;

            var saveResult = SaveOfficeDocumentAsJson.save(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

            if (!saveResult.success)
            {
                UIHelpers.showErrorMessage(saveResult.errorReason);
            }
        }

        public void closeOfficeFile()
        {
            if (!String.IsNullOrEmpty(currentOfficeDocPath))
            {
                UIHelpers.offerToSaveDocumentBeforeExiting(_iProcessOfficeFile);
            }
        }

        public (Boolean success, String errorReason) saveNewTranslationTypedByUser(int row, int col, String newValue)
        {
            var resultGetChangedValue = _iProcessOfficeFile.getShapeFromMemoryAtIndex(row);

            if (!resultGetChangedValue.success)
            {
                return (false, resultGetChangedValue.errorReason);
            }

            var addChangedValueToDic = addCorrectedWordToDictionary(resultGetChangedValue.shape.originalText, resultGetChangedValue.shape.newText);

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

        public (Boolean success, String errorReason) userClickedRow(int row, int col)
        {
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

        public (Boolean success, String errorReason) addCorrectedWordToDictionary(String term, String translation)
        {
            return _iDictionary.addOrUpdateLocalDictionary(term, translation, false);
        }

        public (Boolean success, String errorReason) translateAllShapeElements(String sourceLanguage, String targetLanguage)
        {
            _iDictionary.initializeLocalDictionary();

            //Iterate on all items

            foreach (ShapeElement shapeUnderTranslation in _iProcessOfficeFile.getShapesStoredInMemory().shapes)
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

                UIHelpers.setCursorOnDataGridRowThreadSafe(_dataGridView, shapeUnderTranslation.indexOnPresentation, true);

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
                    _dataGridView.Refresh();
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
                _dataGridView.Refresh();
            }

            //Checks for partial text and replaces

            var partialReplaceResult = _iDictionary.replacePartialExpressions(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

            if (!partialReplaceResult.success)
            {
                UIHelpers.showErrorMessage($"Unable to replace partial words.{partialReplaceResult.errorReason}");
            }

            _iProcessOfficeFile.overwriteShapesStoredInMemory(partialReplaceResult.replacedExpressions);
            _dataGridView.Refresh();

            SaveOfficeDocumentAsJson.save(_iProcessOfficeFile.getShapesStoredInMemory().shapes);

            return (true, "");
        }

        public (Boolean success, String errorReason) applyChangesOnOfficeFile(Boolean useOriginalText, Boolean useTranslatedText)
        {
            _iDictionary.initializeLocalDictionary();

            //Iterate on all items

            foreach (ShapeElement shapeUnderTranslation in _iProcessOfficeFile.getShapesStoredInMemory().shapes)
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

                UIHelpers.setCursorOnDataGridRowThreadSafe(_dataGridView, shapeUnderTranslation.indexOnPresentation, true);

                _iProcessOfficeFile.navigateToShapeOnFile(shapeUnderTranslation);

                var replaceResult = _iProcessOfficeFile.replaceShapeText(shapeUnderTranslation, useOriginalText, useTranslatedText, true);

                if (!replaceResult.success)
                {
                    DialogResult dialogResult = UIHelpers.showYesNoQuestion($"{replaceResult.errorReason} do you want to continue?");
                    if (dialogResult == DialogResult.No)
                    {
                        return (false, replaceResult.errorReason);
                    }
                }
            }

            return (true, "");
        }
    }
}