using LaRottaO.OfficeTranslationTool.Interfaces;
using LaRottaO.OfficeTranslationTool.Models;
using System.Diagnostics;
using System.Text.Json;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

//TODO
//Change System.Text.Json to Newtonsoft

namespace LaRottaO.OfficeTranslationTool.Services
{
    internal class JsonDictionaryService : ILocalDictionary
    {
        private string? jsonDictionaryPath;
        private Dictionary<string, SavedTranslation> translationDictionary;

        public (bool success, string errorReason) initializeLocalDictionary()
        {
            jsonDictionaryPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, $"{selectedSourceLanguage}_{selectedTargetLanguage}_saved_translations.json");
            translationDictionary = LoadTranslations();

            Debug.WriteLine($"Dictionary set to: {selectedSourceLanguage} {selectedTargetLanguage}");

            return (true, "");
        }

        public (bool success, string errorReason, bool termExists, string termTranslation) getTermFromLocalDictionary(string term)
        {
            if (translationDictionary.Count == 0)
            {
                translationDictionary = LoadTranslations();
            }

            string key = createKey(selectedSourceLanguage, selectedTargetLanguage, term);

            if (translationDictionary.TryGetValue(key, out SavedTranslation? translation))
            {
                return (true, "", true, translation.translation);
            }

            return (true, "", false, "");
        }

        public (bool success, string errorReason) addOrUpdateLocalDictionary(string term, string translation, bool isPartial)
        {
            try
            {
                if (String.IsNullOrEmpty(term) || String.IsNullOrEmpty(translation))
                {
                    return (false, "The text cannot be empty");
                }

                string key = createKey(selectedSourceLanguage, selectedTargetLanguage, term);

                if (translationDictionary.TryGetValue(key, out var existingTranslation))
                {
                    if (!existingTranslation.term.Equals(term, StringComparison.OrdinalIgnoreCase))
                    {
                        Debug.WriteLine($"Overwriting existing translation for term '{term}'. Old: '{existingTranslation.translation}', New: '{translation}'");
                    }
                }

                SavedTranslation savedTranslation = new SavedTranslation
                {
                    sourceLanguage = selectedSourceLanguage,
                    targetLanguage = selectedTargetLanguage,
                    term = term,
                    translation = translation,
                    isAPartialText = isPartial
                };

                translationDictionary[key] = savedTranslation;
                Debug.WriteLine($"Translation {savedTranslation.term} {savedTranslation.translation} saved for future use.");
                SaveDictionaryAsJson(translationDictionary);

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to save translation on dictionary {ex.ToString()}");
            }
        }

        private static string createKey(String sourceLanguage, String targetLanguage, string originalText)
        {
            return $"{sourceLanguage}|{targetLanguage}|{originalText}";
        }

        private Dictionary<string, SavedTranslation> LoadTranslations()
        {
            if (File.Exists(jsonDictionaryPath))
            {
                var json = File.ReadAllText(jsonDictionaryPath);
                var translations = JsonSerializer.Deserialize<List<SavedTranslation>>(json);
                var dict = new Dictionary<string, SavedTranslation>();

                foreach (var translation in translations)
                {
                    var key = createKey(translation.sourceLanguage, translation.targetLanguage, translation.term);
                    dict[key] = translation;
                }

                return dict;
            }
            return new Dictionary<string, SavedTranslation>();
        }

        private void SaveDictionaryAsJson(Dictionary<string, SavedTranslation> translations)
        {
            var options = new JsonSerializerOptions { WriteIndented = true };
            var json = JsonSerializer.Serialize(translations.Values, options);
            File.WriteAllText(jsonDictionaryPath, json);
        }

        //TODO HORRIBLY INNEFICIENT
        public (bool success, string errorReason, List<ShapeElement> replacedExpressions) replacePartialExpressions(List<ShapeElement> elementsTobeExamined)
        {
            Debug.WriteLine("Replacing partial expressions...");

            foreach (ShapeElement element in elementsTobeExamined)
            {
                if (element.newText != null)

                {
                    foreach (SavedTranslation partialWordTrans in getPartialExpressionList().partialExpressions)
                    {
                        if (element.newText.Contains(partialWordTrans.term))
                        {
                            element.newText = element.newText.Replace(partialWordTrans.term, partialWordTrans.translation);
                            Debug.WriteLine($"Expression {partialWordTrans.term} replaced with {partialWordTrans.translation}");
                        }
                    }
                }
            }

            return (true, "", elementsTobeExamined);
        }

        public (bool success, string errorReason, List<SavedTranslation> partialExpressions) getPartialExpressionList()
        {
            initializeLocalDictionary();

            var partialTextEntries = new List<SavedTranslation>();

            foreach (var entry in translationDictionary.Values)
            {
                if (entry.isAPartialText)
                {
                    partialTextEntries.Add(entry);
                }
            }

            return (true, "", partialTextEntries.OrderBy(element => element.term).ToList());
        }

        public (bool success, string errorReason) deleteFromLocalDictionary(string term, string translation, bool isPartial)
        {
            try
            {
                if (String.IsNullOrEmpty(term) || String.IsNullOrEmpty(translation))
                {
                    return (false, "The term or translation cannot be empty");
                }

                string key = createKey(selectedSourceLanguage, selectedTargetLanguage, term);

                if (translationDictionary.TryGetValue(key, out var foundOnDictionary))
                {
                    // Verify translation and isPartial match before deletion
                    if (foundOnDictionary.term.Equals(term) &&
                        foundOnDictionary.translation.Equals(translation) &&
                        foundOnDictionary.isAPartialText == isPartial)
                    {
                        translationDictionary.Remove(key);
                        Debug.WriteLine($"Translation  '{foundOnDictionary.term} {foundOnDictionary.translation}' has been deleted.");
                        SaveDictionaryAsJson(translationDictionary);
                        return (true, "");
                    }
                    else
                    {
                        return (false, "The provided translation or partial flag does not match the existing record.");
                    }
                }
                else
                {
                    return (false, "The specified term was not found in the dictionary.");
                }
            }
            catch (Exception ex)
            {
                return (false, $"Unable to delete translation from dictionary: {ex.ToString()}");
            }
        }
    }
}