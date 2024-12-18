using LaRottaO.OfficeTranslationTool.Models;
using Newtonsoft.Json;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Utils.Utils
{
    internal class LoadOfficeDocumentFromJson
    {
        public static (bool success, string errorReason, List<ElementToBeTranslated> shapes) load()
        {
            if (File.Exists(currentOfficeDocPath + ".json"))
            {
                try
                {
                    string json = File.ReadAllText(currentOfficeDocPath + ".json");

                    List<ElementToBeTranslated>? readShapes = JsonConvert.DeserializeObject<List<ElementToBeTranslated>>(json);

                    if (readShapes == null)
                    {
                        return (false, $"Unable to parse project from external .json, incorrect file structure", new List<ElementToBeTranslated>());
                    }

                    return (true, "", readShapes);
                }
                catch (Exception ex)
                {
                    return (false, $"Unable to read project from external .json {ex}", new List<ElementToBeTranslated>());
                }
            }

            return (false, $"Unable to read project from external .json {currentOfficeDocPath + ".json"}, file does not exist", new List<ElementToBeTranslated>());
        }
    }
}