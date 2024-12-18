using LaRottaO.OfficeTranslationTool.Models;
using Newtonsoft.Json;
using System.Diagnostics;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Utils.Utils
{
    internal static class SaveOfficeDocumentAsJson
    {
        public static (bool success, string errorReason) save(List<ElementToBeTranslated> shapesList)
        {
            try
            {
                var settings = new JsonSerializerSettings
                {
                    PreserveReferencesHandling = PreserveReferencesHandling.Objects,
                    Formatting = Newtonsoft.Json.Formatting.Indented // Optional: for pretty printing
                };
                string json = JsonConvert.SerializeObject(shapesList, settings);
                File.WriteAllText(currentOfficeDocPath + ".json", json);

                Debug.WriteLine(DateTime.Now + " changes on document saved on disk.");

                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, $"Unable to save project on external json {ex.ToString()}");
            }
        }
    }
}