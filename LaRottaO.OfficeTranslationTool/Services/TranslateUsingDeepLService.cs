using LaRottaO.OfficeTranslationTool.Interfaces;
using Newtonsoft.Json.Linq;
using RestSharp;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Services
{
    internal class TranslateUsingDeepLService : ITranslation
    {
        public (bool success, string errorReason, string translatedText) translate(string term)
        {
            if (String.IsNullOrEmpty(deepLUrl))
            {
                return (false, "The DeepL URL is empty. Please set it first.", "");
            }

            if (String.IsNullOrEmpty(deepLAuthKey))
            {
                return (false, "The DeepL Auth Key is empty. Please set it first.", "");
            }

            var client = new RestClient(deepLUrl);
            var request = new RestRequest
            {
                Method = Method.Post,
                RequestFormat = DataFormat.Json
            };

            request.AddHeader("Authorization", "DeepL-Auth-Key " + deepLAuthKey);
            request.AddHeader("Content-Type", "application/json");

            //DeepL will screw the line jumps. This code fixes it:

            String originalTextWithProperLineJumps = term.Replace("\r", "\r\n");

            var body = new
            {
                text = new[] { originalTextWithProperLineJumps },
                target_lang = selectedTargetLanguage
            };

            request.AddJsonBody(body);

            try
            {
                var response = client.Execute(request);

                if (response.IsSuccessful)
                {
                    var jsonResponse = JObject.Parse(response.Content);
                    var translatedText = jsonResponse["translations"][0]["text"].ToString();
                    return (true, "", translatedText);
                }
                else
                {
                    String errorExplanation = "";

                    if (response.StatusDescription != null)
                    {
                        errorExplanation = response.StatusDescription.ToString();
                    }
                    else if (response.ErrorMessage != null)
                    {
                        errorExplanation = response.ErrorMessage.ToString();
                    }

                    return (false, errorExplanation, "");
                }
            }
            catch (Exception ex)
            {
                return (false, ex.Message, "");
            }
        }
    }
}