using LaRottaO.OfficeTranslationTool.Interfaces;
using Newtonsoft.Json.Linq;
using RestSharp;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Services
{
    internal class TranslateUsingDeepLService : ITranslation
    {
        public (bool success, string errorReason) init()
        {
            return (true, "");
        }

        public (bool success, string errorReason) terminate()
        {
            return (true, "");
        }

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

            // Preserve special characters like bullets, tabs, and line breaks
            string encodedText = term
                .Replace("\r", "\r\n")   // Normalize line breaks
                .Replace("\t", "\\t")   // Escape tabs explicitly if needed
                .Replace("\u2022", "\\u2022"); // Explicitly preserve bullets (optional)

            var body = new
            {
                text = new[] { encodedText },
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

                    // Post-process translated text to ensure special characters are respected
                    translatedText = translatedText
                        .Replace("\\t", "\t") // Decode tabs
                        .Replace("\\u2022", "\u2022"); // Decode bullets (if encoded previously)

                    return (true, "", translatedText);
                }
                else
                {
                    string errorExplanation = response.StatusDescription ?? response.ErrorMessage ?? "Unknown error occurred.";
                    return (false, errorExplanation, "");
                }
            }
            catch (Exception ex)
            {
                return (false, ex.Message, "");
            }
        }

        public (bool success, string errorReason, string translatedText) translateOriginal(string term)
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