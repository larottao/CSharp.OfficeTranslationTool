using LaRottaO.OfficeTranslationTool.Interfaces;
using Newtonsoft.Json.Linq;

using OpenQA.Selenium;
using OpenQA.Selenium.Firefox;
using OpenQA.Selenium.Interactions;
using OpenQA.Selenium.Support.UI;
using Keys = OpenQA.Selenium.Keys;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;
using System.Diagnostics;
using LaRottaO.OfficeTranslationTool.Utils;

namespace LaRottaO.OfficeTranslationTool.Services
{
    internal class TranslateUsingGoogleTranslate : ITranslation
    {
        private IWebDriver driver;

        private const String ERROR_UNABLE_RETRIEVE_TEXT_FROM_BROWSER = "ERROR: UNABLE TO GET TRANSLATED TEXT FROM BROWSER";
        private const String ERROR_UNABLE_CLEAR_BROWSER = "ERROR: UNABLE TO CLEAR PREVIOUS TEXT ON BROWSER";
        private const String ERROR_UNABLE_SEND_TEXT_TO_BROWSER = "ERROR: UNABLE TO SEND TEXT TO BROWSER";
        private readonly string ERROR_BROWSER_NOT_OPEN = "ERROR: FIREFOX BROWSER IS NOT RUNNING";

        public bool checkIfBrowserIsOpen()
        {
            return driver != null;
        }

        public (bool success, string errorReason, string translatedText) translate(string term)
        {
            IWebElement googleTranslateInputBox;
            IWebElement googleTranslateCopyTextButton;

            WebDriverWait wait1Second = new WebDriverWait(driver, TimeSpan.FromSeconds(Convert.ToInt32(1)));

            try
            {
                ClipboardOps.ClearClipboardFromAnotherThread();

                googleTranslateInputBox =
                    wait1Second.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(
                        By.CssSelector(googleTranslateInputCssSelector)));

                googleTranslateInputBox.Clear();

                Actions actions = new Actions(driver);

                googleTranslateInputBox.Click();

                actions.KeyDown(Keys.Control).SendKeys("a").KeyUp(Keys.Control).Build().Perform();
                Thread.Sleep(200);
                actions.SendKeys(Keys.Delete).Build().Perform();

                Thread.Sleep(500);

                actions.SendKeys("a").Build().Perform();
                Thread.Sleep(200);
                actions.SendKeys(Keys.Backspace).Build().Perform();

                Thread.Sleep(500);
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
                return (false, ERROR_UNABLE_CLEAR_BROWSER, "");
            }

            Thread.Sleep(500);

            try
            {
                googleTranslateInputBox.SendKeys(term);

                //Enough fot 500 words
                Thread.Sleep(2000);
            }
            catch (Exception ex)
            {
                return (false, ERROR_UNABLE_SEND_TEXT_TO_BROWSER, "");
            }

            try
            {
                Debug.WriteLine("Retrieving translated text...");

                IJavaScriptExecutor jsExecutor = (IJavaScriptExecutor)driver;
                jsExecutor.ExecuteScript("window.scrollTo(0, document.body.scrollHeight);");

                Debug.WriteLine($"Awaiting for element with aria-label 'Copy translation'...");
                googleTranslateCopyTextButton = wait1Second.Until(SeleniumExtras.WaitHelpers.ExpectedConditions.ElementToBeClickable(By.CssSelector(googleTranslateCopyButtonCssSelector)));

                if (googleTranslateCopyTextButton == null)
                {
                    Debug.WriteLine($"Element with aria-label 'Copy translation' not found!");
                    return (false, ERROR_UNABLE_RETRIEVE_TEXT_FROM_BROWSER, "");
                }

                Debug.WriteLine($"Clearing clipboard...");

                ClipboardOps.ClearClipboardFromAnotherThread();

                Debug.WriteLine($"Clicking element with aria-label 'Copy translation'...");

                googleTranslateCopyTextButton.Click();

                // Wait for 1 second after clicking for better stability

                Thread.Sleep(1000);

                return (true, ClipboardOps.GetClipboardTextThreadSafe(), "");
            }
            catch (WebDriverTimeoutException)
            {
                Debug.WriteLine($"Element with aria-label 'Copy translation' not found within the timeout period!");
                return (false, ERROR_UNABLE_RETRIEVE_TEXT_FROM_BROWSER, "");
            }
            catch (Exception ex)
            {
                return (false, ERROR_UNABLE_RETRIEVE_TEXT_FROM_BROWSER, "");
            }
        }

        public (bool success, String errorReason) setTranslationLanguages(string sourceLang, string destLang)
        {
            if (!checkIfBrowserIsOpen())
            {
                return (false, ERROR_BROWSER_NOT_OPEN);
            }

            try
            {
                driver.Navigate().GoToUrl(GlobalVariables.googleTranslateURL.Replace("SOURCELANG", sourceLang).Replace("DESTINATIONLANG", destLang));
                return (true, "");
            }
            catch (Exception ex)
            {
                return (false, ex.ToString());
            }
        }

        public (bool success, string errorReason) init()
        {
            try
            {
                deleteStorageContents(GlobalVariables.googleTranslateSeleniumProfileName);

                Debug.WriteLine("Opening web browser...");

                FirefoxProfile profile = new FirefoxProfileManager().GetProfile(GlobalVariables.googleTranslateSeleniumProfileName);

                FirefoxOptions options = new FirefoxOptions();

                options.Profile = profile;

                driver = new FirefoxDriver(options);

                Debug.WriteLine("Opening web browser OK");

                var setLangResult = setTranslationLanguages(GlobalVariables.selectedSourceLanguage, GlobalVariables.selectedTargetLanguage);

                if (!setLangResult.success)
                {
                    return (false, $"Unable to set target / dest lang: {setLangResult.errorReason}");
                }

                return (true, "");
            }
            catch (Exception ex)
            {
                MessageBox.Show(new Form() { TopMost = true }, ex.ToString());
                return (false, ex.ToString());
            }
        }

        private static void deleteStorageContents(string partialProfileName)
        {
            try
            {
                string appDataPath = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                string firefoxProfilesPath = Path.Combine(appDataPath, "Mozilla\\Firefox\\Profiles");

                string[] matchingProfiles = Directory.GetDirectories(firefoxProfilesPath, "*" + partialProfileName + "*");

                foreach (string profile in matchingProfiles)
                {
                    string storagePath = Path.Combine(profile, "storage");

                    if (Directory.Exists(storagePath))
                    {
                        foreach (string file in Directory.GetFiles(storagePath))
                        {
                            File.Delete(file);
                        }

                        foreach (string subDirectory in Directory.GetDirectories(storagePath))
                        {
                            Directory.Delete(subDirectory, true);
                        }

                        //MessageBox.Show(new Form() { TopMost = true }, $"Contents of 'storage' folder in profile '{profile}' deleted.");
                    }
                    else
                    {
                        MessageBox.Show(new Form() { TopMost = true }, $"'storage' folder not found in profile '{profile}'.");
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error: " + ex.Message);
            }
        }

        public (bool success, string errorReason) terminate()
        {
            try
            {
                if (driver != null)
                {
                    driver.Close();
                    return (true, "");
                }
                return (false, "");
            }
            catch (Exception ex)
            {
                return (false, "");
            }
        }
    }
}