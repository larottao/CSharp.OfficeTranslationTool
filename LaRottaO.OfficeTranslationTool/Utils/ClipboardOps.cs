using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaRottaO.OfficeTranslationTool.Utils
{
    public static class ClipboardOps
    {
        public static void SetClipboardTextThreadSafe(String argText)
        {
            if (argText == null)
            {
                return;
            }

            Thread thread = new Thread((ThreadStart)delegate
            {
                Clipboard.SetText(argText);
            });

            try
            {
                thread.TrySetApartmentState(ApartmentState.STA);
                thread.Start();

                while (!thread.IsAlive) { Thread.Sleep(1); }

                Thread.Sleep(1);
                thread.Join();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to set text on clipboard: " + ex.ToString());
            }
        }

        public static String GetClipboardTextThreadSafe()
        {
            String clipboardText = "";

            Thread thread = new Thread((ThreadStart)delegate
            {
                try
                {
                    clipboardText = Clipboard.GetText();
                }
                catch (Exception ex)
                {
                    Console.WriteLine($"ERROR: {ex}");
                }
            });

            try
            {
                thread.TrySetApartmentState(ApartmentState.STA);
                thread.Start();

                while (!thread.IsAlive) { Thread.Sleep(1); }

                Thread.Sleep(1);
                thread.Join();

                return clipboardText;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to get text on clipboard: " + ex.ToString());
                return "";
            }
        }

        public static void ClearClipboardFromAnotherThread()
        {
            Thread thread = new Thread((ThreadStart)delegate
            {
                try
                {
                    Clipboard.Clear();
                }
                catch (Exception ex)
                {
                    Console.WriteLine("Unable to clear clipboard: " + ex.ToString());
                }
            });

            try
            {
                thread.TrySetApartmentState(ApartmentState.STA);
                thread.Start();

                while (!thread.IsAlive) { Thread.Sleep(1); }

                Thread.Sleep(1);
                thread.Join();
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to clear clipboard: " + ex.ToString());
            }
        }
    }
}