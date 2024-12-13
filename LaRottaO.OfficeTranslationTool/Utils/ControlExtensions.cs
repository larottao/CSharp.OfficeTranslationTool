using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace LaRottaO.OfficeTranslationTool.Utils
{
    public static class ControlExtensions
    {
        public static T InvokeFromAnotherThread<T>(this Control control, Func<T> func)
        {
            if (control.InvokeRequired)
            {
                return (T)control.Invoke(func);
            }
            else
            {
                return func();
            }
        }

        public static void InvokeFromAnotherThread(this Control control, Action action)
        {
            if (control.InvokeRequired)
            {
                control.Invoke(action);
            }
            else
            {
                action();
            }
        }
    }
}