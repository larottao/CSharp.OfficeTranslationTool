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

   

        public static void RefreshFromAnotherThreadKeepingFocus(this DataGridView dataGridView)
        {
            dataGridView.Invoke((MethodInvoker)delegate
            {
                // Save the index of the currently selected row and the currently focused cell
                int? selectedRowIndex = dataGridView.CurrentRow?.Index;
                int? selectedColumnIndex = dataGridView.CurrentCell?.ColumnIndex;

                // Refresh the DataGridView
                dataGridView.Refresh();

                // Restore the selected row and cell
                if (selectedRowIndex.HasValue && selectedRowIndex.Value >= 0 && selectedRowIndex.Value < dataGridView.Rows.Count)
                {
                    dataGridView.CurrentCell = dataGridView.Rows[selectedRowIndex.Value].Cells[selectedColumnIndex ?? 0];
                    dataGridView.Rows[selectedRowIndex.Value].Selected = true;

                    // Scroll to the selected row to ensure it's visible
                    dataGridView.FirstDisplayedScrollingRowIndex = selectedRowIndex.Value;
                }
            });
        }
    

}
}