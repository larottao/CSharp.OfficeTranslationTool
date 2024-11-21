using LaRottaO.OfficeTranslationTool.Interfaces;
using System.Windows.Forms;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool.Utils
{
    internal static class UIHelpers
    {
        public static void showInformationMessage(string text)
        {
            MessageBox.Show($"{text}", "", MessageBoxButtons.OK, MessageBoxIcon.Information);
        }

        public static DialogResult showYesNoQuestion(string text)
        {
            return MessageBox.Show($"{text}", "", MessageBoxButtons.YesNo, MessageBoxIcon.Information);
        }

        public static void showErrorMessage(string text)
        {
            MessageBox.Show($"{text}", "", MessageBoxButtons.OK, MessageBoxIcon.Warning);
        }

        public static void offerToSaveDocumentBeforeExiting(IProcessOfficeFile iProcessOfficeFile)
        {
            if (currentOfficeDocPath == null)
            {
                return;
            }

            bool election = false;

            DialogResult saveBeforeClosingElection = MessageBox.Show(new Form() { TopMost = true }, "Do you want to save the changes made on the office document?", "Before exiting", MessageBoxButtons.YesNoCancel, MessageBoxIcon.Question);

            if (saveBeforeClosingElection == DialogResult.Yes) { election = true; }

            //TODO: Add Word

            if (iProcessOfficeFile.isOfficeProgramOpen())
            {
                iProcessOfficeFile.closeCurrentlyOpenFile(election);
            }
        }

        public static bool setCursorOnDataGridRowThreadSafe(this DataGridView dataGridView, int index, bool focusScrollBarToo = true)
        {
            try
            {
                if (index < 0)
                {
                    return false;
                }

                if (Thread.CurrentThread.IsBackground)
                {
                    dataGridView.Invoke(new Action(() =>
                    {
                        dataGridView.Rows[index].Selected = true;
                    }));
                }
                else
                {
                    dataGridView.Rows[index].Selected = true;
                }

                if (focusScrollBarToo)
                {
                    focusScrollOnCurrentSelectedRow(dataGridView);
                }

                return true;
            }
            catch (Exception ex)
            {
                Console.WriteLine("Unable to select DataGridView index: " + ex);
                return false;
            }
        }

        public static void focusScrollOnCurrentSelectedRow(DataGridView dataGridView)
        {
            if (Thread.CurrentThread.IsBackground)
            {
                dataGridView.Invoke(new Action(() =>
                {
                    focus(dataGridView);
                }));
            }
            else
            {
                focus(dataGridView);
            }
        }

        private static void focus(DataGridView dataGridView)
        {
            if (dataGridView.SelectedRows.Count > 0)
            {
                // Get the index of the first selected row
                int rowIndex = dataGridView.SelectedRows[0].Index;

                // Ensure the index is within the valid range
                if (rowIndex >= 0 && rowIndex < dataGridView.RowCount)
                {
                    // Check if the selected row is out of view and adjust the scrolling
                    if (rowIndex < dataGridView.FirstDisplayedScrollingRowIndex ||
                        rowIndex >= dataGridView.FirstDisplayedScrollingRowIndex + dataGridView.DisplayedRowCount(false))
                    {
                        dataGridView.FirstDisplayedScrollingRowIndex = Math.Max(0, rowIndex - dataGridView.DisplayedRowCount(false) / 2);
                    }
                }
            }
            else if (dataGridView.CurrentCell != null)
            {
                // Get the index of the row that contains the current cell
                int rowIndex = dataGridView.CurrentCell.RowIndex;

                // Ensure the index is within the valid range
                if (rowIndex >= 0 && rowIndex < dataGridView.RowCount)
                {
                    // Check if the current cell's row is out of view and adjust the scrolling
                    if (rowIndex < dataGridView.FirstDisplayedScrollingRowIndex ||
                        rowIndex >= dataGridView.FirstDisplayedScrollingRowIndex + dataGridView.DisplayedRowCount(false))
                    {
                        dataGridView.FirstDisplayedScrollingRowIndex = Math.Max(0, rowIndex - dataGridView.DisplayedRowCount(false) / 2);
                    }
                }
            }
        }
    }
}