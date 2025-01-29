using LaRottaO.OfficeTranslationTool.Models;
using LaRottaO.OfficeTranslationTool.Utils;

using System.ComponentModel;
using static LaRottaO.OfficeTranslationTool.GlobalVariables;

namespace LaRottaO.OfficeTranslationTool
{
    public partial class MainForm : Form
    {
        private FormLogic formLogic;

        public MainForm()
        {
            InitializeComponent();

            formLogic = new FormLogic(this);

            mainDataGridView.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            mainDataGridView.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            mainDataGridView.MultiSelect = false;
            mainDataGridView.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;

            dataGridViewPartialExpressions.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.Fill;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            var loadSettingsResult = LoadProgramSettings.load();

            if (!loadSettingsResult.success)
            {
                UIHelpers.showErrorMessage(loadSettingsResult.errorReason);
            }

            foreach (KeyValuePair<String, String> lang in AVAILABLE_LANGUAGES)
            {
                comboBoxSourceLanguage.Items.Add(lang.Key);
                comboBoxDestLanguage.Items.Add(lang.Key);
            }

            foreach (KeyValuePair<String, TRANSLATION_METHOD> method in AVAILABLE_TRANSLATION_METHODS)
            {
                comboBoxTranslationMethod.Items.Add(method.Key);
            }

            comboBoxTranslationMethod.SelectedIndex = 0;
        }

        private async void buttonOpenOfficeFile_Click(object sender, EventArgs e)
        {
            await formLogic.launchSelectFileDialog();
        }

        //**************************************************
        //Changes the color of the DataGridView to stripes
        //**************************************************

        private void dataGridView_RowPostPaint(object sender, DataGridViewRowPostPaintEventArgs e)
        {
            for (int i = 0; i < mainDataGridView.Rows.Count; i++)
            {
                if (i % 2 == 0)
                {
                    mainDataGridView.Rows[i].DefaultCellStyle.BackColor = Color.LightBlue;
                }
            }
        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            formLogic.closeOfficeFile();

            //TODO ALSO CLOSE THE APP
        }

        private void button1_Click(object sender, EventArgs e)
        {
            formLogic.test();
        }

        private void dataGridView_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            foreach (DataGridViewColumn column in mainDataGridView.Columns)
            {
                var property = typeof(ElementToBeTranslated).GetProperty(column.DataPropertyName);
                if (property != null)
                {
                    // Check for [Browsable(false)]
                    if (property.GetCustomAttributes(typeof(BrowsableAttribute), true).FirstOrDefault() is BrowsableAttribute browsable && !browsable.Browsable)
                    {
                        column.Visible = false; // Hide the column
                    }

                    // Check for [ColumnName] and set the header text
                    if (property.GetCustomAttributes(typeof(ColumnNameAttribute), true).FirstOrDefault() is ColumnNameAttribute columnName)
                    {
                        column.HeaderText = columnName.Name;
                    }
                }
            }
        }

        private String? previousCellValue;

        private void dataGridView_CellBeginEdit(object sender, DataGridViewCellCancelEventArgs e)
        {
            if (replaceInProgress)
            {
                return;
            }

            if (!formLogic.areBothSourceAndDestintionLanguagesSet())
            {
                UIHelpers.showInformationMessage("Please select the Source and Target languages first.");

                //TODO this doesnt work
                mainDataGridView.ClearSelection();
                return;
            }

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var cell = mainDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];

                previousCellValue = cell.Value?.ToString();
            }
        }

        private void dataGridView_CellEndEdit(object sender, DataGridViewCellEventArgs e)
        {
            if (replaceInProgress)
            {
                return;
            }

            if (e.RowIndex >= 0 && e.ColumnIndex >= 0)
            {
                var cell = mainDataGridView.Rows[e.RowIndex].Cells[e.ColumnIndex];

                string? newValue = cell.Value?.ToString();

                if (newValue != null && !newValue.Equals(previousCellValue))
                {
                    formLogic.saveNewTranslationTypedByUserOnMainDgv(e.RowIndex, e.ColumnIndex, newValue);
                }
            }
        }

        private void comboBoxSourceLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            formLogic.setDictionaryLanguage(comboBoxSourceLanguage.Text, comboBoxDestLanguage.Text);
        }

        private void comboBoxDestLanguage_SelectedIndexChanged(object sender, EventArgs e)
        {
            formLogic.setDictionaryLanguage(comboBoxSourceLanguage.Text, comboBoxDestLanguage.Text);
        }

        private async void buttonTranslateAll_Click(object sender, EventArgs e)
        {
            if (!formLogic.areBothSourceAndDestintionLanguagesSet())
            {
                UIHelpers.showInformationMessage("Please select the Source and Target languages first.");
                return;
            }

            var transResult = await formLogic.translateAllShapeElements();

            if (transResult.success)
            {
                UIHelpers.showInformationMessage("Process complete");
            }
            else
            {
                UIHelpers.showErrorMessage(transResult.errorReason);
            }
        }

        private void mainDataGridView_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (replaceInProgress)
            {
                return;
            }

            if (mainDataGridView.Rows.Count == 0)
            {
                return;
            }

            formLogic.userClickedMainDataGridRow(e.RowIndex, e.ColumnIndex);
        }

        private async void buttonApplyChanges_Click(object sender, EventArgs e)
        {
            replaceInProgress = true;

            var replaceResult = await formLogic.applyChangesOnOfficeFile(false, true);

            if (replaceResult.success)
            {
                UIHelpers.showInformationMessage("Process complete");
            }
            else

            {
                UIHelpers.showErrorMessage(replaceResult.errorReason);
            }

            replaceInProgress = false;
        }

        private async void buttonRevertChanges_Click(object sender, EventArgs e)
        {
            var replaceResult = await formLogic.applyChangesOnOfficeFile(true, false);

            if (replaceResult.success)
            {
                UIHelpers.showInformationMessage("Process complete");
            }
            else

            {
                UIHelpers.showErrorMessage(replaceResult.errorReason);
            }
        }

        private void buttonLunchConfig_Click(object sender, EventArgs e)
        {
            if (!formLogic.areBothSourceAndDestintionLanguagesSet())
            {
                UIHelpers.showInformationMessage("Please select the Source and Target languages first.");
                return;
            }

            formLogic.populatePartialExpressionsDgv(dataGridViewPartialExpressions);

            setConfigPanelElementsVisibility(true);
        }

        private void setConfigPanelElementsVisibility(Boolean state)
        {
            buttonLunchConfig.Visible = !state;
            buttonSaveConfig.Visible = state;
            tabControlConfig.Visible = state;
            comboBoxSourceLanguage.Enabled = !state;
            comboBoxDestLanguage.Enabled = !state;
        }

        private void buttonSaveConfig_Click(object sender, EventArgs e)
        {
            setConfigPanelElementsVisibility(false);
        }

        private void buttonAddPartialExpressionToDic_Click(object sender, EventArgs e)
        {
            formLogic.addTranslationToDictionary(textBoxNewPartialExpTerm.Text.Trim(), textBoxNewPartialExpTermTrans.Text.Trim(), true);
        }

        private void dataGridViewPartialExpressions_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            if (dataGridViewPartialExpressions.Rows.Count == 0)
            {
                return;
            }

            formLogic.userClickedDgvPartialExpressionsRow(e.RowIndex, e.ColumnIndex);
        }

        private void buttonDeletePartialExpFromDic_Click(object sender, EventArgs e)
        {
            formLogic.deleteEntryFromPartialExpressionDic(textBoxNewPartialExpTerm.Text.Trim(), textBoxNewPartialExpTermTrans.Text.Trim());
        }

        private void dataGridViewPartialExpressions_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
        }
    }
}