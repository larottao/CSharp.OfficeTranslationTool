namespace LaRottaO.OfficeTranslationTool
{
    partial class MainForm
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            mainDataGridView = new DataGridView();
            buttonOpenOfficeFile = new Button();
            panel1 = new Panel();
            buttonLunchConfig = new Button();
            buttonRevertChanges = new Button();
            buttonApplyChanges = new Button();
            label2 = new Label();
            label1 = new Label();
            buttonTranslateAll = new Button();
            comboBoxDestLanguage = new ComboBox();
            comboBoxSourceLanguage = new ComboBox();
            dataGridViewPartialExpressions = new DataGridView();
            tabControlConfig = new TabControl();
            partialExpressionsTab = new TabPage();
            buttonDeletePartialExpFromDic = new Button();
            label4 = new Label();
            label3 = new Label();
            textBoxNewPartialExpTermTrans = new TextBox();
            textBoxNewPartialExpTerm = new TextBox();
            buttonAddPartialExpressionToDic = new Button();
            buttonSaveConfig = new Button();
            ((System.ComponentModel.ISupportInitialize)mainDataGridView).BeginInit();
            panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)dataGridViewPartialExpressions).BeginInit();
            tabControlConfig.SuspendLayout();
            partialExpressionsTab.SuspendLayout();
            SuspendLayout();
            // 
            // mainDataGridView
            // 
            mainDataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            mainDataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            mainDataGridView.Location = new Point(12, 83);
            mainDataGridView.Name = "mainDataGridView";
            mainDataGridView.Size = new Size(1080, 455);
            mainDataGridView.TabIndex = 0;
            mainDataGridView.CellBeginEdit += dataGridView_CellBeginEdit;
            mainDataGridView.CellEndEdit += dataGridView_CellEndEdit;
            mainDataGridView.DataBindingComplete += dataGridView_DataBindingComplete;
            mainDataGridView.RowEnter += mainDataGridView_RowEnter;
            mainDataGridView.RowPostPaint += dataGridView_RowPostPaint;
            // 
            // buttonOpenOfficeFile
            // 
            buttonOpenOfficeFile.FlatAppearance.BorderColor = Color.Silver;
            buttonOpenOfficeFile.FlatAppearance.BorderSize = 2;
            buttonOpenOfficeFile.FlatStyle = FlatStyle.Flat;
            buttonOpenOfficeFile.Location = new Point(12, 12);
            buttonOpenOfficeFile.Name = "buttonOpenOfficeFile";
            buttonOpenOfficeFile.Size = new Size(174, 37);
            buttonOpenOfficeFile.TabIndex = 1;
            buttonOpenOfficeFile.Text = "Open Office File";
            buttonOpenOfficeFile.UseVisualStyleBackColor = true;
            buttonOpenOfficeFile.Click += buttonOpenOfficeFile_Click;
            // 
            // panel1
            // 
            panel1.Controls.Add(buttonLunchConfig);
            panel1.Controls.Add(buttonRevertChanges);
            panel1.Controls.Add(buttonApplyChanges);
            panel1.Controls.Add(label2);
            panel1.Controls.Add(label1);
            panel1.Controls.Add(buttonTranslateAll);
            panel1.Controls.Add(comboBoxDestLanguage);
            panel1.Controls.Add(comboBoxSourceLanguage);
            panel1.Controls.Add(buttonOpenOfficeFile);
            panel1.Dock = DockStyle.Top;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1104, 64);
            panel1.TabIndex = 2;
            // 
            // buttonLunchConfig
            // 
            buttonLunchConfig.FlatAppearance.BorderColor = Color.Silver;
            buttonLunchConfig.FlatAppearance.BorderSize = 2;
            buttonLunchConfig.FlatStyle = FlatStyle.Flat;
            buttonLunchConfig.Location = new Point(1031, 12);
            buttonLunchConfig.Name = "buttonLunchConfig";
            buttonLunchConfig.Size = new Size(61, 37);
            buttonLunchConfig.TabIndex = 10;
            buttonLunchConfig.Text = "Config";
            buttonLunchConfig.UseVisualStyleBackColor = true;
            buttonLunchConfig.Click += buttonLunchConfig_Click;
            // 
            // buttonRevertChanges
            // 
            buttonRevertChanges.FlatAppearance.BorderColor = Color.Silver;
            buttonRevertChanges.FlatAppearance.BorderSize = 2;
            buttonRevertChanges.FlatStyle = FlatStyle.Flat;
            buttonRevertChanges.Location = new Point(840, 12);
            buttonRevertChanges.Name = "buttonRevertChanges";
            buttonRevertChanges.Size = new Size(174, 37);
            buttonRevertChanges.TabIndex = 9;
            buttonRevertChanges.Text = "Revert changes in document";
            buttonRevertChanges.UseVisualStyleBackColor = true;
            buttonRevertChanges.Click += buttonRevertChanges_Click;
            // 
            // buttonApplyChanges
            // 
            buttonApplyChanges.FlatAppearance.BorderColor = Color.Silver;
            buttonApplyChanges.FlatAppearance.BorderSize = 2;
            buttonApplyChanges.FlatStyle = FlatStyle.Flat;
            buttonApplyChanges.Location = new Point(648, 12);
            buttonApplyChanges.Name = "buttonApplyChanges";
            buttonApplyChanges.Size = new Size(174, 37);
            buttonApplyChanges.TabIndex = 8;
            buttonApplyChanges.Text = "Apply changes in document";
            buttonApplyChanges.UseVisualStyleBackColor = true;
            buttonApplyChanges.Click += buttonApplyChanges_Click;
            // 
            // label2
            // 
            label2.Font = new Font("Segoe UI", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label2.Location = new Point(330, 7);
            label2.Name = "label2";
            label2.Size = new Size(108, 13);
            label2.TabIndex = 7;
            label2.Text = "Destination lang.";
            label2.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // label1
            // 
            label1.Font = new Font("Segoe UI", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label1.Location = new Point(206, 7);
            label1.Name = "label1";
            label1.Size = new Size(108, 13);
            label1.TabIndex = 6;
            label1.Text = "Source lang.";
            label1.TextAlign = ContentAlignment.MiddleCenter;
            // 
            // buttonTranslateAll
            // 
            buttonTranslateAll.FlatAppearance.BorderColor = Color.Silver;
            buttonTranslateAll.FlatAppearance.BorderSize = 2;
            buttonTranslateAll.FlatStyle = FlatStyle.Flat;
            buttonTranslateAll.Location = new Point(456, 12);
            buttonTranslateAll.Name = "buttonTranslateAll";
            buttonTranslateAll.Size = new Size(174, 37);
            buttonTranslateAll.TabIndex = 5;
            buttonTranslateAll.Text = "Translate All";
            buttonTranslateAll.UseVisualStyleBackColor = true;
            buttonTranslateAll.Click += buttonTranslateAll_Click;
            // 
            // comboBoxDestLanguage
            // 
            comboBoxDestLanguage.BackColor = SystemColors.ControlLight;
            comboBoxDestLanguage.FlatStyle = FlatStyle.Flat;
            comboBoxDestLanguage.FormattingEnabled = true;
            comboBoxDestLanguage.Location = new Point(330, 24);
            comboBoxDestLanguage.Name = "comboBoxDestLanguage";
            comboBoxDestLanguage.Size = new Size(108, 23);
            comboBoxDestLanguage.TabIndex = 4;
            comboBoxDestLanguage.SelectedIndexChanged += comboBoxDestLanguage_SelectedIndexChanged;
            // 
            // comboBoxSourceLanguage
            // 
            comboBoxSourceLanguage.BackColor = SystemColors.ControlLight;
            comboBoxSourceLanguage.FlatStyle = FlatStyle.Flat;
            comboBoxSourceLanguage.FormattingEnabled = true;
            comboBoxSourceLanguage.Location = new Point(204, 24);
            comboBoxSourceLanguage.Name = "comboBoxSourceLanguage";
            comboBoxSourceLanguage.Size = new Size(108, 23);
            comboBoxSourceLanguage.TabIndex = 3;
            comboBoxSourceLanguage.SelectedIndexChanged += comboBoxSourceLanguage_SelectedIndexChanged;
            // 
            // dataGridViewPartialExpressions
            // 
            dataGridViewPartialExpressions.ColumnHeadersHeight = 50;
            dataGridViewPartialExpressions.Location = new Point(19, 18);
            dataGridViewPartialExpressions.Name = "dataGridViewPartialExpressions";
            dataGridViewPartialExpressions.RowHeadersWidthSizeMode = DataGridViewRowHeadersWidthSizeMode.AutoSizeToAllHeaders;
            dataGridViewPartialExpressions.Size = new Size(454, 389);
            dataGridViewPartialExpressions.TabIndex = 10;
            dataGridViewPartialExpressions.RowEnter += dataGridViewPartialExpressions_RowEnter;
            // 
            // tabControlConfig
            // 
            tabControlConfig.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            tabControlConfig.Controls.Add(partialExpressionsTab);
            tabControlConfig.Location = new Point(12, 83);
            tabControlConfig.Name = "tabControlConfig";
            tabControlConfig.SelectedIndex = 0;
            tabControlConfig.Size = new Size(1080, 455);
            tabControlConfig.TabIndex = 11;
            tabControlConfig.Visible = false;
            // 
            // partialExpressionsTab
            // 
            partialExpressionsTab.Controls.Add(buttonDeletePartialExpFromDic);
            partialExpressionsTab.Controls.Add(label4);
            partialExpressionsTab.Controls.Add(label3);
            partialExpressionsTab.Controls.Add(textBoxNewPartialExpTermTrans);
            partialExpressionsTab.Controls.Add(textBoxNewPartialExpTerm);
            partialExpressionsTab.Controls.Add(buttonAddPartialExpressionToDic);
            partialExpressionsTab.Controls.Add(dataGridViewPartialExpressions);
            partialExpressionsTab.Location = new Point(4, 24);
            partialExpressionsTab.Name = "partialExpressionsTab";
            partialExpressionsTab.Padding = new Padding(3);
            partialExpressionsTab.Size = new Size(1072, 427);
            partialExpressionsTab.TabIndex = 0;
            partialExpressionsTab.Text = "Partial Expressions";
            partialExpressionsTab.UseVisualStyleBackColor = true;
            // 
            // buttonDeletePartialExpFromDic
            // 
            buttonDeletePartialExpFromDic.FlatAppearance.BorderColor = Color.Silver;
            buttonDeletePartialExpFromDic.FlatAppearance.BorderSize = 2;
            buttonDeletePartialExpFromDic.FlatStyle = FlatStyle.Flat;
            buttonDeletePartialExpFromDic.Location = new Point(679, 130);
            buttonDeletePartialExpFromDic.Name = "buttonDeletePartialExpFromDic";
            buttonDeletePartialExpFromDic.Size = new Size(174, 37);
            buttonDeletePartialExpFromDic.TabIndex = 16;
            buttonDeletePartialExpFromDic.Text = "Delete from Dictionary";
            buttonDeletePartialExpFromDic.UseVisualStyleBackColor = true;
            buttonDeletePartialExpFromDic.Click += buttonDeletePartialExpFromDic_Click;
            // 
            // label4
            // 
            label4.Font = new Font("Segoe UI", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label4.Location = new Point(499, 71);
            label4.Name = "label4";
            label4.Size = new Size(322, 13);
            label4.TabIndex = 15;
            label4.Text = "Replace with";
            label4.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // label3
            // 
            label3.Font = new Font("Segoe UI", 8.25F, FontStyle.Regular, GraphicsUnit.Point, 0);
            label3.Location = new Point(499, 18);
            label3.Name = "label3";
            label3.Size = new Size(322, 13);
            label3.TabIndex = 14;
            label3.Text = "Look for this text";
            label3.TextAlign = ContentAlignment.MiddleLeft;
            // 
            // textBoxNewPartialExpTermTrans
            // 
            textBoxNewPartialExpTermTrans.Location = new Point(499, 87);
            textBoxNewPartialExpTermTrans.Name = "textBoxNewPartialExpTermTrans";
            textBoxNewPartialExpTermTrans.Size = new Size(354, 23);
            textBoxNewPartialExpTermTrans.TabIndex = 13;
            // 
            // textBoxNewPartialExpTerm
            // 
            textBoxNewPartialExpTerm.Location = new Point(499, 34);
            textBoxNewPartialExpTerm.Name = "textBoxNewPartialExpTerm";
            textBoxNewPartialExpTerm.Size = new Size(354, 23);
            textBoxNewPartialExpTerm.TabIndex = 12;
            // 
            // buttonAddPartialExpressionToDic
            // 
            buttonAddPartialExpressionToDic.FlatAppearance.BorderColor = Color.Silver;
            buttonAddPartialExpressionToDic.FlatAppearance.BorderSize = 2;
            buttonAddPartialExpressionToDic.FlatStyle = FlatStyle.Flat;
            buttonAddPartialExpressionToDic.Location = new Point(499, 130);
            buttonAddPartialExpressionToDic.Name = "buttonAddPartialExpressionToDic";
            buttonAddPartialExpressionToDic.Size = new Size(174, 37);
            buttonAddPartialExpressionToDic.TabIndex = 11;
            buttonAddPartialExpressionToDic.Text = "Add to Dictionary";
            buttonAddPartialExpressionToDic.UseVisualStyleBackColor = true;
            buttonAddPartialExpressionToDic.Click += buttonAddPartialExpressionToDic_Click;
            // 
            // buttonSaveConfig
            // 
            buttonSaveConfig.FlatAppearance.BorderColor = Color.Silver;
            buttonSaveConfig.FlatAppearance.BorderSize = 2;
            buttonSaveConfig.FlatStyle = FlatStyle.Flat;
            buttonSaveConfig.Location = new Point(914, 544);
            buttonSaveConfig.Name = "buttonSaveConfig";
            buttonSaveConfig.Size = new Size(174, 37);
            buttonSaveConfig.TabIndex = 12;
            buttonSaveConfig.Text = "Save config";
            buttonSaveConfig.UseVisualStyleBackColor = true;
            buttonSaveConfig.Visible = false;
            buttonSaveConfig.Click += buttonSaveConfig_Click;
            // 
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1104, 589);
            Controls.Add(buttonSaveConfig);
            Controls.Add(tabControlConfig);
            Controls.Add(panel1);
            Controls.Add(mainDataGridView);
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "LaRottaO Office Translation Tool";
            FormClosing += Form1_FormClosing;
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)mainDataGridView).EndInit();
            panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)dataGridViewPartialExpressions).EndInit();
            tabControlConfig.ResumeLayout(false);
            partialExpressionsTab.ResumeLayout(false);
            partialExpressionsTab.PerformLayout();
            ResumeLayout(false);
        }

        #endregion
        private Button buttonOpenOfficeFile;
        private Panel panel1;
        private ComboBox comboBoxDestLanguage;
        private ComboBox comboBoxSourceLanguage;
        private Button buttonTranslateAll;
        private Label label1;
        private Label label2;
        private Button buttonApplyChanges;
        private Button buttonRevertChanges;
        private Button buttonLunchConfig;
        private TabControl tabControlConfig;
        private TabPage partialExpressionsTab;
        private Button buttonSaveConfig;
        private Label label4;
        private Label label3;
        private Button buttonAddPartialExpressionToDic;
        public DataGridView mainDataGridView;
        public TextBox textBoxNewPartialExpTermTrans;
        public TextBox textBoxNewPartialExpTerm;
        private Button buttonDeletePartialExpFromDic;
        public DataGridView dataGridViewPartialExpressions;
    }
}
