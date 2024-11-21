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
            dataGridView = new DataGridView();
            buttonOpenOfficeFile = new Button();
            panel1 = new Panel();
            buttonApplyChanges = new Button();
            label2 = new Label();
            label1 = new Label();
            buttonTranslateAll = new Button();
            comboBoxDestLanguage = new ComboBox();
            comboBoxSourceLanguage = new ComboBox();
            buttonRevertChanges = new Button();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView
            // 
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new Point(12, 83);
            dataGridView.Name = "dataGridView";
            dataGridView.Size = new Size(1080, 463);
            dataGridView.TabIndex = 0;
            dataGridView.CellBeginEdit += dataGridView_CellBeginEdit;          
            dataGridView.CellEndEdit += dataGridView_CellEndEdit;         
            dataGridView.DataBindingComplete += dataGridView_DataBindingComplete;
            dataGridView.RowEnter += dataGridView_RowEnter;
            dataGridView.RowPostPaint += dataGridView_RowPostPaint;
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
            // MainForm
            // 
            AutoScaleDimensions = new SizeF(7F, 15F);
            AutoScaleMode = AutoScaleMode.Font;
            ClientSize = new Size(1104, 558);
            Controls.Add(panel1);
            Controls.Add(dataGridView);
            Name = "MainForm";
            StartPosition = FormStartPosition.CenterScreen;
            Text = "LaRottaO Office Translation Tool";
            FormClosing += Form1_FormClosing;
            Load += Form1_Load;
            ((System.ComponentModel.ISupportInitialize)dataGridView).EndInit();
            panel1.ResumeLayout(false);
            ResumeLayout(false);
        }

        #endregion

        private DataGridView dataGridView;
        private Button buttonOpenOfficeFile;
        private Panel panel1;
        private ComboBox comboBoxDestLanguage;
        private ComboBox comboBoxSourceLanguage;
        private Button buttonTranslateAll;
        private Label label1;
        private Label label2;
        private Button buttonApplyChanges;
        private Button buttonRevertChanges;
    }
}
