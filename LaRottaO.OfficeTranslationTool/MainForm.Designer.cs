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
            buttonTranslateAll = new Button();
            comboBoxDestLanguage = new ComboBox();
            comboBoxSourceLanguage = new ComboBox();
            ((System.ComponentModel.ISupportInitialize)dataGridView).BeginInit();
            panel1.SuspendLayout();
            SuspendLayout();
            // 
            // dataGridView
            // 
            dataGridView.Anchor = AnchorStyles.Top | AnchorStyles.Bottom | AnchorStyles.Left | AnchorStyles.Right;
            dataGridView.ColumnHeadersHeightSizeMode = DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            dataGridView.Location = new Point(12, 70);
            dataGridView.Name = "dataGridView";
            dataGridView.Size = new Size(1080, 476);
            dataGridView.TabIndex = 0;
            dataGridView.CellBeginEdit += dataGridView_CellBeginEdit;
            dataGridView.CellContentClick += dataGridView_CellContentClick;
            dataGridView.CellEndEdit += dataGridView_CellEndEdit;
            dataGridView.CellValuePushed += dataGridView_CellValuePushed;
            dataGridView.DataBindingComplete += dataGridView_DataBindingComplete;
            dataGridView.RowEnter += dataGridView_RowEnter;
            dataGridView.RowPostPaint += dataGridView_RowPostPaint;
            // 
            // buttonOpenOfficeFile
            // 
            buttonOpenOfficeFile.FlatAppearance.BorderColor = Color.Silver;
            buttonOpenOfficeFile.FlatAppearance.BorderSize = 2;
            buttonOpenOfficeFile.FlatStyle = FlatStyle.Flat;
            buttonOpenOfficeFile.Location = new Point(12, 7);
            buttonOpenOfficeFile.Name = "buttonOpenOfficeFile";
            buttonOpenOfficeFile.Size = new Size(122, 37);
            buttonOpenOfficeFile.TabIndex = 1;
            buttonOpenOfficeFile.Text = "Open Office File";
            buttonOpenOfficeFile.UseVisualStyleBackColor = true;
            buttonOpenOfficeFile.Click += buttonOpenOfficeFile_Click;
            // 
            // panel1
            // 
            panel1.Controls.Add(buttonTranslateAll);
            panel1.Controls.Add(comboBoxDestLanguage);
            panel1.Controls.Add(comboBoxSourceLanguage);
            panel1.Controls.Add(buttonOpenOfficeFile);
            panel1.Dock = DockStyle.Top;
            panel1.Location = new Point(0, 0);
            panel1.Name = "panel1";
            panel1.Size = new Size(1104, 53);
            panel1.TabIndex = 2;
            // 
            // buttonTranslateAll
            // 
            buttonTranslateAll.FlatAppearance.BorderColor = Color.Silver;
            buttonTranslateAll.FlatAppearance.BorderSize = 2;
            buttonTranslateAll.FlatStyle = FlatStyle.Flat;
            buttonTranslateAll.Location = new Point(409, 7);
            buttonTranslateAll.Name = "buttonTranslateAll";
            buttonTranslateAll.Size = new Size(122, 37);
            buttonTranslateAll.TabIndex = 5;
            buttonTranslateAll.Text = "Translate All";
            buttonTranslateAll.UseVisualStyleBackColor = true;
            buttonTranslateAll.Click += buttonTranslateAll_Click;
            // 
            // comboBoxDestLanguage
            // 
            comboBoxDestLanguage.FormattingEnabled = true;
            comboBoxDestLanguage.Location = new Point(274, 15);
            comboBoxDestLanguage.Name = "comboBoxDestLanguage";
            comboBoxDestLanguage.Size = new Size(108, 23);
            comboBoxDestLanguage.TabIndex = 4;
            comboBoxDestLanguage.SelectedIndexChanged += comboBoxDestLanguage_SelectedIndexChanged;
            // 
            // comboBoxSourceLanguage
            // 
            comboBoxSourceLanguage.FormattingEnabled = true;
            comboBoxSourceLanguage.Location = new Point(160, 15);
            comboBoxSourceLanguage.Name = "comboBoxSourceLanguage";
            comboBoxSourceLanguage.Size = new Size(108, 23);
            comboBoxSourceLanguage.TabIndex = 3;
            comboBoxSourceLanguage.SelectedIndexChanged += comboBoxSourceLanguage_SelectedIndexChanged;
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
    }
}
