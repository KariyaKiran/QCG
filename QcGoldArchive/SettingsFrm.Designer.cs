namespace QcGoldArchive
{
    partial class SettingsFrm
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
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
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(SettingsFrm));
            this.OpenFileButton = new System.Windows.Forms.Button();
            this.FilePath = new System.Windows.Forms.TextBox();
            this.FileNameLabel = new System.Windows.Forms.Label();
            this.ExportCheck = new System.Windows.Forms.CheckBox();
            this.selectPortLabel = new System.Windows.Forms.Label();
            this.PortsSelectionBox = new System.Windows.Forms.ComboBox();
            this.SaveButton = new System.Windows.Forms.Button();
            this.CancelButton = new System.Windows.Forms.Button();
            this.ImportSet = new System.Windows.Forms.GroupBox();
            this.LanguageLabel = new System.Windows.Forms.Label();
            this.LanguageSelectionBox = new System.Windows.Forms.ComboBox();
            this.HeaderSpaceCheckBox = new System.Windows.Forms.CheckBox();
            this.MakePdfReportRadioButton = new System.Windows.Forms.RadioButton();
            this.HeaderSpace = new System.Windows.Forms.TextBox();
            this.HeaderSpaceLabel = new System.Windows.Forms.Label();
            this.ReportSet = new System.Windows.Forms.GroupBox();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.PrinterName = new System.Windows.Forms.TextBox();
            this.ChoosePrinterButton = new System.Windows.Forms.Button();
            this.PrintResultsRadioButton = new System.Windows.Forms.RadioButton();
            this.PrinterNameLabel = new System.Windows.Forms.Label();
            this.AutoPrintCheck = new System.Windows.Forms.CheckBox();
            this.pdfcheck = new System.Windows.Forms.CheckBox();
            this.Digitalsignaturepath = new System.Windows.Forms.TextBox();
            this.addsign = new System.Windows.Forms.Label();
            this.usedefault = new System.Windows.Forms.CheckBox();
            this.DSpath = new System.Windows.Forms.Button();
            this.subtitle = new System.Windows.Forms.Label();
            this.subtitleval = new System.Windows.Forms.TextBox();
            this.titlelab = new System.Windows.Forms.Label();
            this.title = new System.Windows.Forms.TextBox();
            this.rpttitle = new System.Windows.Forms.Label();
            this.ReportType = new System.Windows.Forms.ComboBox();
            this.rpttype = new System.Windows.Forms.Label();
            this.DateFormatSet = new System.Windows.Forms.ComboBox();
            this.DateFormatLabel = new System.Windows.Forms.Label();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.label2 = new System.Windows.Forms.Label();
            this.optionalval2 = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.optionalval = new System.Windows.Forms.TextBox();
            this.removeFooter = new System.Windows.Forms.CheckBox();
            this.ImportSet.SuspendLayout();
            this.ReportSet.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.SuspendLayout();
            // 
            // OpenFileButton
            // 
            this.OpenFileButton.Location = new System.Drawing.Point(390, 46);
            this.OpenFileButton.Name = "OpenFileButton";
            this.OpenFileButton.Size = new System.Drawing.Size(87, 23);
            this.OpenFileButton.TabIndex = 10;
            this.OpenFileButton.Text = "Choose";
            this.OpenFileButton.UseVisualStyleBackColor = true;
            this.OpenFileButton.Click += new System.EventHandler(this.OpenFileButton_Click);
            // 
            // FilePath
            // 
            this.FilePath.Location = new System.Drawing.Point(17, 47);
            this.FilePath.Name = "FilePath";
            this.FilePath.Size = new System.Drawing.Size(359, 20);
            this.FilePath.TabIndex = 9;
            this.FilePath.TextChanged += new System.EventHandler(this.FilePath_Changed);
            // 
            // FileNameLabel
            // 
            this.FileNameLabel.AutoSize = true;
            this.FileNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FileNameLabel.Location = new System.Drawing.Point(15, 31);
            this.FileNameLabel.Name = "FileNameLabel";
            this.FileNameLabel.Size = new System.Drawing.Size(67, 13);
            this.FileNameLabel.TabIndex = 8;
            this.FileNameLabel.Text = "File Name:";
            this.FileNameLabel.Click += new System.EventHandler(this.FileNameLabel_Click);
            // 
            // ExportCheck
            // 
            this.ExportCheck.Location = new System.Drawing.Point(38, 80);
            this.ExportCheck.Name = "ExportCheck";
            this.ExportCheck.Size = new System.Drawing.Size(514, 18);
            this.ExportCheck.TabIndex = 7;
            this.ExportCheck.Text = "Export to Excel";
            this.ExportCheck.UseVisualStyleBackColor = true;
            this.ExportCheck.Visible = false;
            // 
            // selectPortLabel
            // 
            this.selectPortLabel.AutoSize = true;
            this.selectPortLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.selectPortLabel.Location = new System.Drawing.Point(42, 31);
            this.selectPortLabel.Name = "selectPortLabel";
            this.selectPortLabel.Size = new System.Drawing.Size(74, 13);
            this.selectPortLabel.TabIndex = 11;
            this.selectPortLabel.Text = "Select Port:";
            // 
            // PortsSelectionBox
            // 
            this.PortsSelectionBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.PortsSelectionBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PortsSelectionBox.FormattingEnabled = true;
            this.PortsSelectionBox.Location = new System.Drawing.Point(167, 28);
            this.PortsSelectionBox.Name = "PortsSelectionBox";
            this.PortsSelectionBox.Size = new System.Drawing.Size(73, 21);
            this.PortsSelectionBox.TabIndex = 12;
            this.PortsSelectionBox.SelectedIndexChanged += new System.EventHandler(this.Port_Changed);
            // 
            // SaveButton
            // 
            this.SaveButton.Location = new System.Drawing.Point(317, 470);
            this.SaveButton.Name = "SaveButton";
            this.SaveButton.Size = new System.Drawing.Size(87, 23);
            this.SaveButton.TabIndex = 14;
            this.SaveButton.Text = "Save";
            this.SaveButton.UseVisualStyleBackColor = true;
            this.SaveButton.Click += new System.EventHandler(this.SaveButton_Click);
            // 
            // CancelButton
            // 
            this.CancelButton.Location = new System.Drawing.Point(535, 470);
            this.CancelButton.Name = "CancelButton";
            this.CancelButton.Size = new System.Drawing.Size(87, 23);
            this.CancelButton.TabIndex = 15;
            this.CancelButton.Text = "Cancel";
            this.CancelButton.UseVisualStyleBackColor = true;
            this.CancelButton.Click += new System.EventHandler(this.CancelButton_Click);
            // 
            // ImportSet
            // 
            this.ImportSet.Controls.Add(this.FileNameLabel);
            this.ImportSet.Controls.Add(this.FilePath);
            this.ImportSet.Controls.Add(this.OpenFileButton);
            this.ImportSet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ImportSet.Location = new System.Drawing.Point(28, 60);
            this.ImportSet.Name = "ImportSet";
            this.ImportSet.Size = new System.Drawing.Size(493, 79);
            this.ImportSet.TabIndex = 16;
            this.ImportSet.TabStop = false;
            this.ImportSet.Text = "Export to Excel Settings";
            this.ImportSet.Enter += new System.EventHandler(this.ImportSet_Enter);
            // 
            // LanguageLabel
            // 
            this.LanguageLabel.AutoSize = true;
            this.LanguageLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LanguageLabel.Location = new System.Drawing.Point(283, 31);
            this.LanguageLabel.Name = "LanguageLabel";
            this.LanguageLabel.Size = new System.Drawing.Size(67, 13);
            this.LanguageLabel.TabIndex = 17;
            this.LanguageLabel.Text = "Language:";
            // 
            // LanguageSelectionBox
            // 
            this.LanguageSelectionBox.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.LanguageSelectionBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LanguageSelectionBox.FormattingEnabled = true;
            this.LanguageSelectionBox.Location = new System.Drawing.Point(367, 28);
            this.LanguageSelectionBox.Name = "LanguageSelectionBox";
            this.LanguageSelectionBox.Size = new System.Drawing.Size(154, 21);
            this.LanguageSelectionBox.TabIndex = 18;
            this.LanguageSelectionBox.SelectedIndexChanged += new System.EventHandler(this.Languange_Changed);
            // 
            // HeaderSpaceCheckBox
            // 
            this.HeaderSpaceCheckBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HeaderSpaceCheckBox.Location = new System.Drawing.Point(19, 244);
            this.HeaderSpaceCheckBox.Name = "HeaderSpaceCheckBox";
            this.HeaderSpaceCheckBox.Size = new System.Drawing.Size(194, 28);
            this.HeaderSpaceCheckBox.TabIndex = 19;
            this.HeaderSpaceCheckBox.Text = "Add header space:";
            this.HeaderSpaceCheckBox.UseVisualStyleBackColor = true;
            this.HeaderSpaceCheckBox.CheckedChanged += new System.EventHandler(this.HeaderCheck_Changed);
            // 
            // MakePdfReportRadioButton
            // 
            this.MakePdfReportRadioButton.AutoSize = true;
            this.MakePdfReportRadioButton.Checked = true;
            this.MakePdfReportRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.MakePdfReportRadioButton.Location = new System.Drawing.Point(18, 26);
            this.MakePdfReportRadioButton.Name = "MakePdfReportRadioButton";
            this.MakePdfReportRadioButton.Size = new System.Drawing.Size(121, 17);
            this.MakePdfReportRadioButton.TabIndex = 21;
            this.MakePdfReportRadioButton.TabStop = true;
            this.MakePdfReportRadioButton.Text = "Create pdf report";
            this.MakePdfReportRadioButton.UseVisualStyleBackColor = true;
            this.MakePdfReportRadioButton.CheckedChanged += new System.EventHandler(this.MakePdf_Changed);
            // 
            // HeaderSpace
            // 
            this.HeaderSpace.Location = new System.Drawing.Point(210, 249);
            this.HeaderSpace.Name = "HeaderSpace";
            this.HeaderSpace.ShortcutsEnabled = false;
            this.HeaderSpace.Size = new System.Drawing.Size(50, 20);
            this.HeaderSpace.TabIndex = 22;
            this.HeaderSpace.TextChanged += new System.EventHandler(this.HeaderSpace_TextChanged);
            this.HeaderSpace.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.HeaderSpace_KeyPress);
            // 
            // HeaderSpaceLabel
            // 
            this.HeaderSpaceLabel.AutoSize = true;
            this.HeaderSpaceLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.HeaderSpaceLabel.Location = new System.Drawing.Point(266, 253);
            this.HeaderSpaceLabel.Name = "HeaderSpaceLabel";
            this.HeaderSpaceLabel.Size = new System.Drawing.Size(30, 16);
            this.HeaderSpaceLabel.TabIndex = 23;
            this.HeaderSpaceLabel.Text = "mm";
            // 
            // ReportSet
            // 
            this.ReportSet.Controls.Add(this.removeFooter);
            this.ReportSet.Controls.Add(this.groupBox1);
            this.ReportSet.Controls.Add(this.PrinterName);
            this.ReportSet.Controls.Add(this.ChoosePrinterButton);
            this.ReportSet.Controls.Add(this.PrintResultsRadioButton);
            this.ReportSet.Controls.Add(this.PrinterNameLabel);
            this.ReportSet.Controls.Add(this.AutoPrintCheck);
            this.ReportSet.Controls.Add(this.pdfcheck);
            this.ReportSet.Controls.Add(this.Digitalsignaturepath);
            this.ReportSet.Controls.Add(this.addsign);
            this.ReportSet.Controls.Add(this.usedefault);
            this.ReportSet.Controls.Add(this.DSpath);
            this.ReportSet.Controls.Add(this.subtitle);
            this.ReportSet.Controls.Add(this.subtitleval);
            this.ReportSet.Controls.Add(this.titlelab);
            this.ReportSet.Controls.Add(this.title);
            this.ReportSet.Controls.Add(this.rpttitle);
            this.ReportSet.Controls.Add(this.HeaderSpace);
            this.ReportSet.Controls.Add(this.ReportType);
            this.ReportSet.Controls.Add(this.HeaderSpaceCheckBox);
            this.ReportSet.Controls.Add(this.HeaderSpaceLabel);
            this.ReportSet.Controls.Add(this.rpttype);
            this.ReportSet.Controls.Add(this.MakePdfReportRadioButton);
            this.ReportSet.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.ReportSet.Location = new System.Drawing.Point(29, 149);
            this.ReportSet.Name = "ReportSet";
            this.ReportSet.Size = new System.Drawing.Size(853, 312);
            this.ReportSet.TabIndex = 24;
            this.ReportSet.TabStop = false;
            this.ReportSet.Text = "Report Settings";
            this.ReportSet.Enter += new System.EventHandler(this.ReportSet_Enter);
            // 
            // groupBox1
            // 
            this.groupBox1.Location = new System.Drawing.Point(422, 2);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(2, 309);
            this.groupBox1.TabIndex = 46;
            this.groupBox1.TabStop = false;
            // 
            // PrinterName
            // 
            this.PrinterName.Enabled = false;
            this.PrinterName.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(177)));
            this.PrinterName.Location = new System.Drawing.Point(541, 51);
            this.PrinterName.Name = "PrinterName";
            this.PrinterName.Size = new System.Drawing.Size(145, 22);
            this.PrinterName.TabIndex = 40;
            this.PrinterName.TextChanged += new System.EventHandler(this.PrinterName_TextChanged_1);
            // 
            // ChoosePrinterButton
            // 
            this.ChoosePrinterButton.Location = new System.Drawing.Point(702, 51);
            this.ChoosePrinterButton.Name = "ChoosePrinterButton";
            this.ChoosePrinterButton.Size = new System.Drawing.Size(87, 23);
            this.ChoosePrinterButton.TabIndex = 41;
            this.ChoosePrinterButton.Text = "Choose";
            this.ChoosePrinterButton.UseVisualStyleBackColor = true;
            this.ChoosePrinterButton.Click += new System.EventHandler(this.ChoosePrinterButton_Click_1);
            // 
            // PrintResultsRadioButton
            // 
            this.PrintResultsRadioButton.AutoSize = true;
            this.PrintResultsRadioButton.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PrintResultsRadioButton.Location = new System.Drawing.Point(438, 26);
            this.PrintResultsRadioButton.Name = "PrintResultsRadioButton";
            this.PrintResultsRadioButton.Size = new System.Drawing.Size(114, 17);
            this.PrintResultsRadioButton.TabIndex = 43;
            this.PrintResultsRadioButton.Text = "Print result strip";
            this.PrintResultsRadioButton.UseVisualStyleBackColor = true;
            this.PrintResultsRadioButton.CheckedChanged += new System.EventHandler(this.PrintResultsRadioButton_CheckedChanged_2);
            // 
            // PrinterNameLabel
            // 
            this.PrinterNameLabel.AutoSize = true;
            this.PrinterNameLabel.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.PrinterNameLabel.Location = new System.Drawing.Point(436, 56);
            this.PrinterNameLabel.Name = "PrinterNameLabel";
            this.PrinterNameLabel.Size = new System.Drawing.Size(73, 13);
            this.PrinterNameLabel.TabIndex = 39;
            this.PrinterNameLabel.Text = "Select Printer:";
            // 
            // AutoPrintCheck
            // 
            this.AutoPrintCheck.AutoSize = true;
            this.AutoPrintCheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.AutoPrintCheck.Location = new System.Drawing.Point(439, 89);
            this.AutoPrintCheck.Name = "AutoPrintCheck";
            this.AutoPrintCheck.Size = new System.Drawing.Size(72, 17);
            this.AutoPrintCheck.TabIndex = 42;
            this.AutoPrintCheck.Text = "Auto Print";
            this.AutoPrintCheck.UseVisualStyleBackColor = true;
            this.AutoPrintCheck.CheckedChanged += new System.EventHandler(this.AutoPrintCheck_CheckedChanged);
            // 
            // pdfcheck
            // 
            this.pdfcheck.AutoSize = true;
            this.pdfcheck.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.pdfcheck.Location = new System.Drawing.Point(161, 282);
            this.pdfcheck.Name = "pdfcheck";
            this.pdfcheck.Size = new System.Drawing.Size(72, 17);
            this.pdfcheck.TabIndex = 38;
            this.pdfcheck.Text = "Auto Print";
            this.pdfcheck.UseVisualStyleBackColor = true;
            this.pdfcheck.CheckedChanged += new System.EventHandler(this.pdfcheck_CheckedChanged);
            // 
            // Digitalsignaturepath
            // 
            this.Digitalsignaturepath.Location = new System.Drawing.Point(134, 197);
            this.Digitalsignaturepath.Name = "Digitalsignaturepath";
            this.Digitalsignaturepath.Size = new System.Drawing.Size(171, 20);
            this.Digitalsignaturepath.TabIndex = 25;
            this.Digitalsignaturepath.TextChanged += new System.EventHandler(this.Digitalsignaturepath_TextChanged);
            // 
            // addsign
            // 
            this.addsign.AutoSize = true;
            this.addsign.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.addsign.Location = new System.Drawing.Point(15, 200);
            this.addsign.Name = "addsign";
            this.addsign.Size = new System.Drawing.Size(77, 13);
            this.addsign.TabIndex = 27;
            this.addsign.Text = "Add Signature:";
            // 
            // usedefault
            // 
            this.usedefault.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.usedefault.Location = new System.Drawing.Point(106, 92);
            this.usedefault.Name = "usedefault";
            this.usedefault.Size = new System.Drawing.Size(215, 18);
            this.usedefault.TabIndex = 37;
            this.usedefault.Text = "Use default";
            this.usedefault.UseVisualStyleBackColor = true;
            this.usedefault.CheckedChanged += new System.EventHandler(this.Usedefault_CheckedChanged);
            // 
            // DSpath
            // 
            this.DSpath.Location = new System.Drawing.Point(309, 196);
            this.DSpath.Name = "DSpath";
            this.DSpath.Size = new System.Drawing.Size(87, 23);
            this.DSpath.TabIndex = 26;
            this.DSpath.Text = "Choose";
            this.DSpath.UseVisualStyleBackColor = true;
            this.DSpath.Click += new System.EventHandler(this.DSpath_Click);
            // 
            // subtitle
            // 
            this.subtitle.AutoSize = true;
            this.subtitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.subtitle.Location = new System.Drawing.Point(15, 158);
            this.subtitle.Name = "subtitle";
            this.subtitle.Size = new System.Drawing.Size(45, 13);
            this.subtitle.TabIndex = 36;
            this.subtitle.Text = "Subtitle:";
            // 
            // subtitleval
            // 
            this.subtitleval.Location = new System.Drawing.Point(72, 155);
            this.subtitleval.Name = "subtitleval";
            this.subtitleval.ShortcutsEnabled = false;
            this.subtitleval.Size = new System.Drawing.Size(324, 20);
            this.subtitleval.TabIndex = 35;
            this.subtitleval.TextChanged += new System.EventHandler(this.subtitleval_TextChanged);
            // 
            // titlelab
            // 
            this.titlelab.AutoSize = true;
            this.titlelab.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.titlelab.Location = new System.Drawing.Point(15, 120);
            this.titlelab.Name = "titlelab";
            this.titlelab.Size = new System.Drawing.Size(30, 13);
            this.titlelab.TabIndex = 34;
            this.titlelab.Text = "Title:";
            // 
            // title
            // 
            this.title.Location = new System.Drawing.Point(72, 117);
            this.title.Name = "title";
            this.title.ShortcutsEnabled = false;
            this.title.Size = new System.Drawing.Size(324, 20);
            this.title.TabIndex = 33;
            this.title.TextChanged += new System.EventHandler(this.title_TextChanged);
            // 
            // rpttitle
            // 
            this.rpttitle.AutoSize = true;
            this.rpttitle.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.rpttitle.Location = new System.Drawing.Point(15, 93);
            this.rpttitle.Name = "rpttitle";
            this.rpttitle.Size = new System.Drawing.Size(62, 13);
            this.rpttitle.TabIndex = 32;
            this.rpttitle.Text = "Report Title";
            // 
            // ReportType
            // 
            this.ReportType.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.ReportType.FormattingEnabled = true;
            this.ReportType.Location = new System.Drawing.Point(107, 58);
            this.ReportType.Name = "ReportType";
            this.ReportType.Size = new System.Drawing.Size(126, 21);
            this.ReportType.TabIndex = 29;
            this.ReportType.SelectedIndexChanged += new System.EventHandler(this.ReportType_SelectedIndexChanged);
            this.ReportType.TextChanged += new System.EventHandler(this.ReportType_TextChanged_1);
            // 
            // rpttype
            // 
            this.rpttype.AutoSize = true;
            this.rpttype.Location = new System.Drawing.Point(15, 59);
            this.rpttype.Name = "rpttype";
            this.rpttype.Size = new System.Drawing.Size(69, 13);
            this.rpttype.TabIndex = 28;
            this.rpttype.Text = "Report Type:";
            // 
            // DateFormatSet
            // 
            this.DateFormatSet.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.DateFormatSet.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            this.DateFormatSet.FormattingEnabled = true;
            this.DateFormatSet.Items.AddRange(new object[] {
            "DD/MM/YY",
            "MM/DD/YY"});
            this.DateFormatSet.Location = new System.Drawing.Point(696, 23);
            this.DateFormatSet.Name = "DateFormatSet";
            this.DateFormatSet.Size = new System.Drawing.Size(104, 21);
            this.DateFormatSet.TabIndex = 48;
            this.DateFormatSet.Visible = false;
            // 
            // DateFormatLabel
            // 
            this.DateFormatLabel.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.DateFormatLabel.AutoSize = true;
            this.DateFormatLabel.Location = new System.Drawing.Point(619, 28);
            this.DateFormatLabel.Name = "DateFormatLabel";
            this.DateFormatLabel.Size = new System.Drawing.Size(84, 13);
            this.DateFormatLabel.TabIndex = 47;
            this.DateFormatLabel.Text = "Date Format: ";
            this.DateFormatLabel.Visible = false;
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.label2);
            this.groupBox2.Controls.Add(this.optionalval2);
            this.groupBox2.Controls.Add(this.label1);
            this.groupBox2.Controls.Add(this.optionalval);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(521, 60);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(355, 79);
            this.groupBox2.TabIndex = 49;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "Data Entry";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label2.Location = new System.Drawing.Point(15, 54);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(80, 13);
            this.label2.TabIndex = 10;
            this.label2.Text = "OPTIONAL2:";
            // 
            // optionalval2
            // 
            this.optionalval2.Location = new System.Drawing.Point(101, 51);
            this.optionalval2.Name = "optionalval2";
            this.optionalval2.Size = new System.Drawing.Size(165, 20);
            this.optionalval2.TabIndex = 11;
            this.optionalval2.TextChanged += new System.EventHandler(this.Optionalval2_TextChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.label1.Location = new System.Drawing.Point(15, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(80, 13);
            this.label1.TabIndex = 8;
            this.label1.Text = "OPTIONAL1:";
            // 
            // optionalval
            // 
            this.optionalval.Location = new System.Drawing.Point(101, 20);
            this.optionalval.Name = "optionalval";
            this.optionalval.Size = new System.Drawing.Size(165, 20);
            this.optionalval.TabIndex = 9;
            this.optionalval.TextChanged += new System.EventHandler(this.Optionalval_TextChanged);
            // 
            // removeFooter
            // 
            this.removeFooter.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.removeFooter.Location = new System.Drawing.Point(19, 275);
            this.removeFooter.Name = "removeFooter";
            this.removeFooter.Size = new System.Drawing.Size(120, 28);
            this.removeFooter.TabIndex = 47;
            this.removeFooter.Text = "Remove Footer ";
            this.removeFooter.UseVisualStyleBackColor = true;
            // 
            // SettingsFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.SystemColors.ButtonHighlight;
            this.ClientSize = new System.Drawing.Size(911, 508);
            this.Controls.Add(this.groupBox2);
            this.Controls.Add(this.DateFormatSet);
            this.Controls.Add(this.DateFormatLabel);
            this.Controls.Add(this.ReportSet);
            this.Controls.Add(this.LanguageSelectionBox);
            this.Controls.Add(this.LanguageLabel);
            this.Controls.Add(this.ImportSet);
            this.Controls.Add(this.CancelButton);
            this.Controls.Add(this.SaveButton);
            this.Controls.Add(this.PortsSelectionBox);
            this.Controls.Add(this.selectPortLabel);
            this.Controls.Add(this.ExportCheck);
            this.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "SettingsFrm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Settings";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.SaveChanges);
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.SettingsFrm_FormClosed);
            this.Load += new System.EventHandler(this.SettingsFrm_Load);
            this.ImportSet.ResumeLayout(false);
            this.ImportSet.PerformLayout();
            this.ReportSet.ResumeLayout(false);
            this.ReportSet.PerformLayout();
            this.groupBox2.ResumeLayout(false);
            this.groupBox2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button OpenFileButton;
        private System.Windows.Forms.Label FileNameLabel;
        private System.Windows.Forms.CheckBox ExportCheck;
        private System.Windows.Forms.Label selectPortLabel;
        private System.Windows.Forms.ComboBox PortsSelectionBox;
        private System.Windows.Forms.Button SaveButton;
        private new System.Windows.Forms.Button CancelButton;
        private System.Windows.Forms.GroupBox ImportSet;
        private System.Windows.Forms.Label LanguageLabel;
        private System.Windows.Forms.CheckBox HeaderSpaceCheckBox;
        private System.Windows.Forms.TextBox HeaderSpace;
        private System.Windows.Forms.Label HeaderSpaceLabel;
        private System.Windows.Forms.GroupBox ReportSet;
        private System.Windows.Forms.Button DSpath;
        private System.Windows.Forms.Label addsign;
        public System.Windows.Forms.TextBox Digitalsignaturepath;
        public System.Windows.Forms.TextBox FilePath;
        private System.Windows.Forms.Label rpttype;
        public System.Windows.Forms.ComboBox ReportType;
        private System.Windows.Forms.Label rpttitle;
        private System.Windows.Forms.Label subtitle;
        private System.Windows.Forms.Label titlelab;
        private System.Windows.Forms.TextBox PrinterName;
        private System.Windows.Forms.Button ChoosePrinterButton;
        private System.Windows.Forms.Label PrinterNameLabel;
        private System.Windows.Forms.CheckBox AutoPrintCheck;
        public System.Windows.Forms.RadioButton PrintResultsRadioButton;
        public System.Windows.Forms.RadioButton MakePdfReportRadioButton;
        private System.Windows.Forms.GroupBox groupBox1;
        public System.Windows.Forms.TextBox subtitleval;
        public System.Windows.Forms.TextBox title;
        public System.Windows.Forms.CheckBox usedefault;
        public System.Windows.Forms.CheckBox pdfcheck;
        public System.Windows.Forms.ComboBox LanguageSelectionBox;
        private System.Windows.Forms.ComboBox DateFormatSet;
        private System.Windows.Forms.Label DateFormatLabel;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Label label1;
        public System.Windows.Forms.TextBox optionalval;
        private System.Windows.Forms.Label label2;
        public System.Windows.Forms.TextBox optionalval2;
        private System.Windows.Forms.CheckBox removeFooter;
    }
}