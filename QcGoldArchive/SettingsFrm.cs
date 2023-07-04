using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.Drawing.Printing;
using System.IO.Ports;
using System.IO;
using System.Xml;
using System.Text.RegularExpressions;
using System.Globalization;
using QcGoldArchive;
using System.Threading;


namespace QcGoldArchive
{


    public delegate void UpdateSettings();
    public delegate void UpdateIsSettingsOpenedFlag(object sender, EventArgs e);

    public partial class SettingsFrm : Form
    {
        Thread th;
        OldCsv_Support Ocv = new OldCsv_Support();
        Logclass Logs = new Logclass();
        private PrinterSettings printerSet = null;
        private PrintDialog printer = new PrintDialog();
        public bool changesFlag;  //if changes were made
        private bool saveButtonClicked = false;    //save was made by button
        private bool cancelButtonClicked = false;  //exit without saving
        public bool UI_flag = false;
        public string old_csvpath;

        public event UpdateSettings OnSettingsUpdate;
        public event UpdateIsSettingsOpenedFlag OnUpdateIsSettingsOpenedFlag;
        public string Dsign { get; set; }
        public string Dsp { get; set; }
        public string path;
        public int totalcols;
        // private enum Languages { English, German, Italian };

        LanguageManagement LM = LanguageManagement.CreateInstance();

        MainForm mf;
        public SettingsFrm(MainForm MyParent_)
        {
            mf = MyParent_;
            InitializeComponent();
            if (SerialPort.GetPortNames().Length > 0)
                PortsSelectionBox.Items.AddRange(SerialPort.GetPortNames());
            //LanguageSelectionBox.DataSource = Enum.GetValues(typeof(Languages));
            changesFlag = false;
            SetFields();
            LM.GetSubControlsData(this);
        }
    
        Export export;
        string[] settings;
        private string[] appSettings = XmlUtility.ReadSettings(Application.StartupPath + @"\settings.xml");
        private void SetFields()

        {
            ReportType.Items.Add(LM.Translate("BASIC_REPORT", MessagesLabelsEN.BASIC_REPORT));
            ReportType.Items.Add(LM.Translate("ADVANCED_REPORT", MessagesLabelsEN.ADVANCED_REPORT));

            int selectedPort = 0;
            bool portFound = false;
            string path = Application.StartupPath + @"\settings.xml";



            try
            {
                if ((File.Exists(path)) && (File.ReadAllText(path) != String.Empty))
                {

                    settings = XmlUtility.ReadSettings(path);

                    foreach (var item in SerialPort.GetPortNames())
                    {
                        if (item == settings[0])
                        {
                            portFound = true;
                            break;
                        }
                        selectedPort++;
                    }

                    if (portFound == true)
                        PortsSelectionBox.SelectedIndex = selectedPort;
                    else
                    {
                        if (SerialPort.GetPortNames().Length > 0)
                            PortsSelectionBox.SelectedIndex = 0;
                        else
                            PortsSelectionBox.Enabled = false;
                    }

                    if (settings[1] != "NONE") // not translated
                        PrinterName.Text = settings[1];
                    else
                    {
                        PrinterName.Enabled = false;
                        ChoosePrinterButton.Enabled = false;

                    }

                    FilePath.Text = settings[2];
                    Digitalsignaturepath.Text = settings[10];
                    optionalval.Text = settings[15];
                    optionalval2.Text = settings[16];

                    //title.Text = settings[12];
                    //subtitleval.Text = settings[13];
                    AutoPrintCheck.Checked = Boolean.Parse(settings[3]);
                    usedefault.Checked = Boolean.Parse(settings[4]);
                    //signaturecheck.Checked = Boolean.Parse(settings[12]);
                    ReportType.SelectedIndex = Int32.Parse(settings[11]);
                    title.Text = settings[12];
                    subtitleval.Text = settings[14];
                    pdfcheck.Checked = Boolean.Parse(settings[13]);
                    removeFooter.Checked = Boolean.Parse(settings[17]);


                    if (ReportType.SelectedIndex == 0)
                    {
                        //  signaturecheck.Checked = false;
                    }

                    DateFormatSet.SelectedIndex = Int32.Parse(settings[5]);

                    try
                    {
                        string[] filePaths = Directory.GetFiles(Application.StartupPath + "\\Language", "*.resources");
                        string tmpStr = "";
                        int i = 0;
                        string specName = "";

                        DataTable dtLang = new DataTable();
                        dtLang.Columns.Add("LangID");
                        dtLang.Columns.Add("LangName");

                        foreach (string filename in filePaths)
                        {

                            tmpStr = filename.Replace(Application.StartupPath + "\\Language\\VResources.", "");
                            tmpStr = tmpStr.Replace("resources", "");
                            tmpStr = tmpStr.Replace(".", "");

                            if (tmpStr != "")
                            {
                                specName = CultureInfo.CreateSpecificCulture(tmpStr).NativeName;//.DisplayName;
                                dtLang.Rows.Add(tmpStr, specName);
                            }
                            else
                            {
                                specName = CultureInfo.CreateSpecificCulture("en-US").NativeName;//.DisplayName;
                                dtLang.Rows.Add("en-US", specName);
                            }
                            i++;

                        }

                        dtLang.DefaultView.Sort = "LangName";
                        LanguageSelectionBox.DataSource = dtLang;
                        LanguageSelectionBox.DisplayMember = "LangName";
                        LanguageSelectionBox.ValueMember = "LangID";





                        //for (i = 0; i < LanguageSelectionBox.Items.Count; i++)
                        //{
                        //    //  DataRowView drv = (DataRowView)cmbLanguage.Items[i];

                        //    if (CurrentMDI.SetupData.SelectedLang == LanguageSelectionBox.Items[i].DataValue.ToString())
                        //    {
                        //        LanguageSelectionBox.SelectedIndex = i;
                        //        return;
                        //    }

                        //}
                        LanguageSelectionBox.SelectedValue = settings[6];


                    }
                    catch (Exception ex) { }

                    //LanguageSelectionBox.SelectedIndex = Int32.Parse(settings[6]);

                    // added

                    HeaderSpaceCheckBox.Checked = Boolean.Parse(settings[8]);
                    if (HeaderSpaceCheckBox.Checked == true)
                    {
                        HeaderSpace.Enabled = true;
                        HeaderSpaceLabel.Enabled = true;
                    }
                    else
                    {
                        HeaderSpace.Enabled = false;
                        HeaderSpaceLabel.Enabled = false;
                    }
                    HeaderSpace.Text = settings[9];
                    MakePdfReportRadioButton.Checked = Boolean.Parse(settings[7]);

                    if (MakePdfReportRadioButton.Checked == true)
                    {
                        PrintResultsRadioButton.Checked = false;
                        PrinterName.Enabled = false;
                        PrinterNameLabel.Enabled = false;
                        ChoosePrinterButton.Enabled = false;

                        //HeaderSpaceCheckBox.Enabled = true;
                        //HeaderCheck_Changed(sender, e);
                    }
                    else
                    {
                        PrintResultsRadioButton.Checked = true;
                        PrinterName.Enabled = true;
                        PrinterNameLabel.Enabled = true;
                        ChoosePrinterButton.Enabled = true;

                        HeaderSpaceCheckBox.Enabled = false;
                        HeaderSpace.Enabled = false;
                        HeaderSpaceLabel.Enabled = false;
                    }


                }
                else
                {
                    if (SerialPort.GetPortNames().Length > 0)
                        PortsSelectionBox.SelectedIndex = 0;
                    else
                        PortsSelectionBox.Enabled = false;

                    //if (PrinterSettings.InstalledPrinters.Count > 0)
                    //  PrinterName.Text = printer.PrinterSettings.PrinterName;
                    //else
                    //{
                    //    PrinterName.Enabled = false;
                    //    ChoosePrinterButton.Enabled = false;
                    //}
                    ReportType.SelectedIndex = 0;
                    DateFormatSet.SelectedIndex = 0;
                    LanguageSelectionBox.SelectedIndex = 0;

                    //added
                    HeaderSpace.Text = "";

                    MakePdfReportRadioButton.Checked = true;
                    HeaderSpaceCheckBox.Enabled = true;
                    PrintResultsRadioButton.Checked = false;
                    ChoosePrinterButton.Enabled = false;
                    PrinterName.Enabled = false;
                    PrinterNameLabel.Enabled = false;




                }
            }

            catch (Exception e)
            {
                /* MessageBox.Show(LM.Translate("SETTINGS_FILE_CORRUPTED_RELOAD", MessagesLabelsEN.SETTINGS_FILE_CORRUPTED_RELOAD), 
                  LM.Translate("SETTINGS", MessagesLabelsEN.SETTINGS));
                this.Close();*/
            }
            changesFlag = false;
        }

        private void ChoosePrinterButton_Click(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("Choose printer button clicked" + " \tChoosePrinterButton_Click in" + this.FindForm().Name, "Info");
            PrintDialog pd = new PrintDialog();
            //in order to open the dialog in win7 64bits
            pd.UseEXDialog = true;
            pd.ShowDialog();
            Logs.LogFile_Entries("Print dialogue box opened" + " \tChoosePrinterButton_Click in" + this.FindForm().Name, "Info");
            PrinterName.Text = pd.PrinterSettings.PrinterName;
            printerSet = pd.PrinterSettings;
            changesFlag = true;

        }
        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("open file button clicked" + " \tOpenFileButton_Click in" + this.FindForm().Name, "Info");
            OpenFileDialog dialog = new OpenFileDialog();
            dialog.Filter = "csv files (*.csv)|*.csv";
            dialog.InitialDirectory = Application.StartupPath;
            dialog.Title = LM.Translate("SELECT_FILE_TO_EXPORT", MessagesLabelsEN.SELECT_FILE_TO_EXPORT);
            dialog.DefaultExt = "csv";
            dialog.CheckFileExists = false;
            if (dialog.ShowDialog() == DialogResult.OK)
            {
                FilePath.Text = dialog.FileName;
                changesFlag = true;
                Logs.LogFile_Entries("File dialogue box opened" + " \tOpenFileButton_Click in" + this.FindForm().Name, "Info");
            }

        }
        private void ExportCheck_Changed(object sender, EventArgs e)
        {

            if (ExportCheck.Checked == true)
                if (IsExcelExists() == true)
                    ImportSet.Enabled = true;
                else
                {
                    ExportCheck.Checked = false;
                    MessageBox.Show(LM.Translate("EXCEL_NOT_EXIST", MessagesLabelsEN.EXCEL_NOT_EXIST), UniversalStrings.QC_GOLD_ARCHIVE);
                }
            else
                ImportSet.Enabled = false;

            changesFlag = true;

        }


        private void PrintCheck_Changed(object sender, EventArgs e)
        {
            changesFlag = true;
        }

        private void HeaderCheck_Changed(object sender, EventArgs e)
        {
            if (HeaderSpaceCheckBox.Checked == true)
            {
                HeaderSpace.Enabled = true;
                HeaderSpaceLabel.Enabled = true;

            }
            else
            {
                HeaderSpace.Enabled = false;
                HeaderSpaceLabel.Enabled = false;

            }
            changesFlag = true;
        }






        private void MakePdf_Changed(object sender, EventArgs e)
        {


            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1 && usedefault.Checked == false)
            {


                title.Enabled = true;
                subtitleval.Enabled = true;
                titlelab.Enabled = true;
                subtitle.Enabled = true;
                Logs.LogFile_Entries("Use default unchecked" + " \tMakePdf_Changed in" + this.FindForm().Name, "Info");

            }
            else
            {

                title.Enabled = false;
                subtitleval.Enabled = false;
                titlelab.Enabled = false;
                subtitle.Enabled = false;

            }



            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1)
            {
                Logs.LogFile_Entries("Advanced Report option choosen" + " \tMakePdf_Changed in" + this.FindForm().Name, "Info");
                usedefault.Enabled = true;
                rpttitle.Enabled = true;
                Digitalsignaturepath.Enabled = true;
                DSpath.Enabled = true;
                addsign.Enabled = true;

            }

            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 0)
            {
                Logs.LogFile_Entries("Basic Report option choosen" + " \tMakePdf_Changed in" + this.FindForm().Name, "Info");
                title.Enabled = false;
                subtitleval.Enabled = false;
                titlelab.Enabled = false;
                subtitle.Enabled = false;
                pdfcheck.Enabled = true;

            }

            if (MakePdfReportRadioButton.Checked == true)
            {
                //PrintResultsRadioButton.Checked = false;

                Logs.LogFile_Entries("Create pdf report option choosen" + " \tMakePdf_Changed in" + this.FindForm().Name, "Info");
                HeaderSpaceCheckBox.Enabled = true;
                HeaderSpace.Enabled = true;
                HeaderSpaceLabel.Enabled = true;
                PrinterName.Enabled = false;
                PrinterNameLabel.Enabled = false;
                ChoosePrinterButton.Enabled = false;
                rpttype.Enabled = true;
                ReportType.Enabled = true;
                AutoPrintCheck.Enabled = false;
                // HeaderSpaceCheckBox.Enabled = true;

                HeaderCheck_Changed(sender, e);



            }
            else
            {
                //PrintResultsRadioButton.Checked = true;

                titlelab.Enabled = false;
                title.Enabled = false;
                subtitleval.Enabled = false;
                subtitle.Enabled = false;
                addsign.Enabled = false;
                Digitalsignaturepath.Enabled = false;
                DSpath.Enabled = false;
                HeaderSpaceCheckBox.Enabled = false;
                HeaderSpace.Enabled = false;
                HeaderSpaceLabel.Enabled = false;
                rpttype.Enabled = false;
                ReportType.Enabled = false;
                rpttitle.Enabled = false;
                usedefault.Enabled = false;
                AutoPrintCheck.Enabled = true;
                //pdfcheck.Enabled = true;


                PrinterName.Enabled = true;
                PrinterNameLabel.Enabled = true;
                ChoosePrinterButton.Enabled = true;


            }
            UI_flag = true;
            changesFlag = true;
            ReportType.Enabled = false;
        }
        public void linecount()
        {
            string path1 = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path1);
            char seperator = mf.dsChar; // ';';

            try
            {

                path = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";
                var lines = File.ReadAllLines(path);
                totalcols = lines[0].Split(seperator).Length;


            }
            catch (Exception ex)
            {

            }

        }

        private void SaveChanges(object sender, FormClosingEventArgs e)
        {
            //checkOldArchive(true);
            Logs.LogFile_Entries("settings form close button clicked" + " \tSaveChanges in" + this.FindForm().Name, "Info");
            if ((changesFlag == true) && (saveButtonClicked == false) && (cancelButtonClicked == false))

            {

                if (((MessageBox.Show(this, LM.Translate("SAVE_CHANGES", MessagesLabelsEN.SAVE_CHANGES), LM.Translate("CONFIRM_CHANGES", MessagesLabelsEN.CONFIRM_CHANGES), MessageBoxButtons.YesNo)) == DialogResult.Yes))
                {

                    if ((FilePath.Text == String.Empty))
                    {
                        MessageBox.Show(this, LM.Translate("SPECIFY_FILE_NAME", MessagesLabelsEN.SPECIFY_FILE_NAME), LM.Translate("EXPORT_FILENAME_MISSING", MessagesLabelsEN.EXPORT_FILENAME_MISSING));
                        Logs.LogFile_Entries(MessagesLabelsEN.SPECIFY_FILE_NAME + " \tSaveChanges in" + this.FindForm().Name, "Error");
                        FilePath.Focus();
                        e.Cancel = true;
                    }
                    else if (IsValidPath(FilePath.Text) == false && FilePath.Text != string.Empty)
                    {
                        MessageBox.Show(this, LM.Translate("INVALID_FILEPATH", MessagesLabelsEN.INVALID_FILEPATH), LM.Translate("INVALID_FILENAME", MessagesLabelsEN.INVALID_FILENAME));
                        Logs.LogFile_Entries(MessagesLabelsEN.INVALID_FILEPATH + " \tSaveChanges in" + this.FindForm().Name, "Error");
                        FilePath.Focus();
                        e.Cancel = true;
                    }
                    else if (ReportType.SelectedIndex == 1 && IsValidPath(Digitalsignaturepath.Text) == false && Digitalsignaturepath.Text != String.Empty)
                    {
                        MessageBox.Show(this, LM.Translate("INVALID_FILEPATH", MessagesLabelsEN.INVALID_FILEPATH), LM.Translate("INVALID_FILENAME", MessagesLabelsEN.INVALID_FILENAME));
                        Logs.LogFile_Entries(MessagesLabelsEN.INVALID_FILEPATH + " \tSaveChanges in" + this.FindForm().Name, "Error");
                        Digitalsignaturepath.Focus();
                        e.Cancel = true;

                    }
                    else if (string.IsNullOrEmpty(PrinterName.Text) && PrintResultsRadioButton.Checked == true)
                    {
                        Logs.LogFile_Entries(MessagesLabelsEN.Print_Warning + " \tSaveChanges in" + this.FindForm().Name, "Error");
                        if (MessageBox.Show(LM.Translate("select_printer_Name", MessagesLabelsEN.Print_Warning), LM.Translate("Select_printer", MessagesLabelsEN.Print_Error_Warning), MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                        {
                            PrinterName.Focus();
                            e.Cancel = true;
                        }
                    }
                    else if (string.IsNullOrEmpty(HeaderSpace.Text) && HeaderSpaceCheckBox.Checked == true)
                    {
                        Logs.LogFile_Entries(MessagesLabelsEN.HeaderSpace_Warning + " \tSaveChanges in" + this.FindForm().Name, "Error");
                        if (MessageBox.Show(LM.Translate("Enter_Header_value", MessagesLabelsEN.HeaderSpace_Warning), LM.Translate("Enter_value", MessagesLabelsEN.HeaderSpace_Error_Warning), MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                        {
                            HeaderSpace.Focus();
                            e.Cancel = true;
                        }
                    }

                    else if (string.IsNullOrEmpty(title.Text) && MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1 && usedefault.Checked == false)
                    {
                        Logs.LogFile_Entries(MessagesLabelsEN.Title_Warning + " \tSaveChanges in" + this.FindForm().Name, "Error");
                        if (MessageBox.Show(LM.Translate("Please_enter_title", MessagesLabelsEN.Title_Warning), LM.Translate("Enter_value", MessagesLabelsEN.Title_Error_Warning), MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                        {
                            title.Focus();
                            e.Cancel = true;
                        }
                    }

                    else if (checkOldArchive(true))
                    {

                        this.FilePath.Focus();
                        mf.set_check = false;
                        e.Cancel = true;
                    }
                    else
                    {
                        if (FilePath.Text.EndsWith(".csv") == false)   //if no .xls was added manually
                            FilePath.Text += ".csv";
                        saveButtonClicked = true;
                        SaveSettings();
                        Logs.LogFile_Entries("Settings.xml updated" + " \tSaveChanges in" + this.FindForm().Name, "Info");




                    }

                }





                //if (!checkOldArchive(true))
                //{

                //    this.FilePath.Focus();
                //    mf.set_check = false;
                //    e.Cancel = true;
                //}
                //else
                //{
                //    //
                //    saveButtonClicked = true;
                //    SaveSettings();
                //    this.Close();
                //    //  Application.Restart();
                //    refreshfrm();

                //}








                refreshfrm();

            }

            try
            {
                OnUpdateIsSettingsOpenedFlag(this, null);
            }
            catch { }

        }


        public void SaveSettings()
        {
            string path = Application.StartupPath + @"\settings.xml";
           
            string[] settings = new string[18];

            try
            {
                settings[0] = PortsSelectionBox.Items[PortsSelectionBox.SelectedIndex].ToString();
            }
            catch
            {
                settings[0] = "-1";
            }
           
            settings[1] = PrinterName.Text;            
            settings[2] = FilePath.Text;
            settings[3] = AutoPrintCheck.Checked.ToString();
            settings[4] = usedefault.Checked.ToString();
            settings[5] = DateFormatSet.SelectedIndex.ToString();
         //   settings[6] = LanguageSelectionBox.SelectedValue.ToString(); //SelectedIndex.ToString();
            settings[7] = MakePdfReportRadioButton.Checked.ToString();
            settings[8] = HeaderSpaceCheckBox.Checked.ToString();
            settings[9] = HeaderSpace.Text;
            settings[10] = Digitalsignaturepath.Text;
            settings[11] = ReportType.SelectedIndex.ToString();
            settings[12] = title.Text;
            settings[14] = subtitleval.Text;
            settings[13] = pdfcheck.Checked.ToString();
            settings[15] = optionalval.Text;
            settings[16] = optionalval2.Text;
            settings[17] = removeFooter.Checked.ToString();

            try
            {
                Logs.LogFile_Entries("Selected Port number: " + PortsSelectionBox.Items[PortsSelectionBox.SelectedIndex].ToString() + "" + " Received in" + this.FindForm().Name, "Info");
            }
            catch { }
        //    Logs.LogFile_Entries("Current Language: " + LanguageSelectionBox.SelectedValue.ToString(), "Info");
            Logs.LogFile_Entries("Create pdf Report button Status:" + MakePdfReportRadioButton.Checked.ToString(), "Info");
            Logs.LogFile_Entries("Selected Printer Name:" + PrinterName.Text, "Info");
            Logs.LogFile_Entries("Header space value: " + HeaderSpace.Text, "Info");
            Logs.LogFile_Entries("Digital Signature path:" + Digitalsignaturepath.Text, "Info");
            Logs.LogFile_Entries("Selected Report Type:" + ReportType.Text, "Info");
            Logs.LogFile_Entries("Entered Title Text:" + title.Text, "Info");
            Logs.LogFile_Entries("Entered Subtitle Text:" + subtitleval.Text, "Info");
            Logs.LogFile_Entries("Selected file path:" + FilePath.Text, "Info");
            Logs.LogFile_Entries("Auto print(For strip report) Status :" + bool.Parse(AutoPrintCheck.Checked.ToString()), "Info");
            Logs.LogFile_Entries("Use Default Status: " + bool.Parse(usedefault.Checked.ToString()), "Info");
            Logs.LogFile_Entries("Header space Status: " + bool.Parse(HeaderSpaceCheckBox.Checked.ToString()), "Info");
            Logs.LogFile_Entries("Auto print(For PDF report) Status :" + bool.Parse(pdfcheck.Checked.ToString()), "Info");







            try
            {
                XmlUtility.WriteSettings(settings, path);
                OnSettingsUpdate();
            }

            catch (Exception e)
            {
                //MessageBox.Show(this, e.Message, LM.Translate("SETTINGS_FILE_ERROR", MessagesLabelsEN.SETTINGS_FILE_ERROR));
            }

            //Languages currentLang = (Languages)LanguageSelectionBox.SelectedItem;
            //switch (currentLang)
            //{
            //    case Languages.English:
            //        LM.UpdateLanguage(UniversalStrings.ENG_US);
            //        break;
            //    case Languages.German:
            //        LM.UpdateLanguage(UniversalStrings.GER_GERMANY);
            //        break;
            //    case Languages.Italian:
            //        LM.UpdateLanguage(UniversalStrings.IT_ITALIA);
            //        break;
            //    default:
            //        LM.UpdateLanguage(UniversalStrings.ENG_US);
            //        break;
            //}
            LM.UpdateLanguage(settings[6]);

        }

        public void refreshfrm()
        {
            
           
            mf.comments.Visible = false;
            mf.Avg_desg.Visible = false;
            mf.Avg_frclbl.Visible = false;
            mf.Avg_liqlbl.Visible = false;
            mf.Avg_tesname.Visible = false;
            mf.Avg_desg.Visible = false;
            mf.Avg_patname.Visible = false;
            mf.Avg_desg.Text = "";
            mf.Avg_frclbl.Text = "";
            mf.Avg_liqlbl.Text = "";
            mf.Avg_tesname.Text = "";
            mf.Avg_patname.Text = "";
           
            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);


            if (Boolean.Parse(appSettings[7]) == false)
            {
                Logs.LogFile_Entries("Print result strip UI displayed" + " \trefereshfrm in" + this.FindForm().Name, "Info");
                mf.resultstrip.Text = LM.Translate("PRINT_RESULTS_STRIP", MessagesLabelsEN.EXPORT_PDF);
                mf.PrintButton.Image = System.Drawing.Image.FromFile(Application.StartupPath + "\\assets\\report1.png");
                pdfcheck.Enabled = false;
                mf.ArchiveButton.Enabled = false;
                // mf.pictureBox1.Location = new Point(120, 56);

                //pictureBox1.Image = imageList1.Images[2];
                foreach (Control c in mf.Controls)
                {
                    if (c is TextBox || c is Button || c is Label || c is RichTextBox)
                    {
                        c.Visible = false;
                        mf.resultstrip.Visible = true;
                        mf.Plot.Visible = true;
                    }

                }

            }
            else
            {
                mf.ArchiveButton.Enabled = true;
                foreach (Control c in mf.Controls)
                {
                    if (c is TextBox || c is Button || c is Label)
                    {
                        c.Visible = true;
                        mf.richTextBox1.Visible = true;
                        mf.resultstrip.Visible = false;
                    }

                }

            }
            if (appSettings[6] == "zh-CN")
            {
                mf.devicesntxt.Location = new Point(80, 722);
                mf.devicesn.Location = new Point(150, 722);
            }
            else
            {
                mf.devicesntxt.Location = new Point(168, 816);
                mf.devicesn.Location = new Point(256, 816);
            }
            if (appSettings[6] == "de-De" || appSettings[6] == "fr-FR")
            {
                mf.testername.Location = new Point(178, 478);
            }
            else
            {
                mf.testername.Location = new Point(153, 479);
            }
            if (Boolean.Parse(appSettings[7]) == false)
            {
                mf.SaveButton.Enabled = false;
            }
            if (int.Parse(appSettings[11]) == 0)
            {
                mf.SaveButton.Enabled = false;
            }

            //if (int.Parse(appSettings[11]) == 1 && Boolean.Parse(appSettings[7]) == true && mf.check==false && !string.IsNullOrEmpty(mf.patname.Text))
            //{
            //    mf.SaveButton.Enabled = true;
            //}

            if (Boolean.Parse(appSettings[7]) == true)
            {
                mf.Plot.Visible = false;
                mf.refbydr2.Visible = false;
                mf.refbydr3.Visible = false;
                mf.PrintButton.Text = "Export pdf";
                mf.PrintButton.Image = System.Drawing.Image.FromFile(Application.StartupPath + "\\assets\\report.png");

            }


            if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 1)
            {
                Logs.LogFile_Entries("Advanced Report UI displayed" + " \trefereshfrm in" + this.FindForm().Name, "Info");
                mf.Resultinfo.Location = new Point(26, 528);
                mf.label42.Location = new Point(25, 547);
                mf.conc.Location = new Point(185, 555);
                mf.label4.Location = new Point(24, 555);
                mf.label5.Location = new Point(25, 583);
                mf.totmot.Location = new Point(185, 583);
                mf.rapidProglable.Location = new Point(25, 610);
                mf.progmotility.Location = new Point(185, 610);


                mf.label21.Location = new Point(25, 637);
                mf.nonprgmot.Location = new Point(185, 637);

                mf.label20.Location = new Point(26, 662);
                mf.immotility.Location = new Point(185, 663);
                mf.label19.Location = new Point(26, 689);
                mf.morph.Location = new Point(185, 689);

                mf.label27.Location = new Point(284, 560);
                mf.msc.Location = new Point(455, 560);

                mf.rapidpmscLable.Location = new Point(284, 585);
                mf.pmsc.Location = new Point(455, 585);
                mf.label25.Location = new Point(284, 609);
                mf.fsc.Location = new Point(455, 609);

                mf.label24.Location = new Point(284, 635);
                mf.velocity.Location = new Point(455, 635);
                mf.label23.Location = new Point(284, 662);
                mf.smi.Location = new Point(455, 662);

                mf.label31.Location = new Point(528, 557);
                mf.sperm.Location = new Point(705, 557);
                mf.label30.Location = new Point(528, 582);
                mf.motsperm.Location = new Point(705, 582);

                mf.label29.Location = new Point(528, 609);
                mf.progsperm.Location = new Point(705, 609);
                mf.label28.Location = new Point(528, 635);
                mf.funcsperm.Location = new Point(705, 635);
                mf.label22.Location = new Point(528, 662);
                mf.label2.Location = new Point(705, 662);
                //mf.pictureBox1.Location = new Point(138, 57);

            }



            if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 0)
            {

                mf.rptname.Text = LM.Translate("BASIC_REPORT", MessagesLabelsEN.BASIC_REPORT);
                Logs.LogFile_Entries("Basic Report UI displayed" + " \trefereshfrm in" + this.FindForm().Name, "Info");
                mf.pictureBox1.Image = mf.imageList1.Images[1];
                // mf.pictureBox1.Location = new Point(125, 58);
                mf.fructoselbl.Visible = false;
                mf.fructose.Visible = false;
                mf.liquifiction.Visible = false;
                mf.liquilbl.Visible = false;
                mf.refbydr.Visible = false;
                mf.refbtn.Visible = false;
                mf.label40.Visible = false;
                mf.refbydrlbl.Visible = false;
                mf.patname.Visible = false;
                mf.patnamelbl.Visible = false;
                mf.fructosetxt.Visible = false;
                mf.liquifictiontxt.Visible = false;
                mf.refbydr2.Visible = false;
                mf.refbydr3.Visible = false;
                mf.Testerinfo.Visible = false;
                mf.richTextBox1.Visible = false;
                mf.optionallab.Visible = false;
                mf.textBox1.Visible = false;
                mf.textBox2.Visible = false;
                mf.label15.Visible = false;
                mf.label16.Visible = false;
                mf.label1.Visible = false;

                // Plot.Visible = false;
                //  mf.tesname.Visible = false;
                mf.testernamelbl.Visible = false;
                mf.tesdesg.Visible = false;
                mf.testername.Visible = false;
                mf.testerdesg.Visible = false;
                mf.Resultinfo.Location = new Point(27, 415);
                mf.label42.Location = new Point(27, 440);
                mf.conc.Location = new Point(190, 455);
                mf.label4.Location = new Point(27, 455);
                mf.label5.Location = new Point(27, 480);
                mf.totmot.Location = new Point(190, 480);
                mf.rapidProglable.Location = new Point(27, 505);
                mf.progmotility.Location = new Point(190, 505);


                mf.label21.Location = new Point(27, 530);
                mf.nonprgmot.Location = new Point(190, 530);

                mf.label20.Location = new Point(27, 553);
                mf.immotility.Location = new Point(190, 553);
                mf.label19.Location = new Point(27, 575);
                mf.morph.Location = new Point(190, 575);

                mf.label27.Location = new Point(284, 455);
                mf.msc.Location = new Point(455, 455);

                mf.rapidpmscLable.Location = new Point(284, 480);
                mf.pmsc.Location = new Point(455, 480);
                mf.label25.Location = new Point(284, 505);
                mf.fsc.Location = new Point(455, 505);

                mf.label24.Location = new Point(284, 530);
                mf.velocity.Location = new Point(455, 530);
                mf.label23.Location = new Point(284, 553);
                mf.smi.Location = new Point(455, 553);

                mf.label31.Location = new Point(510, 455);
                mf.sperm.Location = new Point(697, 455);
                mf.label30.Location = new Point(510, 480);
                mf.motsperm.Location = new Point(697, 480);

                mf.label29.Location = new Point(510, 505);
                mf.progsperm.Location = new Point(697, 505);
                mf.label28.Location = new Point(510, 530);
                mf.funcsperm.Location = new Point(697, 530);
                mf.label22.Location = new Point(510, 555);
                mf.label2.Location = new Point(697, 555);


            }

            else if (Boolean.Parse(appSettings[7]) == true && ReportType.SelectedIndex == 1)
            {

                mf.rptname.Text = LM.Translate("ADVANCED_REPORT", MessagesLabelsEN.ADVANCED_REPORT);

                // mf.pictureBox1.Location = new Point(155, 58);
                mf.pictureBox1.Image = mf.imageList1.Images[0];
                mf.Plot.Visible = false;


            }
            else if (Boolean.Parse(appSettings[7]) == false && mf.resultstrip.Enabled == true)
            {
                mf.pictureBox1.Image = mf.imageList1.Images[2];
                //mf.pictureBox1.Location = new Point(118, 57);
            }
            if (int.Parse(appSettings[11]) == 1 && string.IsNullOrEmpty(mf.Testdate.Text))
            {
                //mf.patname.Enabled = false;
                //mf.fructosetxt.Enabled = false;
                //mf.liquifictiontxt.Enabled = false;
                //mf.testername.Enabled = false;
                //mf.testerdesg.Enabled = false;
                //mf.refbydr.Enabled = false;
                //mf.refbydr2.Enabled = false;
                //mf.refbydr3.Enabled = false;
            }

            if (ReportType.SelectedIndex == 1 && MakePdfReportRadioButton.Checked == true)
            {
                if (mf.refbydr2.Text != "")
                {
                    mf.refbydr2.Visible = true;
                }
                if (mf.refbydr3.Text != "")
                {
                    mf.refbydr3.Visible = true;
                }
            }
            //if (!string.IsNullOrEmpty(mf.type.Text))
            //{

            //    if (mf.type.Text == LM.Translate("FRESH", MessagesLabelsEN.FRESH))
            //    {
            //        mf.isFresh = true;
            //    }
            //   else if (mf.type.Text == LM.Translate("FROZEN", MessagesLabelsEN.FROZEN))
            //    {
            //        mf.isFrozen = true;
            //    }
            //   else if (mf.type.Text == LM.Translate("WASHED", MessagesLabelsEN.WASHED))
            //    {
            //        mf.isWashed = true;
            //    }
            //    
            //}
            mf.translateType1();
            //if (!string.IsNullOrEmpty(mf.type.Text) )
            //{
            //    string test = LM.Translate(mf.type.Text, mf.type.Text);
            //    mf.type.Text = test;

            //}
            if (!mf.checkBox2.Checked)
            {
                mf.patidField.Visible = false;
                mf.testdateField1.Visible = false;
                mf.collecteddateField.Visible = false;
                mf.receivedDateField.Visible = false;
                mf.absField.Visible = false;
                mf.accessionField.Visible = false;
                mf.volField.Visible = false;
                mf.phField.Visible = false;
                mf.wbcField.Visible = false;
                mf.typeField.Visible = false;
            }
        }

        private bool checkOldArchive(bool clickedX)
        {
           
            try
            {
                var path = this.FilePath.Text.Substring(0, this.FilePath.Text.Length - 3) + "csv";
                if (File.Exists(path))
                {
                    var lines = File.ReadAllLines(path);
                    char LineStr;
                    if (mf.dsChar == char.Parse(","))
                        LineStr = ';';
                    else
                        LineStr = ',';

                    totalcols = lines[0].Split(LineStr).Length;
                   
                    //if (totalcols <=30)
                    //{
                       
                    //    Logs.LogFile_Entries(MessagesLabelsEN.Old_File_MESSAGE + " \tcheckOldArchive in" + this.FindForm().Name, "Error");
                    //    if (MessageBox.Show(LM.Translate("You_choose_old_Record", MessagesLabelsEN.Old_File_MESSAGE), LM.Translate("Warning", MessagesLabelsEN.Warning), MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                    //    {
                          
                    //        SaveSettings();
                    //        Ocv.old_csv(FilePath.Text);
                            

                    //        return false;

                    //    }
                    //    else
                    //    {

                    //        //sf.changesFlag = true;
                    //        FilePath.Text = appSettings[2];
                    //        return true;

                    //    }
                    //}
                    if (totalcols <= 42)
                    {
                        //Update Indian CSV
                        Logs.LogFile_Entries(MessagesLabelsEN.Old_File_MESSAGE + " \tcheckOldArchive in" + this.FindForm().Name, "Error");
                        if (MessageBox.Show(LM.Translate("You_choose_old_Record", MessagesLabelsEN.Old_File_MESSAGE), LM.Translate("Warning", MessagesLabelsEN.Warning), MessageBoxButtons.OKCancel, MessageBoxIcon.Warning) == DialogResult.OK)
                        {

                            SaveSettings();
                            Ocv.IN_old_csv(FilePath.Text);


                            return false;

                        }
                        else
                        {

                            //sf.changesFlag = true;
                            FilePath.Text = appSettings[2];
                            return true;

                        }

                    }
                    return false;
                }
                return false;
            }
            catch (Exception ex)
            {
                try
                {
                    if (!string.IsNullOrEmpty(this.FilePath.Text))
                    {
                        var path = this.FilePath.Text.Substring(0, this.FilePath.Text.Length - 3) + "csv";
                        bool cancelled = false;
                        //while (IsFileOpen(filename) && cancelled == false)
                        while (mf.CanAccessFile(path) == false && cancelled == false)
                        {
                            this.Invoke(new Action(() =>
                            {
                                Logs.LogFile_Entries(MessagesLabelsEN.CSV_OPEN_ERROR + " \tcheckOldArchive in" + this.FindForm().Name, "Error");
                                if (MessageBox.Show(this, LM.Translate("csv_file_not_open", MessagesLabelsEN.CSV_OPEN_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.OK) == DialogResult.Cancel)
                                {
                                    this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("", MessagesLabelsEN.CSV_OPEN_ERROR), UniversalStrings.QC_GOLD_ARCHIVE); }));
                                    cancelled = true;
                                };
                            }));
                        }

                        if (!cancelled)
                        {
                            //checkOldArchive(true);
                        }
                        return false;
                    }
                    return false;
                }
                catch { return false; }
            }
        }

        private void SaveButton_Click(object sender, EventArgs e)
        {

            Logs.LogFile_Entries("Save button clicked" + "in" + this.FindForm().Name, "Info");
            // Dsp = Digitalsignaturepath.Text;
            //mf.check2();
            mf.counter = 0;


            if (!checkOldArchive(true))
            {
                // this.SaveSettingsSettings();
                mf.set_check = true;
            }



            else
            {
                this.FilePath.Focus();
                mf.set_check = false;
            }

            bool found = false;
            if (mf.DEL_csv_records == true || mf.archive_data == true)
            {
                if ((mf.SaveButton.Enabled == true && UI_flag == true) || (mf.SaveButton.Enabled == true && ReportType.SelectedIndex == 0))
                {
                    Logs.LogFile_Entries(MessagesLabelsEN.SAVE_DATA + " \tSaveButton_Click in" + this.FindForm().Name, "Error");
                    if (((MessageBox.Show(this, LM.Translate("SAVE_DATA", MessagesLabelsEN.SAVE_DATA), LM.Translate("CONFIRM_CHANGES", MessagesLabelsEN.CONFIRM_CHANGES), MessageBoxButtons.YesNo)) == DialogResult.Yes))
                    {
                        mf.updatecsv();


                    }
                    //else
                    //{
                    //    ArchiveFrm archive_frm = new ArchiveFrm(mf);
                    //    archive_frm.display_csv_data();
                    //}

                }
            }

            if (string.IsNullOrEmpty(PrinterName.Text) && PrintResultsRadioButton.Checked == true)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Print_Warning + " \tSaveButton_Click in" + this.FindForm().Name, "Error");
                if (MessageBox.Show(LM.Translate("select_printer_Name", MessagesLabelsEN.Print_Warning), LM.Translate("Select_printer", MessagesLabelsEN.Print_Error_Warning), MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    PrinterName.Focus();

                }
            }
            else if ((FilePath.Text == String.Empty))
            {
                Logs.LogFile_Entries(MessagesLabelsEN.SPECIFY_FILE_NAME + " \tSaveButton_Click in" + this.FindForm().Name, "Error");
                MessageBox.Show(this, LM.Translate("SPECIFY_FILE_NAME", MessagesLabelsEN.SPECIFY_FILE_NAME), LM.Translate("EXPORT_FILENAME_MISSING", MessagesLabelsEN.EXPORT_FILENAME_MISSING));
                FilePath.Focus();
            }
            else if (IsValidPath(FilePath.Text) == false)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.INVALID_FILEPATH + " \tSaveButton_Click in" + this.FindForm().Name, "Error");
                MessageBox.Show(this, LM.Translate("INVALID_FILEPATH", MessagesLabelsEN.INVALID_FILEPATH), LM.Translate("INVALID_FILENAME", MessagesLabelsEN.INVALID_FILENAME));
                FilePath.Focus();
            }
            else if (string.IsNullOrEmpty(HeaderSpace.Text) && HeaderSpaceCheckBox.Checked == true)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.HeaderSpace_Warning + " \tSaveButton_Click in" + this.FindForm().Name, "Error");
                if (MessageBox.Show(LM.Translate("Enter_Header_value", MessagesLabelsEN.HeaderSpace_Warning), LM.Translate("Enter_value", MessagesLabelsEN.HeaderSpace_Error_Warning), MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    HeaderSpace.Focus();

                }
            }
            else if (string.IsNullOrEmpty(title.Text) && MakePdfReportRadioButton.Checked == true && usedefault.Checked == false && ReportType.SelectedIndex==1)
            {

                Logs.LogFile_Entries(MessagesLabelsEN.Title_Warning + " \tSaveButton_Click in" + this.FindForm().Name, "Error");
                if (MessageBox.Show(LM.Translate("Please_enter_title", MessagesLabelsEN.Title_Warning), LM.Translate("Enter_value", MessagesLabelsEN.Title_Error_Warning), MessageBoxButtons.OK, MessageBoxIcon.Information) == DialogResult.OK)
                {
                    title.Focus();

                }
            }


            else
            {

                if (FilePath.Text.EndsWith(".csv") == false && (FilePath.Text != String.Empty))
                {//if no .xls was added manually
                    FilePath.Text += ".csv";
                }

                //if(ReportType.SelectedIndex!=int.Parse(appSettings[11]) || PrintResultsRadioButton.Checked !=Boolean.Parse(appSettings[7]))
                //{


                
                
                saveButtonClicked = true;
                if (mf.set_check == true)
                {
                    SaveSettings();
                    this.Close();
                }
                else
                {
                    FilePath.Focus();
                }


                

                //if (!found)
                //{
                //    MainForm aForm = new MainForm();
                //    aForm.Name = "MainForm";
                //    aForm.Refresh();


                //}

                //Application.DoEvents();

                // / mf.ShowDialog();



                /* th = new Thread(opennewfrm);
                 th.SetApartmentState(ApartmentState.STA);
                 th.Start();
             */
            }
            changesFlag = true;
            saveButtonClicked = false;
            refreshfrm();
            //if (totalcols < 42)
            //{
            //    Ocv.old_csv(appSettings[2]);
            //}
            Logs.LogFile_Entries("Settings form closed" + " \tSaveButton_Click in" + this.FindForm().Name, "Info");

        }


        private void opennewfrm(object obj)
        {
            Application.Run(new MainForm());
        }
        private void CancelButton_Click(object sender, EventArgs e)
        {
            cancelButtonClicked = true;
            this.Close();
            Logs.LogFile_Entries("Settings form closed" + " \tCancelButton_Click in" + this.FindForm().Name, "Info");
        }

        private bool IsValidPath(string path)
        {
            Regex r = new Regex(@"^(([a-zA-Z]\:)|(\\))(\\{1}|((\\{1})[^\\]([^/:*?<>""|]*))+)$");
            return r.IsMatch(path);
        }

        private void FilePath_Changed(object sender, EventArgs e)
        {
            changesFlag = true;
            
        }

        private bool IsExcelExists()
        {
            Microsoft.Win32.RegistryKey key = Microsoft.Win32.Registry.ClassesRoot;
            Microsoft.Win32.RegistryKey excelKey = key.OpenSubKey("Excel.Application");
            if (excelKey == null)
                return false;
            else
                return true;

        }

        private void Port_Changed(object sender, EventArgs e)
        {


            changesFlag = true;
        }

        private void Languange_Changed(object sender, EventArgs e)
        {
            changesFlag = true;
        }

        private void DateFormat_Changed(object sender, EventArgs e)
        {
            changesFlag = true;
        }

        private void HeaderSpace_Changed(object sender, EventArgs e)
        {
            try
            {
                int val = int.Parse(HeaderSpace.Text);
                if (val > 30)
                {
                    MessageBox.Show(LM.Translate("Max_Value", MessagesLabelsEN.Max_Exceed));
                    Logs.LogFile_Entries(MessagesLabelsEN.Max_Exceed + " \tHeaderSpace_Changed in" + this.FindForm().Name, "Info");

                    int num = 30;
                    HeaderSpace.Text = num.ToString();

                }
            }
            catch (Exception ex)
            {

            }
            changesFlag = true;
        }
        private void Digitalsignaturepath_TextChanged(object sender, EventArgs e)
        {


            changesFlag = true;
        }

        private void ReportType_SelectedIndexChanged(object sender, EventArgs e)
        {

            //saveButtonClicked = true;

            if (HeaderSpaceCheckBox.Checked == true)
            {
                HeaderSpace.Enabled = true;
                Logs.LogFile_Entries("Header space check box checked" + " \tReportType_SelectedIndexChanged in" + this.FindForm().Name, "Info");
            }
            else
            {
                HeaderSpace.Enabled = false;
                Logs.LogFile_Entries("Header space check box unchecked" + " \tReportType_SelectedIndexChanged in" + this.FindForm().Name, "Info");
            }

            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1 && usedefault.Checked == false)
            {
                titlelab.Enabled = true;
                title.Enabled = true;
                subtitleval.Enabled = true;
                subtitle.Enabled = true;

            }
            if (!MakePdfReportRadioButton.Checked)
            {
                pdfcheck.Enabled = false;
            }
            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1)
            {

                addsign.Enabled = true;
                Digitalsignaturepath.Enabled = true;
                DSpath.Enabled = true;
                HeaderSpaceCheckBox.Enabled = true;
                //HeaderSpace.Enabled = true;
                HeaderSpaceLabel.Enabled = true;
                rpttitle.Enabled = true;
                usedefault.Enabled = true;
                AutoPrintCheck.Enabled = false;
                pdfcheck.Enabled = false;



            }

            else if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 0)
            {
                titlelab.Enabled = false;
                title.Enabled = false;
                subtitleval.Enabled = false;
                subtitle.Enabled = false;
                addsign.Enabled = false;
                Digitalsignaturepath.Enabled = false;
                DSpath.Enabled = false;
                HeaderSpaceCheckBox.Enabled = true;
                //HeaderSpace.Enabled = true;
                HeaderSpaceLabel.Enabled = true;

                rpttitle.Enabled = false;
                usedefault.Enabled = false;
                pdfcheck.Enabled = true;
                AutoPrintCheck.Enabled = false;


            }

            else
            {
                titlelab.Enabled = false;
                title.Enabled = false;
                subtitleval.Enabled = false;
                subtitle.Enabled = false;
                addsign.Enabled = false;
                Digitalsignaturepath.Enabled = false;
                DSpath.Enabled = false;
                rpttitle.Enabled = false;
                usedefault.Enabled = false;
            }

            changesFlag = true;

        }



        private void SettingsFrm_Load(object sender, EventArgs e)
        {
            
            PrintResultsRadioButton.Text = LM.Translate("PRINT_RESULTS_STRIP", MessagesLabelsEN.EXPORT_PDF);
            MainForm mf = new MainForm();
            //mf. 
            if (appSettings[6] == "en-US")
            {
                HeaderSpace.Location = new Point(163, 249);
                HeaderSpaceLabel.Location = new Point(219, 253);
            }



            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1 && usedefault.Checked == true)
            {

                title.Enabled = false;
                titlelab.Enabled = false;
                subtitle.Enabled = false;
                subtitleval.Enabled = false;
            }
            else
            {

                title.Enabled = true;
                titlelab.Enabled = true;
                subtitle.Enabled = true;
                subtitleval.Enabled = true;
            }
            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 0)
            {
                title.Enabled = false;
                titlelab.Enabled = false;
                subtitle.Enabled = false;
                subtitleval.Enabled = false;
            }
            if (PrintResultsRadioButton.Checked == true)
            {
                subtitle.Enabled = false;
                subtitleval.Enabled = false;
                title.Enabled = false;
                titlelab.Enabled = false;
            }
            /*
            if (ReportType.Text == "Normal Report")
            {
                signaturecheck.Checked = false;
                signaturecheck.Enabled = false;
                dsgrp.Enabled = false;
            }
            else
            {
                signaturecheck.Checked = true;
                signaturecheck.Enabled = true;
                dsgrp.Enabled = true;

            }
           // ReportType.SelectedIndex = Int32.Parse(settings[11]);

    */

            ReportType.Enabled = false;
        }

        OpenFileDialog ofd = new OpenFileDialog();
        private void DSpath_Click(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("choose button clicked on digital signature path" + " in" + this.FindForm().Name, "Info");
            ofd.Filter = "(*.BMP;*.JPG;*.PNG)|*.BMP;*.JPG;*.PNG|All files (*.*)|*.*";
            if (ofd.ShowDialog() == DialogResult.OK)
            {
                Logs.LogFile_Entries("File dialogue open" + " in" + this.FindForm().Name, "Info");
                Digitalsignaturepath.Text = ofd.FileName;
                changesFlag = true;

            }
        }

        private void ImportSet_Enter(object sender, EventArgs e)
        {

        }

        private void FileNameLabel_Click(object sender, EventArgs e)
        {

        }

        private void PrinterName_TextChanged(object sender, EventArgs e)
        {

        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            /* if(signaturecheck.Checked==true)
             {
                 dsgrp.Enabled = true;
             }
             else
             {
                 dsgrp.Enabled = false;

             }
             changesFlag = true;*/
        }


        private void PrinterNameLabel_Click(object sender, EventArgs e)
        {

        }

        private void PrintResultsRadioButton_CheckedChanged(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("Print result strip option selected" + "\tPrintResultsRadioButton_CheckedChanged in" + this.FindForm().Name, "Info");
            if (PrintResultsRadioButton.Checked == false)
            {
                AutoPrintCheck.Enabled = false;
                pdfcheck.Enabled = true;
            }
            else
            {
                pdfcheck.Enabled = false;
            }


            ReportType.Enabled = false;
        }

        private void ReportSet_Enter(object sender, EventArgs e)
        {

        }

        private void PrintResultsRadioButton_CheckedChanged_1(object sender, EventArgs e)
        {

        }
        private void ChoosePrinterButton_Click_1(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("Choose printer button clicked" + " \tChoosePrinterButton_Click_1" + this.FindForm().Name, "Info");
            PrintDialog pd = new PrintDialog();
            //in order to open the dialog in win7 64bits
            pd.UseEXDialog = true;
            pd.ShowDialog();
            PrinterName.Text = pd.PrinterSettings.PrinterName;
            printerSet = pd.PrinterSettings;
            changesFlag = true;
        }

        private void ReportType_TextChanged(object sender, EventArgs e)
        {





        }

        private void PrintResultsRadioButton_CheckedChanged_2(object sender, EventArgs e)
        {
            ReportType.Enabled = false;
            pdfcheck.Enabled = false;
            UI_flag = true;
            changesFlag = true;
        }

        private void AutoPrintCheck_CheckedChanged(object sender, EventArgs e)
        {
            changesFlag = true;
        }

        private void Usedefault_CheckedChanged(object sender, EventArgs e)
        {
            if (MakePdfReportRadioButton.Checked == true && ReportType.SelectedIndex == 1 && usedefault.Checked == true)

            {
                Logs.LogFile_Entries("Use default check box checked true" + " in" + this.FindForm().Name, "Info");
                title.Enabled = false;
                subtitleval.Enabled = false;
                titlelab.Enabled = false;
                subtitle.Enabled = false;

            }
            else
            {

                title.Enabled = true;
                subtitleval.Enabled = true;
                titlelab.Enabled = true;
                subtitle.Enabled = true;
            }
            changesFlag = true;
        }

        public void Check_Limit()
        {
            try
            {
                HeaderSpace.MaxLength = 2;
                int val = int.Parse(HeaderSpace.Text);
                if (val > 50)
                {
                    Logs.LogFile_Entries(MessagesLabelsEN.Max_Exceed + " \tCheck_Limit in" + this.FindForm().Name, "Error");
                    MessageBox.Show("Enter value less than 50", LM.Translate("Settings_Error", MessagesLabelsEN.SETTINGS_ERROR));
                    int num = 30;
                    HeaderSpace.Text = num.ToString();

                }
            }
            catch (Exception ex)
            {

            }
        }

        private void HeaderSpace_TextChanged(object sender, EventArgs e)
        {
            Check_Limit();
            changesFlag = true;
        }

        private void HeaderSpace_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void SettingsFrm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Logs.LogFile_Entries("Settings form closed" + " \tSettingsFrm_FormClosed in" + this.FindForm().Name, "Info");

        }

        private void pdfcheck_CheckedChanged(object sender, EventArgs e)
        {
            changesFlag = true;
        }

        private void title_TextChanged(object sender, EventArgs e)
        {
            title.MaxLength = 46;
            changesFlag = true;
        }

        private void subtitleval_TextChanged(object sender, EventArgs e)
        {
            subtitleval.MaxLength = 59;
            changesFlag = true;
        }

        private void PrinterName_TextChanged_1(object sender, EventArgs e)
        {

        }

        private void ReportType_TextChanged_1(object sender, EventArgs e)
        {
            //  UI_flag = true;
        }

        private void Optionalval2_TextChanged(object sender, EventArgs e)
        {
            optionalval2.MaxLength = 15;

        }

        private void Optionalval_TextChanged(object sender, EventArgs e)
        {
            optionalval.MaxLength = 15;
        }
    }
}
