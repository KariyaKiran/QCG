using System;
using System.Collections.Generic;
using System.Linq;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using System.IO.Ports;
using System.Threading;
using iTextSharp.text;
using iTextSharp;
using System.IO;
using iTextSharp.text.pdf;
using System.Data.OleDb;
using System.Drawing.Imaging;
//using System.Windows.Forms.ScrollBars;

using System.Drawing.Printing;
using System.Globalization;
using System.Xml;
using System.Runtime.InteropServices;
using Spire.Pdf.Tables;
using Spire.Pdf.Graphics;
using System.Windows.Forms.DataVisualization.Charting;
using System.Security.Permissions;
using System.Collections;
using Rectangle = iTextSharp.text.Rectangle;

namespace QcGoldArchive
{
    public delegate void guiDelegate(List<string> data);
    public delegate void UpdateControlsDelegate();
    public partial class MainForm : Form
    {

        private SerialPort sp;
        private string currentLang;
        private List<string> printResults = new List<string>();   //buffer for printing
        private List<string> printData;                           //temporary buffer for each print session
        private List<string> oneTestPrintData;
        public List<string> Avg_data = new List<string>();
        public List<string> Avg_rowdata1 = new List<string>();
        public List<string> Avg_rowdata2 = new List<string>();
        List<string> labeldata = new List<string>();
        iTextSharp.text.Image Arrow;
        public List<string> reportData = new List<string>();
        bool test = true;
        public string duplicateData;
        List<string> results = new List<string>();   //List for data, each row is in separate string
        public string footerData;
        public string awData; public string old_patname = string.Empty;
        public string AVGStr1; public string old_tesname = string.Empty;
        public int count; public string old_tesdesg = string.Empty;
        public string awData1; public string old_Liq = string.Empty;
        public string ODStr1; public string old_Fru = string.Empty;
        public string CNTStr1; public string old_Refdr1 = string.Empty;
        public List<string> Archive_strip = new List<string>();
        public bool check = false; public string old_Refdr2 = string.Empty;
        public bool footer = false; public string old_Refdr3 = string.Empty;
        public bool error_check = false;
        public bool archive_data = false;
        public string patname1;

        int cnt = 0;
        public string testername1;
        public string testerdesg1;
        public string fructose1;
        public string liqduraion;
        public string refdr1;
        public string refdr2;
        public string refdr3;
        public byte expType;
        public bool set_check = false;
        public bool DEL_csv_records = false;
        public string AVGStr = "AVG ";
        int max_lines = 0;
        public string ODStr = "OD ";
        public string CNTStr = "CNT ";
        public string Try_AVGStr;
        public string Try_ODStr;
        public string Try_CNTStr;
        public bool exported = true;

        public string test_date_save;

        public string Patient_name;
        public string Tester_name;
        public string Tester_desg;
        public string Fructose_;
        public string Liquifiction_;
        public string Ref_dr;
        public string Ref_dr1;
        public string Ref_dr2;

        public string age_;
        public string optional1_;
        public string manFructose_;
        public string manVitality_;
        public string manRbc_;
        public string manRoundcell_;
        public string managgregation_;
        public string managglutination_;
        public string manOptional_;
        public string manNforms_;
        public string manHeadedefects_;
        public string manpinHeads_;
        public string manneck_midpiece_;
        public string manTailDefects_;
        public string mancytoplasmicDroplets_;
        public string manAcrosome_;
        public string comments_;

        public float immotile;
        public float prog;
        public float SMI_val;
        public float nonprog;
        bool serviceData = false;
        public int row_count;
        private string[] appSettings;
        private Thread guiThread;
        private CultureInfo ExcelCultureInfo = new CultureInfo(UniversalStrings.ENG_US, false);
        private CultureInfo CurrentCulture = CultureInfo.CurrentCulture;
        public char dsChar = Convert.ToChar(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);
        public int counter = 0;
        private const int RF_MESSAGE = 0xA123;
        private bool IsSettingsFrmOpened = false;
        public int csv_empty_check = 0;
        public bool isLowVol;
        public bool isFrozen;
        public bool isFresh;
        public bool isWashed;
        public int totalcols;
        float spacer;

        //ReportTable T_MorpChart;
        LanguageManagement LM = LanguageManagement.CreateInstance();
        Logclass Logs = new Logclass();

        Export export;

        public MainForm()
        {
            InitializeComponent();
            //LM.CreateResxFromDB(); // create new resource file after changes
            guiThread = Thread.CurrentThread;
            //  Plot.ScrollBars = ScrollBars.Vertical;
            this.Text += UniversalStrings.QC_GOLD_ARCHIVE;
            App_Start();                           //initialize all the settings, and update status bar

            Thread.CurrentThread.CurrentCulture = CurrentCulture;
            Thread.CurrentThread.CurrentUICulture = CurrentCulture;
            export = new Export(dsChar, ExcelCultureInfo, appSettings[5]);
            export.OnProcessError += new QcGoldArchive.ProcessError(Export_OnProcessError);
        }

        private void Export_OnProcessError(string exceptionMessage, string title)
        {
            Logs.LogFile_Entries("Export onProcessError" + exceptionMessage, "Error");
            this.Invoke((MethodInvoker)delegate

            {

                MessageBox.Show(this, exceptionMessage, title);

            });
        }




        /*    ***************************************
               Configuring COM port for data receive 
              *************************************** */


        /// <summary>
        /// Get Data from the QwikCheck Gold device 
        /// </summary>
        public void GetData()
        {
            Process p = Process.GetCurrentProcess();
            string myProcessName = p.ProcessName;

            try
            {
                if (Process.GetProcessesByName(myProcessName).Length < 2)  //don't initialize port if there are more than one application
                {
                    sp = new SerialPort(appSettings[0]);
                    sp.BaudRate = 19200;
                    sp.DataBits = 8;
                    sp.StopBits = StopBits.One;
                    sp.Parity = Parity.None;
                    sp.Handshake = Handshake.None;
                    sp.Open();
                    sp.DataReceived += Recieved;
                }
            }

            catch
            {

                MessageBox.Show(this, LM.Translate("PORT_ERROR", MessagesLabelsEN.PORT_ERROR), LM.Translate("PORT_ERROR_CAPTION", MessagesLabelsEN.PORT_ERROR_CAPTION));
                Logs.LogFile_Entries(MessagesLabelsEN.PORT_ERROR + " GetData in " + this.FindForm().Name, "Error");

            }
        }

        /*    *********************************
               Get relevant data from COM port
              *********************************   */

        public void Recieved(object sender, SerialDataReceivedEventArgs e)
        {
            Thread.Sleep(1000);     //Allow the data to be transferred to the buffer
            int numOfBytes = sp.BytesToRead;
            byte[] dataRecieved = new byte[numOfBytes];
            string result = String.Empty;

            byte temp = 0;
            byte prev = 0;
            byte befPrev = 0;
            bool flag = false;
            int count = 0;
            int maxC = 20;
            int maxC2 = 2;
            int countSp = 0;//count for spaces

            isLowVol = false;
            isFrozen = false;
            isFresh = false;
            isWashed = false;
            results = new List<string>(); // old data gets cleared before recieving new test
            serviceData = false;

            sp.Read(dataRecieved, 0, numOfBytes);         //Read the buffer
            Thread.Sleep(150);     // Wait to read the whole buffer
            Logs.LogFile_Entries("Data Received from device of Length " + dataRecieved.Length + "", "Info");
            if (dataRecieved.Length <= 1)
                return;

            for (int i = 0; i < dataRecieved.Length; i++)    //Translate the data into readable text
            {
                befPrev = prev;
                prev = temp;
                temp = dataRecieved[i];

                if (temp == 10 & (prev != 10) | (prev == 27 & temp == 107))
                {
                    results.Add(result);
                    result = String.Empty;
                }


                if (temp == 36 && prev == 27)
                {
                    result += "  ";
                }
                // changed !(prev == 0 & temp == 48) to !(befPrev == 88 & prev == 0 & temp == 48)
                // because it would mistakenly delete 0 of AVG and AW of low quality samples
                if ((temp == 32 | (temp >= 48 & temp <= 57) | (temp >= 65 & temp <= 90) | (temp >= 97 & temp <= 122) | temp == 43 | temp == 46 | temp == 47 | temp == 58 | temp == 35 | temp == 60 | temp == 61 | temp == 62 | temp == 37) & prev != 27 & (!(befPrev == 88 & prev == 0 & temp == 48)) & prev != 36 & !(befPrev == 27 & prev == 105))
                {
                    if (temp == 32)
                        countSp++;
                    else
                        countSp = 0;
                    if (countSp < 10)
                    {
                        result += ((char)temp).ToString();
                        count++;
                    }
                }
                else
                    if (prev == 27 & temp == 36)
                    flag = true;
                else
                        if (flag & prev != 36)
                {
                    flag = false;

                    int j;
                    if (prev == 8 | prev == 164 | prev == 114)
                        j = maxC - count;
                    else
                    {
                        if (prev == 90)
                            j = maxC2 - count;
                        else
                        {
                            if (prev == 74)
                                j = 12 - count;
                            else
                            {
                                if (prev == 224 | prev == 234)
                                    j = 19 - count;
                                else
                                {
                                    if (prev == 240)
                                        j = 9 - count;
                                    else
                                        if (prev == 94)
                                        j = 18 - count;
                                    else
                                            if (prev == 38)
                                        j = 25 - count;
                                    else
                                        j = 10 - count;
                                }
                            }
                        }
                    }

                    if (j < 0)
                        j = 0;
                    while (j != 0)
                    {
                        result += " ";
                        j--;
                        count++;
                    }
                }


            }

            results.Add(result);    //Add last row to the List of data
            string logdata = "";
            for (int i = 0; i <= results.Count - 1; i++)
            {
                logdata += results[i] + "\n";
            }
            Logs.LogFile_Entries("Data Received from Device " + logdata + this.FindForm().Name, "info");
            logdata = "";            // added:
            if (results[2].Contains("SEMEN ANALYSIS REPORT")) // if (results.Count > 27)
            { // is TEST DATA not service data
                int i = 0;
                //if (results.Count <= 30) // low volume or frozen
                //{

                foreach (string item in results)
                {
                    if (item.Contains("MOTILITY RESULTS ONLY")) // low volume
                    {
                        isLowVol = true;
                        break;
                    }
                    else if (item.Contains("LOW QUALITY SAMPLE"))
                    {
                        if (results.Count <= 31)
                            isLowVol = true;
                        break;
                    }
                    else if (item.Contains("FROZEN"))
                    {
                        isFrozen = true;
                    }
                    else if (item.Contains("FRESH"))
                    {
                        isFresh = true;
                    }
                    else if (item.Contains("WASHED"))
                    {
                        isWashed = true;
                    }
                    i++;
                }
                //}
                if (isLowVol == true)
                {
                    List<string> editedres = new List<string>(results);
                    editedres.RemoveAt(i);
                    reportData = RemoveFrozenLabels(editedres);
                    reportData = RemoveUnits(reportData);
                    reportData = RemoveSpaces(reportData);
                }

                else if (isFrozen) //results.Count < 37
                { // is frozen test
                    reportData = RemoveFrozenLabels(results);
                    reportData = RemoveUnits(reportData);
                    reportData = RemoveSpaces(reportData);
                }
                else
                { // is fresh or washed test
                    reportData = RemoveLabels(results);
                    reportData = RemoveUnits(reportData);
                    reportData = RemoveSpaces(reportData);
                }


                footerData = results[results.Count - 3]; // contains AVG, OD, CNT
                awData = results[results.Count - 2]; // contains AW
            }
            else
            { // is SERVICE or SETTINGS DATA
                serviceData = true;
                reportData.Clear(); // service data sent
            }




            this.Invoke(new guiDelegate(PrintOnScreen), results);


            //Synchronize for GUI
            //duplicateData.AddRange(reportData);


            if (!serviceData)
            {
                this.footer = false;
                this.Invoke(new UpdateControlsDelegate(get_old_date));
                this.Invoke(new UpdateControlsDelegate(display));
                try
                {

                    string csvPath = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";
                    if (reportData[18] == "N.")
                    { reportData[18] = "N.A."; }
                    string path = appSettings[2];
                    if (File.Exists(csvPath))
                        expType = 2;
                    else
                        expType = 1;
                    bool cancelled = false;
                    if (expType == 2)
                    {

                        while (CanAccessFile(csvPath) == false && cancelled == false)
                        {
                            this.Invoke(new Action(() =>
                            {
                                exported = false;

                                error_check = true;
                                Logs.LogFile_Entries(MessagesLabelsEN.ACCESS_ERROR + "\t Received in " + this.FindForm().Name, "Error");
                                if (MessageBox.Show(this, csvPath + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                                {
                                    Logs.LogFile_Entries(MessagesLabelsEN.IMPORT_FAILED + "\t Received in " + this.FindForm().Name, "Error");
                                    this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE); }));
                                    cancelled = true;

                                };
                            }));
                        }

                    }
                    if (!cancelled)
                    {





                        if (path != "")
                        {
                            error_check = true;
                            //FileStream aFile;
                            byte exptypes;
                            if (File.Exists(csvPath))
                            { exptypes = 2; }
                            else { exptypes = 1; }

                            if (exptypes == 2)
                            {
                                try
                                {
                                    export.linecount();
                                }
                                catch { }
                            }



                            if (File.Exists(csvPath))
                                expType = 2;
                            else
                                expType = 1;
                            if ((Boolean.Parse(appSettings[7]) == true) || (Boolean.Parse(appSettings[7]) == false))
                            {

                                ExportCsv(csvPath, expType);
                                //  ExportCsv(Application.StartupPath + @"\QC_gold_archive.csv", expType);
                                Logs.LogFile_Entries("Export to CSV on Process", "Info");
                            }
                        }
                    }
                }

                catch (Exception ex)       // Any unexpected error
                {
                    //MessageBox.Show(ex.ToString());
                    //if (!serviceData)
                    //    this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("EXPORT_ERROR", MessagesLabelsEN.EXPORT_ERROR), UniversalStrings.QC_GOLD_ARCHIVE); }));
                    MessageBox.Show(LM.Translate("EXPORT_ERROR", MessagesLabelsEN.EXPORT_ERROR), UniversalStrings.QC_GOLD_ARCHIVE);
                    Logs.LogFile_Entries(MessagesLabelsEN.EXPORT_ERROR + " Received in " + this.FindForm().Name, "Error");
                }
            }
        }


        public void ExportCsv(string filepath, byte ExportType)
        {



            FileStream aFile = null;
            //for (int i = 0; i < reportData.Count; i++)
            //{
            //    if (reportData[i].Contains(patdob.Text))
            //        reportData[i] = "Calculators";
            //}

            if (ExportType == 2)
            {

                try
                {
                    aFile = new FileStream(filepath, FileMode.Append, FileAccess.Write);
                }
                catch
                {
                    string filename;
                    filename = GetFileName(filepath);
                    bool cancelled = false;
                    //while (IsFileOpen(filename) && cancelled == false)

                    while (CanAccessFile(filepath) == false && cancelled == false)
                    {
                        this.Invoke(new Action(() =>
                        {
                            //exported = false;
                            Logs.LogFile_Entries(MessagesLabelsEN.ACCESS_ERROR + " \tExportCsv in " + this.FindForm().Name, "Error");
                            if (MessageBox.Show(this, filepath + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                            {
                                this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE); }));
                                cancelled = true;
                                Logs.LogFile_Entries(MessagesLabelsEN.IMPORT_FAILED + " \tExportCsv in " + this.FindForm().Name, "Error");
                            };
                        }));
                    }

                    if (!cancelled)
                    {





                        aFile = new FileStream(filepath, FileMode.Append, FileAccess.Write);

                        export.FillCsv(filepath, reportData, ExportType, aFile);

                        return;
                    }
                    else
                        return;

                }

            }

            //if (string.IsNullOrEmpty(patname.Text))
            //{
            //    reportData.Insert(4, "---");
            //}
            //if (!string.IsNullOrEmpty(patname.Text))
            //{
            //    reportData.Insert(4, "---");
            //}


            export.FillCsv(filepath, reportData, ExportType, aFile);

            //exported = true;

        }

        private void ExportExcel(string path)
        {
            List<string> Results = new List<string>(reportData);
            List<string> manualentry = new List<string>();
            manualentry.Add(patname.Text);
            manualentry.Add(testername.Text);
            manualentry.Add(testerdesg.Text);
            manualentry.Add(fructosetxt.Text);
            manualentry.Add(liquifictiontxt.Text);
            manualentry.Add(refbydr.Text);
            manualentry.Add(refbydr2.Text);
            manualentry.Add(refbydr3.Text);
            Results.AddRange(manualentry);


            Results = RemoveUnits(Results);
            Results = RemoveSpaces(Results);
            bool success = true;
            if (reportData[18] == "N.")
            { reportData[18] = "N.A"; }
            if (!File.Exists(path))
                success = export.MakeNewExcelFile(path);
            if (success)
            {
                bool cancelled = false;
                while (CanAccessFile(path) == false && cancelled == false)
                {
                    this.Invoke(new Action(() =>
                    {
                        Logs.LogFile_Entries(MessagesLabelsEN.ACCESS_ERROR + "ExportExcel in " + this.FindForm().Name, "Error");
                        if (MessageBox.Show(this, path + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                        {
                            this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE); }));
                            cancelled = true;
                            Logs.LogFile_Entries(MessagesLabelsEN.IMPORT_FAILED + "ExportExcel in " + this.FindForm().Name, "Error");
                        };
                    }));
                }
                if (!cancelled)
                {
                    export.FillExcel(path, Results);
                }
            }
        }
        public void get_old_date()
        {


            if (test == true && duplicateData == null)
            {
                duplicateData = reportData[2];

            }


            //test = false;
            //
            //if (duplicateData==reportData[2] && test == true)
            //{
            //    test = false;
            //}
            //else
            //{
            //    test = true;
            //}

            //if (test != true || cnt==0)
            //{
            //    duplicateData = reportData[2];
            //}
            //else if(duplicateData!=reportData[2])
            //{
            //    duplicateData = reportData[2];
            //}


        }
        public void csv_man_entry()
        {

            char temp = '\0';
            char prev = '\0';
            char befPrev = '\0';
            int pntCnt = 0; // count for points in string
            bool AVGflag = false;
            bool ODflag = false;
            bool CNTflag = false;
            bool CNTpreflag = false;
            string AVGStr = "AVG ";
            string ODStr = "OD ";
            string CNTStr = "CNT ";


            foreach (char c in footerData)
            {
                befPrev = prev;
                prev = temp;
                temp = c;

                // count points
                if (temp == '.')
                    pntCnt++;

                // write char to string
                if (AVGflag)
                    if (temp == ' ' & pntCnt == 2 & prev != ' ')
                        AVGflag = false;
                    else
                        AVGStr += ((char)temp).ToString();
                if (ODflag)
                    if (temp == ' ' & pntCnt == 4 & prev != ' ')
                        ODflag = false;
                    else
                        ODStr += ((char)temp).ToString();
                if (CNTflag)
                    if (temp == ' ' & pntCnt == 5 & prev != ' ' & prev != '.')
                        CNTflag = false;
                    else
                        CNTStr += ((char)temp).ToString();

                // check for label and set flag
                if (befPrev == ' ' & temp == '.')
                    if (prev == '5' & pntCnt == 1)
                        AVGflag = true;
                    else if (prev == '7' & pntCnt == 3)
                        ODflag = true;
                if (befPrev == ' ' & prev == '1' & temp == '2')
                    CNTpreflag = true;
                if (CNTpreflag & temp == '.' & pntCnt == 5)
                    CNTflag = true;

            }


            // remove unnecessary empty spaces
            AVGStr = RemoveSpaces(AVGStr, 'G');
            CNTStr = RemoveSpaces(CNTStr, 'T');
            ODStr = RemoveSpaces(ODStr, 'D');
            awData = RemoveSpaces(awData, 'W');
            Try_AVGStr = AVGStr;
            Try_CNTStr = CNTStr;
            Try_ODStr = ODStr;
            List<string> nullval = new List<string>();
            List<string> nullval2 = new List<string>();
            List<string> labeldata = new List<string>();
            labeldata.Add(AVGStr);
            labeldata.Add(CNTStr);
            labeldata.Add(ODStr);
            labeldata.Add(awData);

            for (int i = 0; i <= 6; i++)
            {
                nullval.Add("---");

            }
            for (int i = 0; i <= 16; i++)
            {
                nullval2.Add("---");


            }

            reportData.AddRange(nullval);
            reportData.AddRange(labeldata);
            reportData.AddRange(nullval2);

        }



        public string GetFileName(string FilePath)
        {
            string fileName;
            int startIndex = 0;
            int endIndex = FilePath.Length - 1;
            for (int i = FilePath.Length - 1; i >= 0; i--)
            {
                char currChar = FilePath[i];
                if (currChar == '.')
                {
                    endIndex = i;
                }
                if (currChar == '\\') // '\'
                {
                    startIndex = i + 1;
                    //i = 0;
                    break;
                }
            }
            fileName = FilePath.Substring(startIndex, (endIndex - startIndex));
            return fileName;
        }

        private bool IsFileOpen(string FileName)
        {
            foreach (Process process in Process.GetProcesses())
            {
                if (process.ProcessName.Contains("EXCEL"))
                {
                    if (process.MainWindowTitle.Contains(FileName))
                        return true;
                }
            }
            return false;
        }
        void line()
        {
            for (int k = 0; k <= 10; k++)
            {

                Plot.Text += "------";


            }
            Plot.Text += "\n";
            Plot.Text += "\n";
        }
        int j = 0;
        public void PrintOnScreen(List<string> data)
        {


            oneTestPrintData = new List<string>();

            if (printResults.Count != 0)
                printResults.Add(LM.Translate("END_OF_REPORT", MessagesLabelsEN.END_OF_REPORT));   //mark to start next page (atleast one report was shown)


            if (Boolean.Parse(appSettings[7]) == true && serviceData)
            {
                MessageBox.Show(LM.Translate("No_Report", MessagesLabelsEN.NO_REPORT_ERROR), LM.Translate("Invalid_data", MessagesLabelsEN.REPORT_Error));
                Logs.LogFile_Entries(MessagesLabelsEN.NO_REPORT_ERROR, "Error");
            }
            else
            {

                if (j >= 1)
                {
                    line();
                }
                for (int i = 0; i < data.Count; i++)
                {


                    Plot.Text += data[i] + "\r\n";
                    printResults.Add(data[i]);
                    oneTestPrintData.Add(data[i]);

                }
                Plot.Text += "\n";
                Logs.LogFile_Entries("Result printed on strip screen" + " \tPrintOnScreen in " + this.FindForm().Name, "Info");
                j++;
                DEL_csv_records = true;
                archive_data = false;
                // SaveSettings();



                if (printResults.Count == 0)
                    printResults.Add("end of report");   //mark to start next page (no report was shown before)
            }

            if ((Boolean.Parse(appSettings[3]) == true && Boolean.Parse(appSettings[7]) == false))  //autoprint
            {
                if (Boolean.Parse(appSettings[7]) == true)
                {
                    if (reportData.Count != 0 && Plot.Text != null)
                    {
                        MakePdf(); // creates new pdf report
                        Logs.LogFile_Entries("Auto print basic Report", "Info");
                    }
                    else
                    {


                    }
                }
                else
                    Printing(true); // makes old report 
                Logs.LogFile_Entries("printing Strip report", "Info");


            }

            if (Boolean.Parse(appSettings[13]) == true && Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 0)
            {
                MakePdf();
                Logs.LogFile_Entries("Auto print basic Report", "Info");
            }


            //  Plot.Text += "\r\n";



        }
        public void display()
        {
            translateType();
            refreshfrm();
            try
            {

                if (SaveButton.Enabled == true)
                {
                    if (((MessageBox.Show(this, LM.Translate("SAVE_DATA", MessagesLabelsEN.SAVE_DATA), LM.Translate("CONFIRM_CHANGES", MessagesLabelsEN.CONFIRM_CHANGES), MessageBoxButtons.YesNo)) == DialogResult.Yes))
                    {
                        updatecsv0();
                        SaveButton.Enabled = false;
                    }
                    else
                    {
                        SaveButton.Enabled = false;
                    }
                }

                Testdate.Text = reportData[2];
                patid.Text = reportData[3];
                patdob.Text = reportData[4];
                absent.Text = reportData[5];
                accession.Text = reportData[6];
                collecteddt.Text = reportData[7];
                rcvdt.Text = reportData[8];

                type.Text = reportData[9];
                volume.Text = reportData[10];
                wbc.Text = reportData[11];
                ph.Text = reportData[12];
                conc.Text = reportData[13];
                concbox.Text = reportData[13];
                totmot.Text = reportData[14];
                totmotbox.Text = reportData[14];
                progmotility.Text = reportData[15];
                rapidProg.Text = reportData[15];
                nonprgmot.Text = reportData[16];
                npmbox.Text = reportData[16];
                immotility.Text = reportData[17];
                immotbox.Text = reportData[17];
                morph.Text = reportData[18];
                mnfbox.Text = reportData[18];
                msc.Text = reportData[19];
                mscbox.Text = reportData[19];
                pmsc.Text = reportData[20];
                rapid_pmscbox.Text = reportData[20];
                fsc.Text = reportData[21];
                fscbox.Text = reportData[21];
                velocity.Text = reportData[22];
                velocitybox.Text = reportData[22];
                smi.Text = reportData[23];
                smibox.Text = reportData[23];
                sperm.Text = reportData[24];
                spermbox.Text = reportData[24];
                motsperm.Text = reportData[25];
                motspermbox.Text = reportData[25];
                progsperm.Text = reportData[26];
                progspermbox.Text = reportData[26];
                funcsperm.Text = reportData[27];
                funcspermbox.Text = reportData[27];
                label2.Text = reportData[28];
                mnfsbox.Text = reportData[28];
                devicesn.Text = reportData[0];

                patname.Enabled = true;
                fructosetxt.Enabled = true;
                liquifictiontxt.Enabled = true;
                testername.Enabled = true;
                testerdesg.Enabled = true;
                refbydr.Enabled = true;
                refbydr2.Enabled = true;
                refbydr3.Enabled = true;
                SaveButton.Enabled = false;
                textBox1.Enabled = true;
                textBox2.Enabled = true;

                cleartext();
                csv_man_entry();
                if (string.IsNullOrEmpty(patname.Text))
                {
                    reportData.Insert(4, "---");
                }
                else
                {
                    reportData.Insert(4, patname.Text);
                }
                Logs.LogFile_Entries("Result printed on main screen" + " \tPrintOnScreen in " + this.FindForm().Name, "Info");

            }
            catch (Exception ex)
            {
                MessageBox.Show(ex.ToString());

            }
            stateOfFields(false);
        }

        public void updatecsv0()
        {
            Logs.LogFile_Entries("Updating CSV in process", "Info");
            bool success = true;

            String path = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";
            List<String> lines = new List<String>();
            Patient_name = "---";
            Tester_name = "---";
            Tester_desg = "---";
            Liquifiction_ = "---";
            Fructose_ = "---";
            Ref_dr = "---";
            Ref_dr1 = "---";
            Ref_dr2 = "---";
            if (File.Exists(path)) ;
            {
                try

                {
                    if (!string.IsNullOrEmpty(patname.Text))
                    { Patient_name = patname.Text; }
                    if (!string.IsNullOrEmpty(testername.Text))
                    { Tester_name = testername.Text; }
                    if (!string.IsNullOrEmpty(testerdesg.Text))
                    { Tester_desg = testerdesg.Text; }
                    if (!string.IsNullOrEmpty(liquifictiontxt.Text))
                    { Liquifiction_ = liquifictiontxt.Text; }
                    if (!string.IsNullOrEmpty(fructosetxt.Text))
                    { Fructose_ = fructosetxt.Text; }
                    if (!string.IsNullOrEmpty(refbydr.Text))
                    { Ref_dr = refbydr.Text; }
                    if (!string.IsNullOrEmpty(refbydr2.Text))
                    { Ref_dr1 = refbydr2.Text; }
                    if (!string.IsNullOrEmpty(refbydr3.Text))
                    { Ref_dr2 = refbydr3.Text; }

                    using (StreamReader reader = new StreamReader(path))
                    {
                        String line;

                        string LineStr;
                        if (dsChar == char.Parse(","))
                            LineStr = ";";
                        else
                            LineStr = ",";
                        while ((line = reader.ReadLine()) != null)
                        {
                            if (line.Contains(LineStr))
                            {
                                String[] split = line.Split(char.Parse(LineStr));

                                if (split[0] == duplicateData)
                                {

                                    split[2] = Patient_name;
                                    split[28] = Tester_name;
                                    split[29] = Tester_desg;
                                    split[30] = Fructose_;
                                    split[31] = Liquifiction_;
                                    split[32] = Ref_dr;
                                    split[33] = Ref_dr1;
                                    split[34] = Ref_dr2;
                                    line = String.Join(LineStr, split);
                                }

                            }

                            lines.Add(line);
                        }
                    }

                    using (StreamWriter writer = new StreamWriter(path, false))
                    {
                        foreach (String line in lines)
                            writer.WriteLine(line);
                    }
                    SaveButton.Enabled = false;
                    duplicateData = reportData[2];
                    test = false;
                    Logs.LogFile_Entries("Updating CSV Completed", "Info");
                }
                catch (Exception ex)
                {
                    MessageBox.Show(ex.ToString());
                    Logs.LogFile_Entries("Updating CSV Failed \t " + ex.ToString() + "\t" + "updatecsv in " + this.FindForm().Name, "Error");
                }
            }


        }

        public void updatecsv()
        {
            Logs.LogFile_Entries("Updating CSV in process", "Info");
            bool success = true;
            string path = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";
            List<String> lines = new List<String>();
            test_date_save = "---";
            Patient_name = "---";
            Tester_name = "---";
            Tester_desg = "---";
            Liquifiction_ = "---";
            Fructose_ = "---";
            Ref_dr = "---";
            Ref_dr1 = "---";
            Ref_dr2 = "---";
            age_ = "---";
            optional1_ = "---";
            manFructose_ = "---";
            manVitality_ = "---";
            manRbc_ = "---";
            manRoundcell_ = "---";
            managgregation_ = "---";
            managglutination_ = "---";
            manOptional_ = "---";
            manNforms_ = "---";
            manHeadedefects_ = "---";
            manpinHeads_ = "---";
            manneck_midpiece_ = "---";
            manTailDefects_ = "---";
            mancytoplasmicDroplets_ = "---";
            manAcrosome_ = "---";
            comments_ = "---";
            accession_save = "---";
            test_date_save = "---";
            abstinance_save = "---";
            patid_save = "---";
            collecteddate_save = "---";
            received_field_save = "---";
            type_save = "---";
            volume_save = "---";
            wbc_save = "---";
            ph_save = "---";
            conc_save = "---";
            totalMotile_save = "---";
            progressiveMot_save = "---";
            nonProg_motility_save = "---";
            immotile_save = "---";
            mnf_save = "---";
            msc_save = "---";
            pmsc_save = "---";
            fsc_save = "---";
            velocity_save = "---";
            smi_save = "---";
            sperm_save = "---";
            motileSperm_save = "---";
            progsperm_save = "---";
            functionSperm_save = "---";
            sperm_save = "---";
            avg_save = "---";
            cnt_str = "---";
            odstr_save = "---";
            aw_save = "---";
            softwareV_save = "---";
            deviceSN_save = "---";
            FileStream aFile = null;



            try

            {
                if (File.Exists(path))
                {
                    if (checkBox2.Checked != true)
                    {
                        if (!string.IsNullOrEmpty(patname.Text))
                        { Patient_name = patname.Text; }
                        if (!string.IsNullOrEmpty(testername.Text))
                        { Tester_name = testername.Text; }
                        if (!string.IsNullOrEmpty(testerdesg.Text))
                        { Tester_desg = testerdesg.Text; }
                        if (!string.IsNullOrEmpty(liquifictiontxt.Text))
                        { Liquifiction_ = liquifictiontxt.Text; }
                        if (!string.IsNullOrEmpty(fructosetxt.Text))
                        { Fructose_ = fructosetxt.Text; }
                        if (!string.IsNullOrEmpty(refbydr.Text))
                        { Ref_dr = refbydr.Text; }
                        if (!string.IsNullOrEmpty(refbydr2.Text))
                        { Ref_dr1 = refbydr2.Text; }
                        if (!string.IsNullOrEmpty(refbydr3.Text))
                        { Ref_dr2 = refbydr3.Text; }
                        if (!string.IsNullOrEmpty(textBox2.Text))
                        { age_ = textBox2.Text; }
                        if (!string.IsNullOrEmpty(textBox1.Text))
                        { optional1_ = textBox1.Text; }

                        if (!string.IsNullOrEmpty(textBox16.Text))
                        { manFructose_ = textBox16.Text; }
                        if (!string.IsNullOrEmpty(textBox7.Text))
                        { manVitality_ = textBox7.Text; }
                        if (!string.IsNullOrEmpty(textBox6.Text))
                        { manRbc_ = textBox6.Text; }
                        if (!string.IsNullOrEmpty(textBox3.Text))
                        { manRoundcell_ = textBox3.Text; }
                        if (!string.IsNullOrEmpty(textBox5.Text))
                        { managgregation_ = textBox5.Text; }
                        if (!string.IsNullOrEmpty(textBox4.Text))
                        { managglutination_ = textBox4.Text; }
                        if (!string.IsNullOrEmpty(textBox15.Text))
                        { manOptional_ = textBox15.Text; }
                        if (!string.IsNullOrEmpty(textBox12.Text))
                        { manNforms_ = textBox12.Text; }
                        if (!string.IsNullOrEmpty(textBox11.Text))
                        { manHeadedefects_ = textBox11.Text; }
                        if (!string.IsNullOrEmpty(textBox8.Text))
                        { manpinHeads_ = textBox8.Text; }
                        if (!string.IsNullOrEmpty(textBox10.Text))
                        { manneck_midpiece_ = textBox10.Text; }
                        if (!string.IsNullOrEmpty(textBox9.Text))
                        { manTailDefects_ = textBox9.Text; }
                        if (!string.IsNullOrEmpty(textBox14.Text))
                        { mancytoplasmicDroplets_ = textBox14.Text; }
                        if (!string.IsNullOrEmpty(textBox13.Text))
                        { manAcrosome_ = textBox13.Text; }
                        if (!string.IsNullOrEmpty(richTextBox1.Text))
                        { comments_ = richTextBox1.Text; }


                        using (StreamReader reader = new StreamReader(path))
                        {
                            String line;

                            string LineStr;
                            if (dsChar == char.Parse(","))
                                LineStr = ";";
                            else
                                LineStr = ",";
                            while ((line = reader.ReadLine()) != null)
                            {
                                if (line.Contains(LineStr))
                                {
                                    String[] split = line.Split(char.Parse(LineStr));

                                    if (split[0].Contains(Testdate.Text))
                                    {

                                        split[2] = Patient_name;
                                        split[28] = Tester_name;
                                        split[29] = Tester_desg;
                                        split[30] = Fructose_;
                                        split[31] = Liquifiction_;
                                        split[32] = Ref_dr;
                                        split[33] = Ref_dr1;
                                        split[34] = Ref_dr2;
                                        split[39] = age_;
                                        split[40] = optional1_;
                                        split[41] = manFructose_;
                                        split[42] = manVitality_;
                                        split[43] = manRbc_;
                                        split[44] = manRoundcell_;
                                        split[45] = managgregation_;
                                        split[46] = managglutination_;
                                        split[47] = manOptional_;
                                        split[48] = manNforms_;
                                        split[49] = manHeadedefects_;
                                        split[50] = manpinHeads_;
                                        split[51] = manneck_midpiece_;
                                        split[52] = manTailDefects_;
                                        split[53] = mancytoplasmicDroplets_;
                                        split[54] = manAcrosome_;
                                        split[55] = comments_;
                                        line = String.Join(LineStr, split);
                                    }

                                }

                                lines.Add(line);
                            }
                        }

                        using (StreamWriter writer = new StreamWriter(path, false))
                        {
                            foreach (String line in lines)
                                writer.WriteLine(line);
                        }
                        SaveButton.Enabled = false;
                        Logs.LogFile_Entries("Updating CSV Completed", "Info");
                    }
                }
                if(checkBox2.Checked==true)
                {
                    if (!string.IsNullOrEmpty(testdateField1.Text))
                    { test_date_save = testdateField1.Text; }
                    if (!string.IsNullOrEmpty(patidField.Text))
                    { patid_save = patidField.Text; }
                    if (!string.IsNullOrEmpty(absField.Text))
                    { abstinance_save = absField.Text; }
                    if (!string.IsNullOrEmpty(accessionField.Text))
                    { accession_save = accessionField.Text; }
                    if (!string.IsNullOrEmpty(collecteddateField.Text))
                    { collecteddate_save = collecteddateField.Text; }
                    if (!string.IsNullOrEmpty(receivedDateField.Text))
                    { received_field_save = receivedDateField.Text; }
                    if (!string.IsNullOrEmpty(typeField.Text))
                    { type_save = typeField.Text; }
                    if (!string.IsNullOrEmpty(volField.Text))
                    { volume_save = volField.Text; }
                    if (!string.IsNullOrEmpty(wbcField.Text))
                    { wbc_save = wbcField.Text; }
                    if (!string.IsNullOrEmpty(absField.Text))
                    { ph_save = phField.Text; }
                    if (!string.IsNullOrEmpty(concbox.Text))
                    { conc_save = concbox.Text; }
                    if (!string.IsNullOrEmpty(totmotbox.Text))
                    { totalMotile_save = totmotbox.Text; }
                    if (!string.IsNullOrEmpty(rapidProg.Text))
                    { progressiveMot_save = rapidProg.Text; }
                    if (!string.IsNullOrEmpty(npmbox.Text))
                    { nonProg_motility_save = npmbox.Text; }
                    if (!string.IsNullOrEmpty(immotbox.Text))
                    { immotile_save = immotbox.Text; }
                    if (!string.IsNullOrEmpty(mnfbox.Text))
                    { mnf_save = mnfbox.Text; }
                    if (!string.IsNullOrEmpty(mscbox.Text))
                    { msc_save = mscbox.Text; }
                    if (!string.IsNullOrEmpty(rapid_pmscbox.Text))
                    { pmsc_save = rapid_pmscbox.Text; }
                    if (!string.IsNullOrEmpty(fscbox.Text))
                    { fsc_save = fscbox.Text; }
                    if (!string.IsNullOrEmpty(velocitybox.Text))
                    { velocity_save = velocitybox.Text; }
                    if (!string.IsNullOrEmpty(smibox.Text))
                    { smi_save = smibox.Text; }
                    if (!string.IsNullOrEmpty(spermbox.Text))
                    { sperm_save = spermbox.Text; }
                    if (!string.IsNullOrEmpty(motspermbox.Text))
                    { motileSperm_save = motspermbox.Text; }
                    if (!string.IsNullOrEmpty(progspermbox.Text))
                    { progsperm_save = progspermbox.Text; }
                    if (!string.IsNullOrEmpty(funcspermbox.Text))
                    { functionSperm_save = funcspermbox.Text; }
                    if (!string.IsNullOrEmpty(spermbox.Text))
                    { sperm_save = spermbox.Text; }
                    if (!string.IsNullOrEmpty(patname.Text))
                    { Patient_name = patname.Text; }
                    if (!string.IsNullOrEmpty(testername.Text))
                    { Tester_name = testername.Text; }
                    if (!string.IsNullOrEmpty(testerdesg.Text))
                    { Tester_desg = testerdesg.Text; }
                    if (!string.IsNullOrEmpty(liquifictiontxt.Text))
                    { Liquifiction_ = liquifictiontxt.Text; }
                    if (!string.IsNullOrEmpty(fructosetxt.Text))
                    { Fructose_ = fructosetxt.Text; }
                    if (!string.IsNullOrEmpty(refbydr.Text))
                    { Ref_dr = refbydr.Text; }
                    if (!string.IsNullOrEmpty(refbydr2.Text))
                    { Ref_dr1 = refbydr2.Text; }
                    if (!string.IsNullOrEmpty(refbydr3.Text))
                    { Ref_dr2 = refbydr3.Text; }
                    if (!string.IsNullOrEmpty(textBox2.Text))
                    { age_ = textBox2.Text; }
                    if (!string.IsNullOrEmpty(textBox1.Text))
                    { optional1_ = textBox1.Text; }
                    if (!string.IsNullOrEmpty(textBox16.Text))
                    { manFructose_ = textBox16.Text; }
                    if (!string.IsNullOrEmpty(textBox7.Text))
                    { manVitality_ = textBox7.Text; }
                    if (!string.IsNullOrEmpty(textBox6.Text))
                    { manRbc_ = textBox6.Text; }
                    if (!string.IsNullOrEmpty(textBox3.Text))
                    { manRoundcell_ = textBox3.Text; }
                    if (!string.IsNullOrEmpty(textBox5.Text))
                    { managgregation_ = textBox5.Text; }
                    if (!string.IsNullOrEmpty(textBox4.Text))
                    { managglutination_ = textBox4.Text; }
                    if (!string.IsNullOrEmpty(textBox15.Text))
                    { manOptional_ = textBox15.Text; }
                    if (!string.IsNullOrEmpty(textBox12.Text))
                    { manNforms_ = textBox12.Text; }
                    if (!string.IsNullOrEmpty(textBox11.Text))
                    { manHeadedefects_ = textBox11.Text; }
                    if (!string.IsNullOrEmpty(textBox8.Text))
                    { manpinHeads_ = textBox8.Text; }
                    if (!string.IsNullOrEmpty(textBox10.Text))
                    { manneck_midpiece_ = textBox10.Text; }
                    if (!string.IsNullOrEmpty(textBox9.Text))
                    { manTailDefects_ = textBox9.Text; }
                    if (!string.IsNullOrEmpty(textBox14.Text))
                    { mancytoplasmicDroplets_ = textBox14.Text; }
                    if (!string.IsNullOrEmpty(textBox13.Text))
                    { manAcrosome_ = textBox13.Text; }
                    if (!string.IsNullOrEmpty(richTextBox1.Text))
                    { comments_ = richTextBox1.Text; }

                    reportData.Insert(0, test_date_save);
                    reportData.Insert(1, patid_save);
                    reportData.Insert(2, Patient_name);
                    reportData.Insert(3, "00/00/00");
                    reportData.Insert(4, abstinance_save);
                    reportData.Insert(5, accession_save);
                    reportData.Insert(6, collecteddate_save);
                    reportData.Insert(7, received_field_save);
                    reportData.Insert(8, type_save);
                    reportData.Insert(9, volume_save);
                    reportData.Insert(10, wbc_save);
                    reportData.Insert(11, ph_save);

                    reportData.Insert(12, conc_save);
                    reportData.Insert(13, totalMotile_save);
                    reportData.Insert(14, progressiveMot_save);
                    reportData.Insert(15, nonProg_motility_save);
                    reportData.Insert(16, immotile_save);
                    reportData.Insert(17, mnf_save);
                    reportData.Insert(18, msc_save);
                    reportData.Insert(19, pmsc_save);
                    reportData.Insert(20, fsc_save);
                    reportData.Insert(21, velocity_save);

                    reportData.Insert(22, smi_save);
                    reportData.Insert(23, sperm_save);
                    reportData.Insert(24, motileSperm_save);
                    reportData.Insert(25, progsperm_save);
                    reportData.Insert(26, functionSperm_save);
                    reportData.Insert(27, sperm_save);


                    reportData.Insert(28, Tester_name);
                    reportData.Insert(29, Tester_desg);
                    reportData.Insert(30, Fructose_);
                    reportData.Insert(31, Liquifiction_);
                    reportData.Insert(32, Ref_dr);
                    reportData.Insert(33, Ref_dr1);
                    reportData.Insert(34, Ref_dr2);

                    reportData.Insert(35, avg_save);
                    reportData.Insert(36, cnt_str);
                    reportData.Insert(37, odstr_save);
                    reportData.Insert(38, aw_save);

                    reportData.Insert(39, age_);
                    reportData.Insert(40, optional1_);
                    reportData.Insert(41, manFructose_);
                    reportData.Insert(42, manVitality_);
                    reportData.Insert(43, manRbc_);
                    reportData.Insert(44, manRoundcell_);
                    reportData.Insert(45, managgregation_);
                    reportData.Insert(46, managglutination_);
                    reportData.Insert(47, manOptional_);
                    reportData.Insert(48, manNforms_);
                    reportData.Insert(49, manHeadedefects_);
                    reportData.Insert(50, manpinHeads_);
                    reportData.Insert(51, manneck_midpiece_);
                    reportData.Insert(52, manTailDefects_);
                    reportData.Insert(53, mancytoplasmicDroplets_);
                    reportData.Insert(54, manAcrosome_);
                    reportData.Insert(55, comments_);
                    reportData.Insert(56, "---");
                    reportData.Insert(57, "---");


                    if (File.Exists(path))
                    {
                        expType = 2;
                        aFile = new FileStream(path, FileMode.Append, FileAccess.Write);

                    }
                    else
                        expType = 1;
                    export.FillCsv(path, reportData, expType, aFile, true);
                    //  ExportCsv(Application.StartupPath + @"\QC_gold_archive.csv", expType);
                    Logs.LogFile_Entries("manual entry Export to CSV on Process", "Info");
                    SaveButton.Enabled = false;
                    Logs.LogFile_Entries("Updating CSV Completed", "Info");
                }
            }
            catch (Exception ex)
            {
                Logs.LogFile_Entries("Updating CSV Failed \t " + ex.ToString() + "\t" + "updatecsv in " + this.FindForm().Name, "Error");
            }


        }
        public void cleartext()
        {
            foreach (Control c in Controls)
            {
                if (c is TextBox)
                {
                    if (!string.IsNullOrEmpty(c.Text))
                    {
                        patname.Text = "";
                        fructosetxt.Text = "";
                        liquifictiontxt.Text = "";
                        testerdesg.Text = "";
                        testername.Text = "";
                        refbydr.Text = "";
                        refbydr2.Text = "";
                        refbydr3.Text = "";
                    }
                }

            }
            SaveButton.Enabled = false;
        }
        public List<string> RemoveLabels(List<string> data)
        {
            List<string> Results = new List<string>();  //list containing all the rows in the printout
            string current = String.Empty;    //represents current string being analyzed
            LM.UpdateLanguage("en");
            string[] labels = LM.Translate(UniversalStrings.REMOVE_LABELS);
            LM.UpdateLanguage(currentLang);

            int posStart;
            int posEnd;

            for (int i = 3; i < labels.Length + 3; i++)    //Removing Labels
            {
                if (((i != 6) & (i != 10)) && i <= (labels.Length + 2))
                {
                    if (!string.IsNullOrEmpty(labels[i - 3]))
                    {

                        current = data[i];
                        posStart = labels[i - 3].Length;
                        posEnd = current.Length - 1;
                        Results.Add(current.Substring(posStart, (posEnd - posStart + 1)));
                    }
                }
                else
                {
                    if ((data[i + 1].Contains("BIRTH")) || (data[i + 1].Contains("COLLECTED")))
                    {
                        data.Insert(i + 1, String.Empty);
                        Results.Add(String.Empty);
                    }
                    else
                    {
                        Results.Add(data[i + 1]);
                    }

                }
            }
            return Results;

        }

        public List<string> RemoveFrozenLabels(List<string> data)
        {
            List<string> Results = new List<string>();  //list containing all the rows in the printout
            string current = String.Empty;    //represents current string being analyzed
            LM.UpdateLanguage("en");
            string[] labels = LM.Translate(UniversalStrings.REMOVE_LABELS);
            string[] frozenLabels = LM.Translate(UniversalStrings.FROZEN_LABELS);
            LM.UpdateLanguage(currentLang);

            int count = 0;
            int posStart;
            int posEnd;

            for (int i = 3; i < labels.Length + 3; i++)    //Removing Labels
            {
                if (((i != 6) & (i != 10)) && i <= (labels.Length + 2))
                {
                    if (!string.IsNullOrEmpty(labels[i - 3]))
                    {

                        if (i > 17)                   // until i=18 no difference between frozen and normal
                        {
                            if ((count < frozenLabels.Length) && (labels[i - 3] == frozenLabels[count]))
                            {
                                if (count < 4)
                                {
                                    current = data[18 + count];              // until i=18 there is no difference between frozen and normal
                                    if ((count == 0) && (current.ToLower().Contains("MOTILITY_RESULTS")))
                                    {
                                        data.RemoveAt(18 + count);
                                        current = data[18 + count];
                                    }
                                }
                                else                                        // eliminate "TOTALS PER VOLUME" label
                                    current = data[18 + count + 1];
                                posStart = labels[i - 3].Length;
                                posEnd = current.Length - 1;
                                Results.Add(current.Substring(posStart, (posEnd - posStart + 1)));

                                count++;
                            }
                            else
                            {
                                Results.Add(UniversalStrings.EMPTY_VALUE);
                            }

                        }
                        else
                        {
                            current = data[i];
                            posStart = labels[i - 3].Length;
                            posEnd = current.Length - 1;
                            Results.Add(current.Substring(posStart, (posEnd - posStart + 1)));
                        }
                    }
                }
                else
                {
                    if ((data[i + 1].Contains("BIRTH")) || (data[i + 1].Contains("COLLECTED")))

                    {
                        data.Insert(i + 1, String.Empty);
                        Results.Add(String.Empty);
                    }
                    else
                    {
                        Results.Add(data[i + 1]);
                    }

                }
            }
            return Results;

        }


        public bool CanAccessFile(string path)
        {
            try
            {
                using (FileStream stream = new FileStream(path, FileMode.Open)) { }
                return true;
            }

            catch
            {
                return false;
            }
        }

        // added function to remove Units from Results
        private List<string> RemoveUnits(List<string> Results)
        {
            string current = String.Empty;    //represents current string being analyzed
            LM.UpdateLanguage("en");
            string[] units = LM.Translate(UniversalStrings.UNITS);
            LM.UpdateLanguage(currentLang);

            int index;

            for (int i = 0; i < Results.Count; i++)        //Removing units
            {
                current = Results[i];
                for (int j = 0; j < units.Length; j++)
                {
                    index = current.IndexOf(units[j]);
                    if (index != -1)
                    {
                        current = current.Remove(index);
                        Results.RemoveAt(i);
                        Results.Insert(i, current);
                    }
                }
            }
            return Results;
        }

        // added function to Remove unnecessary spaces
        private List<string> RemoveSpaces(List<string> Results)
        {
            string current = String.Empty;    //represents current string being analyzed

            for (int i = 0; i < Results.Count; i++)         //Removing unnecessary spaces
            {
                current = Results[i];
                if ((i != 2) && (i != 4) && (i != 7) && (i != 8))
                {
                    for (int j = 0; j < current.Length; j++)
                    {
                        if (current[j] == (char)(32))
                        {
                            current = current.Remove(j, 1);
                            j = j - 1;
                        }
                    }
                }
                else
                {
                    while (current[0] == (char)(32))
                        current = current.Remove(0, 1);
                }

                Results.RemoveAt(i);
                if (current != String.Empty)
                    Results.Insert(i, current);
                else
                    Results.Insert(i, UniversalStrings.EMPTY_VALUE);
            }
            return Results;
        }



        private void Size_Changed(object sender, EventArgs e)
        {
            if (this.WindowState == FormWindowState.Minimized)
            {
                notifyIcon.Visible = true;
            }

        }


        private void ClearButton_Click(object sender, EventArgs e)
        {


            if (rptname.Text == LM.Translate("ADVANCED_REPORT", MessagesLabelsEN.ADVANCED_REPORT) || rptname.Text == LM.Translate("BASIC_REPORT", MessagesLabelsEN.BASIC_REPORT) || rptname.Text == LM.Translate("PRINT_RESULTS_STRIP", MessagesLabelsEN.PRINT_TEST_RESULTS))
            {
                comments.Visible = false;
                Avg_desg.Visible = false;
                Avg_frclbl.Visible = false;
                Avg_liqlbl.Visible = false;
                Avg_tesname.Visible = false;
                Avg_desg.Visible = false;
                Avg_patname.Visible = false;
            }
            Logs.LogFile_Entries("Clear Button clicked" + " \tClearButton_Click in " + this.FindForm().Name, "Info");
            if (MessageBox.Show(LM.Translate("Clear_The_Result", MessagesLabelsEN.CLEAR_MESSAGE), LM.Translate("Clear_Result", MessagesLabelsEN.Clear_Result), MessageBoxButtons.OKCancel, MessageBoxIcon.Question) == DialogResult.OK)
            {
                counter = 0;
                Erase();
                refreshfrm();
                Avg_data.Clear();
                Avg_rowdata2.Clear();
                Avg_rowdata1.Clear();
                SettingsButton.Enabled = true;
                Plot.Text = String.Empty;
                printResults.RemoveRange(0, printResults.Count);
                reportData.Clear();
                DEL_csv_records = false;
                exported = true;
                results.Clear();
                //exported = false;
                archive_data = false;
                j = 0;
                Logs.LogFile_Entries(" Data Cleared on Main screen", "Info");
            }

            stateOfFields(true);

        }
        public void Erase()
        {
            foreach (Control c in Controls)
            {


                Testdate.Text = "";
                patid.Text = "";
                patdob.Text = "";
                absent.Text = "";
                accession.Text = "";
                collecteddt.Text = "";
                rcvdt.Text = "";
                type.Text = "";
                volume.Text = "";
                wbc.Text = "";
                ph.Text = "";
                conc.Text = "";
                totmot.Text = "";
                progmotility.Text = "";
                nonprgmot.Text = "";
                immotility.Text = "";
                morph.Text = "";
                msc.Text = "";
                pmsc.Text = "";
                fsc.Text = "";
                velocity.Text = "";
                smi.Text = "";
                richTextBox1.Text = "";
                sperm.Text = "";
                motsperm.Text = "";
                progsperm.Text = "";
                funcsperm.Text = "";
                label2.Text = " ";
                devicesn.Text = " ";
                Avg_desg.Text = " ";
                Avg_frclbl.Text = " ";
                Avg_liqlbl.Text = " ";
                Avg_patname.Text = " ";
                Avg_tesname.Text = " ";
                comments.Text = " ";
                if (c is TextBox)
                {
                    c.Text = "";
                    c.Enabled = false;
                }
                SaveButton.Enabled = false;
            }
        }
        public void EmptyFields()
        {
            foreach (Control c in Controls)
            {
                if (c is TextBox)
                {
                    c.Text = "";
                    c.Enabled = false;
                }
            }
            foreach (Control c in Controls)
            {
                if (c is Panel)
                {
                    c.Text = "";
                    c.Enabled = false;
                    checkBox1.Checked = false;
                    textBox16.Text = "";
                    textBox7.Text = "";
                    textBox6.Text = "";
                    textBox3.Text = "";
                    textBox5.Text = "";
                    textBox4.Text = "";
                    textBox15.Text = "";
                    textBox12.Text = "";
                    textBox11.Text = "";
                    textBox8.Text = "";
                    textBox10.Text = "";
                    textBox9.Text = "";
                    textBox14.Text = "";
                    textBox13.Text = "";
                    opt.Text = "";
                }
            }
        }
        private void NotifyIcon_Clicked(object sender, MouseEventArgs e)
        {
            this.WindowState = FormWindowState.Normal;
        }

        private void savebtn()
        {
            if (checkBox1.Checked == false && checkBox2.Checked == false)
            {
                if ((fructose.Text != null || liquifiction.Text != null || testername.Text != null || testerdesg.Text != null || patname.Text != null || refbydr.Text != null) || archive_data == true)

                {
                    //if (checkBox2.Checked == false)
                    //{
                    SaveButton.Enabled = true;

                    //}
                }
            }
            else if (checkBox2.Checked == true)
            {
                foreach (Control c in Controls)
                {
                    if (c is TextBox)
                    {
                        if (c.Text != null)
                            SaveButton.Enabled = true;
                    }
                }
            }


        }

        private void svbtndisable()
        {
            if (checkBox1.Checked == false && checkBox2.Checked == false)
            {
                if (fructose.Text == null && liquifiction.Text == null && testername.Text == null && testerdesg.Text == null && patname.Text == null && refbydr.Text == null)
                {
                    SaveButton.Enabled = false;
                }
            }
            else if (checkBox2.Checked == true)
            {
                foreach (Control c in Controls)
                {
                    if (c is TextBox)
                    {
                        if (c.Text == null)
                            SaveButton.Enabled = false;
                    }
                }
            }

        }

        private void pd_PrintPage(object sender, PrintPageEventArgs ev)
        {
            float linesPerPage = 0;
            float yPos = 0;
            int count = 0;
            float leftMargin = ev.MarginBounds.Left;
            float topMargin = ev.MarginBounds.Top;
            String line = null;
            System.Drawing.Font printFont = new System.Drawing.Font(UniversalStrings.Font, 10);

            // Calculate the number of lines per page.
            linesPerPage = ev.MarginBounds.Height / printFont.GetHeight(ev.Graphics);


            // Iterate over the file, printing each line.

            while ((count < linesPerPage) && (printData.Count > 0) && (printData[0] != LM.Translate("END_OF_REPORT", MessagesLabelsEN.END_OF_REPORT)))
            {
                line = printData[0];
                printData.RemoveAt(0);
                yPos = topMargin + (count * printFont.GetHeight(ev.Graphics));
                ev.Graphics.DrawString(line, printFont, Brushes.Black,
                   leftMargin, yPos, new StringFormat());
                count++;

            }

            // If more lines exist, print another page.
            if (printData.Count > 0)
            {
                if (printData[0] == LM.Translate("END_OF_REPORT", MessagesLabelsEN.END_OF_REPORT))
                    printData.RemoveAt(0);
                ev.HasMorePages = true;
            }
            else
                ev.HasMorePages = false;

        }

        // Print the file.
        public void Printing(bool autoPrintFlag)
        {
            FillPrintData(autoPrintFlag);

            try
            {
                using (PrintDocument pd = new PrintDocument())
                {
                    pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);
                    pd.PrinterSettings.PrinterName = appSettings[1];
                    if (pd.PrinterSettings.IsValid == true)
                        pd.Print();
                    else
                        MessageBox.Show(this, LM.Translate("PRINTER_ERROR", MessagesLabelsEN.PRINTER_ERROR), UniversalStrings.QC_GOLD_ARCHIVE);
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show(this, ex.Message);
                Logs.LogFile_Entries(ex.Message + " \tPrinting " + this.FindForm().Name, "Error");

            }
        }

        private void FillPrintData(bool autoPrintFlag)              //adds data for printing (for current printing session)
        {
            try
            {
                printData = new List<string>();
                if (autoPrintFlag == true)
                    printData.AddRange(oneTestPrintData);
                else
                    printData.AddRange(printResults);
            }
            catch (Exception ex)
            {
                Logs.LogFile_Entries(ex.ToString() + " \tFillPrintData " + this.FindForm().Name, "Error");

            }
        }

        private void PrintButton_Click(object sender, System.EventArgs e)
        {
            max_lines = 0;
            Logs.LogFile_Entries("Print Button Clicked", "Info");
            try
            {

                if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 0 && !string.IsNullOrEmpty(patid.Text) && comments.Visible == false)
                {
                    if (reportData.Count == 41)
                    {

                        reportData.RemoveAt(4);
                        //count = count + 1;
                    }
                    MakePdf();
                    Logs.LogFile_Entries("Basic Report Generated", "Info");
                }
                else if (comments.Visible == true && Avg_data.Count != 0)
                {
                    if (patname.Visible == true)
                    {
                        Avg_data[2] = patname.Text;
                    }
                    else
                    {
                        Avg_data[2] = Avg_patname.Text;
                    }
                    Avg_data[32] = comments.Text;
                    Average_report avg_rpt = new Average_report();
                    avg_rpt.averagerpt(Avg_data, Avg_rowdata1, Avg_rowdata2);
                    Avg_data.Clear();
                    Avg_rowdata1.Clear();
                    Avg_rowdata2.Clear();

                    Erase();
                    refreshfrm();

                }

                else if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 1 && !string.IsNullOrEmpty(patid.Text) && refbydr.Visible == true)

                {
                    if (reportData.Count == 40 /*&& reportData[9]==LM.Translate("FRESH",MessagesLabelsEN.FRESH)*/ )
                    {
                        if (string.IsNullOrEmpty(patname.Text))
                        {
                            reportData.Insert(4, "---");
                            count = 0;
                        }
                        else
                        {
                            reportData.Insert(4, patname.Text);
                            count = 0;
                        }

                    }


                    advancedrpt();
                    Logs.LogFile_Entries("Advance Report Generated", "Info");

                }


                else if (Boolean.Parse(appSettings[7]) == false && !string.IsNullOrEmpty(Plot.Text) && results[1].Contains("QwikCheck Gold "))
                {
                    Printing(true);
                    Logs.LogFile_Entries("strip report printed", "Info");

                }

                else
                {
                    if (checkBox2.Checked)
                    {
                        advancedrpt();
                        Logs.LogFile_Entries(MessagesLabelsEN.REPORT_ERROR + " \t Generate report with enable fields " + this.FindForm().Name, "Error");
                    }
                    else
                    {

                        MessageBox.Show(this, LM.Translate("No_Data", MessagesLabelsEN.REPORT_ERROR), LM.Translate("REPORT_ERROR_CAPTION", MessagesLabelsEN.REPORT_ERROR_CAPTION));
                        //MessageBox.Show(this, LM.Translate("REPORT_ERROR", MessagesLabelsEN.REPORT_ERROR), LM.Translate("REPORT_ERROR_CAPTION", MessagesLabelsEN.REPORT_ERROR_CAPTION));
                        Logs.LogFile_Entries(MessagesLabelsEN.REPORT_ERROR + " \tPrintButton_Click in " + this.FindForm().Name, "Error");
                    }
                }

            }

            catch (Exception ex)
            {
                Logs.LogFile_Entries(ex.ToString() + " \tPrintButton_Click " + this.FindForm().Name, "Error");

                //MessageBox.Show(this, LM.Translate("REPORT_ERROR", MessagesLabelsEN.REPORT_ERROR), LM.Translate("REPORT_ERROR_CAPTION", MessagesLabelsEN.REPORT_ERROR_CAPTION));
            }

        }
        List<string> deviceValues = new List<string>();

        List<string> patientValues = new List<string>();
        List<string> sampleValues = new List<string>();
        List<string> resultValues = new List<string>();

        public void emptylines()
        {

            patname1 = patname.Text;
            testername1 = testername.Text;
            testerdesg1 = testerdesg.Text;
            fructose1 = fructosetxt.Text;
            liqduraion = liquifictiontxt.Text;
            refdr1 = refbydr.Text;
            refdr2 = refbydr2.Text;
            refdr3 = refbydr3.Text;

            if (patname1 == UniversalStrings.EMPTY_VALUE)
            { patname1 = ""; }
            if (testername1 == UniversalStrings.EMPTY_VALUE)
            { testername1 = ""; }
            if (testerdesg1 == UniversalStrings.EMPTY_VALUE)
            { testerdesg1 = ""; }
            if (string.IsNullOrEmpty(fructosetxt.Text))
            { fructose1 = "---"; }
            if (string.IsNullOrEmpty(liquifictiontxt.Text))

            { liqduraion = "---"; }
            if (string.IsNullOrEmpty(refbydr.Text))
            { refdr1 = "---"; }
            if (string.IsNullOrEmpty(refbydr2.Text))
            { refdr2 = "---"; }
            if (string.IsNullOrEmpty(refbydr3.Text))
            { refdr3 = "---"; }





        }
        // new function.

        public void MakePdf()
        {


            bool pdfCreated;
            int i;
            char temp = '\0';
            char prev = '\0';
            char befPrev = '\0';
            int pntCnt = 0; // count for points in string
            bool AVGflag = false;
            bool ODflag = false;
            bool CNTflag = false;
            bool CNTpreflag = false;

            AVGStr = "";
            CNTStr = "";
            ODStr = "";
            AVGStr = "AVG ";
            ODStr = "OD ";
            CNTStr = "CNT ";
            List<string> deviceValues = new List<string>();

            List<string> patientValues = new List<string>();
            List<string> sampleValues = new List<string>();
            List<string> resultValues = new List<string>();

            float headerSpace, defaultHeaderSpace;
            bool frozen;
            //string defStr;

            defaultHeaderSpace = 0;

            // translate test type
            frozen = translateType();

            for (i = 0; i < reportData.Count; i++)
            {
                if (i < 2)
                    deviceValues.Add(reportData[i]);
                else if (i > 2 & i < 5)
                    patientValues.Add(reportData[i]);
                else if (i == 2 | i >= 5 & i <= 12)
                {
                    if (i != 2 & i != 3 & i != 4 & i != 7 & i != 8 & reportData[i] != "N.A."/*& (Double.TryParse(reportData[i], out num))*/)
                    {
                        if (dsChar != char.Parse("."))
                            sampleValues.Add(reportData[i].Replace('.', ','));
                        else
                            sampleValues.Add(reportData[i]);
                    }
                    else
                        sampleValues.Add(reportData[i]);
                }

                else
                {
                    if (i != 2 & i != 3 & i != 4 & i != 7 & i != 8 & reportData[i] != "N.A." /*& (Double.TryParse(reportData[i], out num))*/)
                    {
                        if (dsChar != char.Parse("."))
                            resultValues.Add(reportData[i].Replace('.', ','));
                        else
                            resultValues.Add(reportData[i]);
                    }
                    else
                        resultValues.Add(reportData[i]);
                }
            }

            // instantiation


            ReportTableBuilder reportTableBuilder = new ReportTableBuilder();

            PdfReportGenerator pdfReportGen = new PdfReportGenerator();

            DataTable DeviceInformationTable = new DataTable();
            reportTableBuilder.CreateDeviceInformation(DeviceInformationTable, deviceValues);

            DataTable PatientInformationTable = new DataTable();
            reportTableBuilder.CreatePatientInformation(PatientInformationTable, patientValues);

            DataTable SampleInformationTable = new DataTable();
            reportTableBuilder.CreateSampleInformationTable(SampleInformationTable, sampleValues);

            DataTable SemAnalParametersTable = new DataTable();
            bool shortReport = false;
            if (frozen || isLowVol)
            {
                shortReport = true;
            }
            reportTableBuilder.CreateSemAnalTestResultsTable(SemAnalParametersTable, resultValues, shortReport);
            DataRow SemAnalParametersDR = SemAnalParametersTable.NewRow();
            SemAnalParametersDR["parameter"] = LM.Translate("PARAMETER", MessagesLabelsEN.PARAMETER);
            SemAnalParametersDR["result"] = LM.Translate("RESULT", MessagesLabelsEN.RESULT);

            // file path in temp folder
            try
            {
                var dateTime = DateTime.Now;
                string formatedDate = dateTime.ToString();
                var replacement = formatedDate.Replace('/', '_');
                replacement = replacement.Replace(':', '_');
                formatedDate = replacement;

                string filePath = Path.GetTempPath()
                       + LM.Translate("REPORT_FILE_NAME", MessagesLabelsEN.REPORT_FILE_NAME) + "_" + deviceValues[0] + "_" + formatedDate + ".pdf";


                // extract values for AVG, CNT, OD
                try
                {
                    if (footerData != null)
                    {

                        foreach (char c in footerData)
                        {
                            befPrev = prev;
                            prev = temp;
                            temp = c;

                            // count points
                            if (temp == '.')
                                pntCnt++;

                            // write char to string
                            if (AVGflag)
                                if (temp == ' ' & pntCnt == 2 & prev != ' ')
                                    AVGflag = false;
                                else
                                    AVGStr += ((char)temp).ToString();
                            if (ODflag)
                                if (temp == ' ' & pntCnt == 4 & prev != ' ')
                                    ODflag = false;
                                else
                                    ODStr += ((char)temp).ToString();
                            if (CNTflag)
                                if (temp == ' ' & pntCnt == 5 & prev != ' ' & prev != '.')
                                    CNTflag = false;
                                else
                                    CNTStr += ((char)temp).ToString();

                            // check for label and set flag
                            if (befPrev == ' ' & temp == '.')
                                if (prev == '5' & pntCnt == 1)
                                    AVGflag = true;
                                else if (prev == '7' & pntCnt == 3)
                                    ODflag = true;
                            if (befPrev == ' ' & prev == '1' & temp == '2')
                                CNTpreflag = true;
                            if (CNTpreflag & temp == '.' & pntCnt == 5)
                                CNTflag = true;

                        }

                        // remove unnecessary empty spaces
                        AVGStr = RemoveSpaces(AVGStr, 'G');
                        CNTStr = RemoveSpaces(CNTStr, 'T');
                        ODStr = RemoveSpaces(ODStr, 'D');
                        awData = RemoveSpaces(awData, 'W');

                    }
                    if (this.footer == true)
                    {
                        AVGStr = AVGStr1;
                        CNTStr = CNTStr1;
                        ODStr = ODStr1;
                        awData = awData1;
                    }


                }
                catch
                {

                }
                // create footer string
                string footerStr = LM.Translate("PRINTED_FROM", MessagesLabelsEN.PRINTED_FROM) + " " + UniversalStrings.QC_GOLD_ARCHIVE + " | " + LM.Translate("DEVICE_SN_WU", MessagesLabelsEN.DEVICE_SN_WU) + " " + reportData[0] + " | " + dateTime + " | " +
                AVGStr + " | " + awData + " | " + CNTStr + " | " + ODStr;

                // add header space
                if (Boolean.Parse(appSettings[8]) == true)
                    headerSpace = float.Parse(appSettings[9]);
                else
                    headerSpace = defaultHeaderSpace;

                // generate pdf report

                pdfCreated = pdfReportGen.GenerateSEMENANALYSISReport(headerSpace,/*Environment.CurrentDirectory*/ filePath,
                    LM.Translate("REPORT_FILE_NAME", MessagesLabelsEN.REPORT_FILE_NAME),
                    LM.Translate("TEST_REPORT", MessagesLabelsEN.TEST_REPORT),
                    LM.Translate("REPORT_CONTINUED", MessagesLabelsEN.REPORT_CONTINUED) + " | " +
                    LM.Translate("PATIENT_ID", MessagesLabelsEN.PATIENT_ID) + patientValues[0] + " | " +
                    "TEST DATE / TIME: " + sampleValues[0], footerStr,
                    LM.Translate("DEVICE_INFO", MessagesLabelsEN.DEVICE_INFO), DeviceInformationTable,
                    LM.Translate("PATIENT_INFO", MessagesLabelsEN.PATIENT_INFO), PatientInformationTable,
                    LM.Translate("SAMPLE_INFO", MessagesLabelsEN.SAMPLE_INFO), SampleInformationTable,
                    SemAnalParametersDR, SemAnalParametersTable);


                // open pdf document
                if (pdfCreated == true)
                {
                    Process.Start(filePath);
                }
                else
                {
                    MessageBox.Show(this, LM.Translate("REPORT_CREATION_ERROR", MessagesLabelsEN.REPORT_CREATION_ERROR), LM.Translate("REPORT_ERROR_CAPTION", MessagesLabelsEN.REPORT_ERROR_CAPTION));
                    Logs.LogFile_Entries(MessagesLabelsEN.REPORT_CREATION_ERROR + " \tmakepdf in " + this.FindForm().Name, "Error");
                }
            }
            catch (Exception ex)
            {
                Logs.LogFile_Entries(ex.ToString() + " \tmakepdf in " + this.FindForm().Name, "Error");

            }
        }



        public void advancedrpt()
        {

            Logs.LogFile_Entries("Generating Advanced Report", "Info");
            char temp = '\0';
            char prev = '\0';
            char befPrev = '\0';
            int pntCnt = 0; // count for points in string
            bool AVGflag = false;
            bool ODflag = false;
            bool CNTflag = false;
            bool CNTpreflag = false;
            string AVGStr = "AVG ";
            string ODStr = "OD ";
            string CNTStr = "CNT ";
            try
            {
                immotile = float.Parse(immotbox.Text);
                prog = float.Parse(rapidProg.Text);
                nonprog = float.Parse(npmbox.Text);
            }
            catch { }
            try
            {
                SMI_val = float.Parse(smibox.Text);
            }
            catch
            {

            }
            emptylines();

            //string path = Application.StartupPath + @"\settings.xml";
            //appSettings = XmlUtility.ReadSettings(path);
            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);
            var dateTime = DateTime.Now;

            LoadMorph();
            LoadSMI();





            MyPdfPageEventHelpPageNo rptgenerator = new MyPdfPageEventHelpPageNo();

            var stringDate = dateTime.ToString().Replace('/', '_');
            stringDate = stringDate.Replace(':', '_');

            string filePath = Path.GetTempPath() + "Semen_Analysis_Report" + stringDate + ".pdf";


            Document doc = new Document(PageSize.A4, 5F, 5F, 0F, 45F);
            try
            {

                PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create));

                doc.Open();




                string subheader2;
                string head;
                string subhead;
                int headerspace;
                int age = 0;
                float space;
                try
                {

                    //string dob = "15/11/1986";
                    int fourDigitYear = 0;
                    string trimtext = reportData[5].Substring(6);
                    int year = int.Parse(trimtext);
                    fourDigitYear = year + (DateTime.Today.Year - DateTime.Today.Year % 100);
                    if (fourDigitYear > DateTime.Today.Year)
                        fourDigitYear = fourDigitYear - 100;
                    // int fourDigitYear = CultureInfo.CurrentCulture.Calendar.ToFourDigitYear(year);
                    if (reportData[5] != "00/00/00")
                    {
                        try
                        {
                            DateTime userdate = Convert.ToDateTime(reportData[5]);
                            int usermonth = int.Parse(reportData[5].Substring(3, 2));
                            if (DateTime.Now.Month > usermonth)
                            {
                                age = DateTime.Now.Year - fourDigitYear;
                            }
                            else
                            {
                                if ((DateTime.Now.Month >= usermonth) && DateTime.Today.Date.Day >= userdate.Day)
                                {

                                    age = DateTime.Now.Year - fourDigitYear;
                                }
                                else
                                {
                                    age = (DateTime.Now.Year - fourDigitYear) - 1;
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            age = (DateTime.Now.Year - fourDigitYear) - 1;
                        }
                    }
                }
                catch { }

                if (Boolean.Parse(appSettings[8]) == true)

                {

                    headerspace = int.Parse(appSettings[9]);

                    space = Utilities.MillimetersToPoints(headerspace);

                }

                else
                {
                    space = Utilities.MillimetersToPoints(15);
                }
                if (Boolean.Parse(appSettings[4]) == true)
                {
                    head = LM.Translate("TEST_REPORT", MessagesLabelsEN.Report_head);
                    subhead = " ";
                    subheader2 = " ";
                }

                else
                {
                    head = appSettings[12];
                    if (appSettings[14].Contains(":"))
                    {
                        subhead = appSettings[14].Substring(0, appSettings[14].LastIndexOf(":"));
                        subheader2 = appSettings[14].Substring(appSettings[14].LastIndexOf(":") + 1);
                    }
                    else
                    {
                        subhead = appSettings[14];
                        subheader2 = "";
                    }

                }



                iTextSharp.text.Font pfont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 15, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font headers = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 11, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font headers2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font headers3 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 9, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                iTextSharp.text.Font Charttxt = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                iTextSharp.text.Font pfont2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 13, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font datafont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font datafooter = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.GRAY);
                iTextSharp.text.Font datafont2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 10, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

                if (appSettings[6] == "zh-CN")
                {
                    var fontPath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
                    FontFactory.Register(fontPath);
                    pfont1 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 13, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    headers = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 11, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    headers2 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    headers3 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    Charttxt = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    pfont2 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 11, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                    datafont = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                    datafooter = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.GRAY);
                    datafont2 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);

                }




                Paragraph para = new Paragraph(head, pfont1);
                Paragraph para1 = new Paragraph(subhead, pfont2);
                Paragraph para2 = new Paragraph(subheader2, pfont1);
                para.SpacingBefore = space;



                //para.IndentationLeft = 240f;
                //para1.IndentationLeft = 220f;
                para.Alignment = Element.ALIGN_CENTER;
                para1.Alignment = Element.ALIGN_CENTER;
                para2.Alignment = Element.ALIGN_CENTER;

                PdfPTable tab = new PdfPTable(2);
                Paragraph patinfo = new Paragraph(LM.Translate("PATIENT_INFO", MessagesLabelsEN.PATIENT_INFO), headers2);
                Paragraph SPACE = new Paragraph("\n");
                Paragraph footername = new Paragraph(LM.Translate("NAME", MessagesLabelsEN.NAME), datafont2);
                Paragraph footerdesg = new Paragraph(LM.Translate("DESIGNATION", MessagesLabelsEN.DESIGNATION), datafont2);
                PdfContentByte pcb = wri.DirectContent;

                PdfPCell header1 = new PdfPCell(new Phrase(patinfo));

                header1.Colspan = 2;

                header1.BorderWidthBottom = 0.6f;
                header1.BorderWidthLeft = 0;
                header1.BorderWidthRight = 0;
                header1.BorderWidthTop = 0;
                header1.PaddingBottom = 8;
                header1.BackgroundColor = new BaseColor(176, 224, 230);
                header1.BorderWidthRight = 0.6f;
                header1.BorderWidthLeft = 0.6f;
                header1.BorderWidthTop = 0.6f;
                header1.VerticalAlignment = Element.ALIGN_CENTER;
                header1.PaddingTop = 5;

                header1.HorizontalAlignment = 0;
                PdfPCell cell1 = new PdfPCell(new Phrase(LM.Translate("PATIENT_ID", MessagesLabelsEN.PATIENT_ID) + ": ", datafont));
                PdfPCell cell1value;
                if (checkBox2.Checked)
                {
                    cell1value = new PdfPCell(new Phrase(patidField.Text, datafont));
                }
                else
                {
                    cell1value = new PdfPCell(new Phrase(reportData[3], datafont));
                }
                PdfPCell cell2 = new PdfPCell(new Phrase(LM.Translate("PAT_NAME", MessagesLabelsEN.PATIENT_NAME) + ": ", datafont));
                PdfPCell cell2value = new PdfPCell(new Phrase(patname1, datafont));
                PdfPCell cell3 = new PdfPCell(new Phrase(LM.Translate("AGE", MessagesLabelsEN.AGE) + ": ", datafont));
                PdfPCell cell3value = new PdfPCell(new Phrase(textBox2.Text, datafont));

                //tab.HorizontalAlignment = 0;//new BaseColor(176, 224, 230)
                tab.AddCell(header1);
                AddTextCell(tab, cell1, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                AddTextCell(tab, cell1value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab, cell2, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                AddTextCell(tab, cell2value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab, cell3, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                AddTextCell(tab, cell3value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                tab.TotalWidth = 266f;
                tab.WriteSelectedRows(0, -1, 20, 770 - space, pcb);
                tab.HorizontalAlignment = 2;


                //reference by doctor
                PdfPTable tab2 = new PdfPTable(2);
                PdfPCell header2 = new PdfPCell(new Phrase(" ", datafont));
                header2.Colspan = 2;

                header2.BorderWidthBottom = 0.6f;


                PdfPCell nill = new PdfPCell(new Phrase(" ", pfont1));

                PdfPCell fill = new PdfPCell(new Phrase(" ", datafont));
                PdfPCell cell4 = new PdfPCell(new Phrase(LM.Translate("REF_BY_DR", MessagesLabelsEN.REF_BY_DR) + ":", datafont));
                PdfPCell cell4value = new PdfPCell(new Phrase(refdr1, datafont));
                PdfPCell cell4avalue = new PdfPCell(new Phrase(refdr2, datafont));
                PdfPCell cell4bvalue = new PdfPCell(new Phrase(refdr3, datafont));
                header1.HorizontalAlignment = 0;
                tab2.HorizontalAlignment = 2;

                AddTextCell(tab2, nill, Element.ALIGN_LEFT, 1, 1, 0.6f, 0.6f, 0.6f, 0, 5, 0, new BaseColor(176, 224, 230));
                AddTextCell(tab2, nill, Element.ALIGN_LEFT, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 0, new BaseColor(176, 224, 230));

                AddTextCell(tab2, cell4, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                AddTextCell(tab2, cell4value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab2, fill, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                AddTextCell(tab2, cell4avalue, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab2, fill, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                AddTextCell(tab2, cell4bvalue, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                tab2.TotalWidth = 266f;
                tab2.WriteSelectedRows(0, -1, 305, 770 - space, pcb);


                if (reportData.Count != 0 || checkBox2.Checked)
                {
                    //sample information table
                    PdfPTable tab3 = new PdfPTable(3);

                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("SAMPLE_INFO", MessagesLabelsEN.SAMPLE_INFO), headers2)), Element.ALIGN_LEFT, 1, 3, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    //headercolor(tab3, nill, Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0.6f, 0.6f, 5, 5);



                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("COLLECTED_DATE_WU", MessagesLabelsEN.COLLECTED_DATE_WU), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(collecteddateField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(reportData[8], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }

                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("RECEIVED_WU", MessagesLabelsEN.RECEIVED_WU), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(receivedDateField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(reportData[9], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }

                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("TEST_DATE_WU", MessagesLabelsEN.TEST_DATE_WU), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(testdateField1.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(reportData[2], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }

                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("ACCESSION", MessagesLabelsEN.ACCESSION), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(accessionField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(reportData[7], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }


                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("SAMPLE_TYPE", MessagesLabelsEN.SAMPLE_TYPE), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(typeField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(type.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }



                    AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("ABSTINENCE_WU", MessagesLabelsEN.ABSTINENCE_WU), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(absField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab3, new PdfPCell(new Phrase(reportData[6], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }

                    tab3.TotalWidth = 266f;

                    tab3.WriteSelectedRows(0, -1, 20, 692 - space, pcb);

                    tab3.HorizontalAlignment = 0;



                    //Semen examination table



                    PdfPTable tab4 = new PdfPTable(3);





                    AddTextCell(tab4, new PdfPCell(new Phrase(LM.Translate("SEMEN_EXAMINATION", MessagesLabelsEN.SEMEN_EXAMINATION), headers2)), Element.ALIGN_LEFT, 1, 3, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    //headercolor(tab4, new PdfPCell(new Phrase("RESULT", headers)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0.6f, 0.6f, 5, 5);

                    AddTextCell(tab4, new PdfPCell(new Phrase(LM.Translate("VOLUME_WU", MessagesLabelsEN.VOLUME), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab4, new PdfPCell(new Phrase(volField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }
                    else
                    {
                        AddTextCell(tab4, new PdfPCell(new Phrase(reportData[11], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    }

                    AddTextCell(tab4, new PdfPCell(new Phrase(LM.Translate("LIQUIFACTION", MessagesLabelsEN.LIQUIFICTION), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab4, new PdfPCell(new Phrase(liqduraion, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    if (string.IsNullOrEmpty(appSettings[16]))
                    {
                        appSettings[16] = "---";
                    }
                    AddTextCell(tab4, new PdfPCell(new Phrase(appSettings[16].ToUpper(), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab4, new PdfPCell(new Phrase(fructose1, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab4, new PdfPCell(new Phrase(LM.Translate("WBC_CONC_WU", MessagesLabelsEN.WBC_CONC_WU), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab4, new PdfPCell(new Phrase(wbcField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    }
                    else
                    {
                        AddTextCell(tab4, new PdfPCell(new Phrase(reportData[12], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    }

                    AddTextCell(tab4, new PdfPCell(new Phrase(LM.Translate("pH", MessagesLabelsEN.PH), datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                    if (checkBox2.Checked)
                    {
                        AddTextCell(tab4, new PdfPCell(new Phrase(phField.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    }
                    else
                    {
                        AddTextCell(tab4, new PdfPCell(new Phrase(reportData[13], datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    }
                    if (string.IsNullOrEmpty(textBox1.Text))
                    {
                        textBox1.Text = "---";
                    }
                    if (string.IsNullOrEmpty(appSettings[15]))
                    {
                        appSettings[15] = "---";
                    }
                    AddTextCell(tab4, new PdfPCell(new Phrase(appSettings[15].ToUpper(), datafont)), Element.ALIGN_LEFT, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    AddTextCell(tab4, new PdfPCell(new Phrase(textBox1.Text.ToLower(), datafont)), Element.ALIGN_CENTER, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    tab4.TotalWidth = 266f;

                    tab4.WriteSelectedRows(0, -1, 305, 692 - space, pcb);



                    //parameters table


                    PdfPTable tab5 = new PdfPTable(55);



                    AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("PARAMETER", MessagesLabelsEN.PARAMETER), headers2)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("RESULT", MessagesLabelsEN.RESULT), headers2)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell(tab5, new PdfPCell(new Phrase("UNIT", headers2)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell1(tab5, new PdfPCell(new Phrase(LM.Translate("Ref_val", MessagesLabelsEN.Ref_Val), headers2)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0, 5, 5, 0, new BaseColor(176, 224, 230), 18, BaseColor.BLACK);
                    AddTextCell(tab5, new PdfPCell(new Phrase("CONCENTRATION", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase(concbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("M/ml", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);


                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 15", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    string red_arrow_path = Application.StartupPath + @"\assets\RED.png";
                    string green_arrow_path = Application.StartupPath + @"\assets\GREEN.png";

                    try
                    {
                        if (float.Parse(concbox.Text) >= 15)
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(green_arrow_path);
                            Arrow.ScaleToFit(10, 10);
                        }
                        else
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                            Arrow.ScaleToFit(10, 10);

                        }
                    }
                    catch (Exception ex)
                    {

                        Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                        Arrow.ScaleToFit(10, 10);



                    }
                    //AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 3, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("TOTAL MOTILITY <PR+NP>", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase(totmotbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("%", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    try
                    {
                        if (float.Parse(totmotbox.Text) >= 40)
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(green_arrow_path);
                            Arrow.ScaleToFit(10, 10);
                        }
                        else
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                            Arrow.ScaleToFit(10, 10);

                        }
                    }
                    catch (Exception ex)
                    {

                        Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                        Arrow.ScaleToFit(10, 10);



                    }

                    headercolor(tab5, new PdfPCell(new Phrase(">= 40", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                    //AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, new BaseColor(243, 243, 243));
                    AddTextCell(tab5, new PdfPCell(new Phrase("PROG. MOTILITY <PR>", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase(rapidProg.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("%", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);


                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 32", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    try
                    {
                        if (float.Parse(rapidProg.Text) >= 32)
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(green_arrow_path);
                            Arrow.ScaleToFit(10, 10);
                        }
                        else
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                            Arrow.ScaleToFit(10, 10);

                        }
                    }
                    catch (Exception ex)
                    {

                        Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                        Arrow.ScaleToFit(10, 10);



                    }


                    //AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("NON-PROG. MOTILITY <NP>", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase(npmbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("%", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    headercolor(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase("IMMOTILE <IM>", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase(immotbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("%", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("MORPH.NORM.FORMS < WHO 5th >", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    if (mnfbox.Text == "N.A.")
                    { mnfbox.Text = "N/A"; }
                    AddTextCell(tab5, new PdfPCell(new Phrase(mnfbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("%", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase(">= 4", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    try
                    {
                        if (float.Parse(mnfbox.Text) >= 4)
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(green_arrow_path);
                            Arrow.ScaleToFit(10, 10);
                        }
                        else
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                            Arrow.ScaleToFit(10, 10);

                        }
                    }
                    catch (Exception ex)
                    {

                        Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                        Arrow.ScaleToFit(10, 10);



                    }


                    //AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase("MOTILE SPERM CONCENTRATION", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    if (mscbox.Text == "N.A.")
                    { mscbox.Text = "N/A"; }
                    headercolor(tab5, new PdfPCell(new Phrase(mscbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("M/ml", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 6", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("PROG. MOTILE SPERM CONC.", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    if (rapid_pmscbox.Text == "N.A.")
                    { rapid_pmscbox.Text = "N/A"; }
                    AddTextCell(tab5, new PdfPCell(new Phrase(rapid_pmscbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("M/ml", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase(">= 5", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase("FUNCTIONAL SPERM CONC.", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    if (fscbox.Text == "N.A.")
                    { fscbox.Text = "N/A"; }
                    headercolor(tab5, new PdfPCell(new Phrase(fscbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("M/ml", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);


                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 0.2", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("VELOCITY", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    if (velocitybox.Text == "N.A.")
                    { velocitybox.Text = "N/A"; }
                    AddTextCell(tab5, new PdfPCell(new Phrase(velocitybox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("(mic/sec)", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase(">= 5", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                   
                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase("SPERM MOTILITY INDEX", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    if (smibox.Text == "N.A.")
                    { smibox.Text = "N/A"; }
                    headercolor(tab5, new PdfPCell(new Phrase(smibox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 80", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("TOTAL EJACULATION ", headers2)), Element.ALIGN_LEFT, 1, 55, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));



                    headercolor(tab5, new PdfPCell(new Phrase("TOTAL SPERM CONC.", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    if (spermbox.Text == "N.A.")
                    { spermbox.Text = "N/A"; }
                    AddTextCell(tab5, new PdfPCell(new Phrase(spermbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("M/ejac", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase(">= 39", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    try
                    {
                        if (float.Parse(spermbox.Text) >= 39)
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(green_arrow_path);
                            Arrow.ScaleToFit(10, 10);
                        }
                        else
                        {

                            Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                            Arrow.ScaleToFit(10, 10);

                        }
                    }
                    catch (Exception ex)
                    {

                        Arrow = iTextSharp.text.Image.GetInstance(red_arrow_path);
                        Arrow.ScaleToFit(10, 10);



                    }


                    //AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase("TOTAL MOTILE SPERM CONC.", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    if (motspermbox.Text == "N.A.")
                    { motspermbox.Text = "N/A"; }
                    headercolor(tab5, new PdfPCell(new Phrase(motspermbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("M/ejac", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 16", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("TOTAL PROG. MOT. SPERM CONC.", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    if (rapidProg.Text == "N.A.")
                    { rapidProg.Text = "N/A"; }
                    AddTextCell(tab5, new PdfPCell(new Phrase(progspermbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("M/ejac", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    headercolor(tab5, new PdfPCell(new Phrase(">= 12", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                    AddTextCell(tab5, new PdfPCell(new Phrase("TOTAL FUNC. SPERM CONC.", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    if (funcspermbox.Text == "N.A.")
                    { funcspermbox.Text = "N/A"; }
                    headercolor(tab5, new PdfPCell(new Phrase(funcspermbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase("M/ejac", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(tab5, new PdfPCell(new Phrase(">= 0.6", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                    headercolor(tab5, new PdfPCell(new Phrase("TOTAL MORPH.NORM.SPERM", datafont)), Element.ALIGN_LEFT, 1, 29, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                    if (mnfsbox.Text == "N.A.")
                    { mnfsbox.Text = "N/A"; }

                    AddTextCell(tab5, new PdfPCell(new Phrase(mnfsbox.Text, datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0.6f, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase("M/ejac", datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0.6f, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    headercolor(tab5, new PdfPCell(new Phrase(">= 2", datafont)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0.6f, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                    //AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0.6f, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                    tab5.TotalWidth = 286f;


                    tab5.WriteSelectedRows(0, -1, 20, 562 - space, pcb);

                    tab5.HorizontalAlignment = 0;
                }

                PdfPTable comments = new PdfPTable(1);
                var phrase = new Phrase();
                phrase.Add(new Chunk("COMMENTS : ", headers2));
                phrase.Add(new Chunk(richTextBox1.Text, datafont));

                AddTextCell(comments, new PdfPCell(phrase), Element.ALIGN_LEFT, 1, 1, 0, 0f, 0, 0, 5, 5, BaseColor.WHITE);

                comments.TotalWidth = 550f;
                comments.WriteSelectedRows(0, -1, 20, 230 - space, pcb);



                PdfPTable tabfooter = new PdfPTable(1);

                tabfooter.WidthPercentage = 45;

                AddTextCell(tabfooter, new PdfPCell(new Phrase("Reviewed by : " + testername1, datafont)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);

                AddTextCell(tabfooter, new PdfPCell(new Phrase(LM.Translate("DESIGNATION", MessagesLabelsEN.DESIGNATION) + ": " + testerdesg1, datafont)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);

                tabfooter.HorizontalAlignment = 0;

                tabfooter.TotalWidth = 266f;
                if (bool.Parse(appSettings[17]) == false)
                {
                    tabfooter.WriteSelectedRows(0, -1, 20, 195 - space, pcb);
                }



                PdfPTable sign = new PdfPTable(1);

                sign.TotalWidth = 266f;

                AddTextCell(sign, new PdfPCell(new Phrase(LM.Translate("SIGNATURE", MessagesLabelsEN.SIGNATURE) + ": ", datafont)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);

                //AddCellheight(sign, new PdfPCell(png), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE, 1);

                sign.HorizontalAlignment = 5;
                if (bool.Parse(appSettings[17]) == false)
                {
                    sign.WriteSelectedRows(0, -1, 305, 177 - space, pcb);
                }
                //footer Data

                string p = appSettings[10];




                PdfPTable signImg = new PdfPTable(1);
                try

                {
                    iTextSharp.text.Image png = iTextSharp.text.Image.GetInstance(p);
                    string signature = appSettings[10];

                    if (!string.IsNullOrEmpty(appSettings[10].ToString()))

                    {





                        // png.ScalePercent(15f);

                        png.ScaleAbsoluteHeight(30f);

                        var newWidth = png.Width * png.ScaledHeight / png.Height;

                        png.ScaleAbsoluteWidth(newWidth);


                        signImg.TotalWidth = 150f;

                        AddTextCell(signImg, new PdfPCell(png), Element.ALIGN_LEFT, 1, 1, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                        if (bool.Parse(appSettings[17]) == false)
                        {
                            signImg.WriteSelectedRows(0, -1, 375, 200 - space, pcb);
                        }
                    }


                }

                catch (Exception ex)

                {
                    //MessageBox.Show(ex.ToString());
                    Logs.LogFile_Entries(ex.ToString() + " \tAdvancedrpt in " + this.FindForm().Name, "Error");



                }









                Rectangle page1 = doc.PageSize;

                PdfPTable footeraddress = new PdfPTable(25);
                footeraddress.WidthPercentage = 45;
                string footeraddressstring = "Plot no. 1751, Door no. 107, 1st Floor, I- block, 13th Main Road, Anna Nagar West, Chennai-600040, PH: +91 9944125909 (Corporate).";
                AddTextCell(footeraddress, new PdfPCell(new Phrase(footeraddressstring, datafooter)), Element.ALIGN_CENTER, 1, 25, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(footeraddress, new PdfPCell(new Phrase(" ", datafont)), Element.ALIGN_CENTER, 1, 0, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                footeraddress.TotalWidth = page1.Width - doc.LeftMargin - doc.RightMargin;
                footeraddress.HorizontalAlignment = 0;
                //footeraddress.WriteSelectedRows(0, -1, doc.LeftMargin, doc.BottomMargin - 30, pcb);

                try
                {
                    if (footerData != null)
                    {
                        foreach (char c in footerData)
                        {
                            befPrev = prev;
                            prev = temp;
                            temp = c;

                            // count points
                            if (temp == '.')
                                pntCnt++;

                            // write char to string
                            if (AVGflag)
                                if (temp == ' ' & pntCnt == 2 & prev != ' ')
                                    AVGflag = false;
                                else
                                    AVGStr += ((char)temp).ToString();
                            if (ODflag)
                                if (temp == ' ' & pntCnt == 4 & prev != ' ')
                                    ODflag = false;
                                else
                                    ODStr += ((char)temp).ToString();
                            if (CNTflag)
                                if (temp == ' ' & pntCnt == 5 & prev != ' ' & prev != '.')
                                    CNTflag = false;
                                else
                                    CNTStr += ((char)temp).ToString();

                            // check for label and set flag
                            if (befPrev == ' ' & temp == '.')
                                if (prev == '5' & pntCnt == 1)
                                    AVGflag = true;
                                else if (prev == '7' & pntCnt == 3)
                                    ODflag = true;
                            if (befPrev == ' ' & prev == '1' & temp == '2')
                                CNTpreflag = true;
                            if (CNTpreflag & temp == '.' & pntCnt == 5)
                                CNTflag = true;

                        }

                    }
                    // remove unnecessary empty spaces
                    AVGStr = RemoveSpaces(AVGStr, 'G');
                    CNTStr = RemoveSpaces(CNTStr, 'T');
                    ODStr = RemoveSpaces(ODStr, 'D');
                    awData = RemoveSpaces(awData, 'W');



                    string footerStr = "Tested By Automated QwikCheck Gold" + " | " + LM.Translate("DEVICE_SN_WU", MessagesLabelsEN.DEVICE_SN_WU) + " " + reportData[0] + " | " + dateTime + " | " +
           AVGStr + " | " + awData + " | " + CNTStr + " | " + ODStr;

                    PdfPTable footer = new PdfPTable(55);
                    footer.WidthPercentage = 45;
                    AddTextCell(footer, new PdfPCell(new Phrase(footerStr, datafooter)), Element.ALIGN_CENTER, 1, 55, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                    AddTextCell(footer, new PdfPCell(new Phrase(" ", datafont)), Element.ALIGN_CENTER, 1, 0, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                    footer.HorizontalAlignment = 0;
                    Rectangle page = doc.PageSize;
                    footer.TotalWidth = page.Width - doc.LeftMargin - doc.RightMargin;
                    if (bool.Parse(appSettings[17]) == false)
                    {
                        footer.WriteSelectedRows(0, -1, doc.LeftMargin, doc.BottomMargin - 28, pcb);
                    }
                    //footer.TotalWidth = 500f;
                    //footer.WriteSelectedRows(0, -1, 230, 80 - 25, pcb);


                }
                catch { }
                if (this.footer == true)
                {
                    PdfPTable footer1 = new PdfPTable(25);
                    footer1.WidthPercentage = 45;
                    string footerStr1 = "Tested By Automated QwikCheck Gold" + " | " + LM.Translate("DEVICE_SN_WU", MessagesLabelsEN.DEVICE_SN_WU) + " " + reportData[0] + " | " + dateTime + " | " +
                    AVGStr1 + " | " + awData1 + " | " + CNTStr1 + " | " + ODStr1;
                    AddTextCell(footer1, new PdfPCell(new Phrase(footerStr1, datafooter)), Element.ALIGN_CENTER, 1, 25, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                    AddTextCell(footer1, new PdfPCell(new Phrase(" ", datafont)), Element.ALIGN_CENTER, 1, 0, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                    footer1.HorizontalAlignment = 0;
                    Rectangle page = doc.PageSize;
                    footer1.TotalWidth = page.Width - doc.LeftMargin - doc.RightMargin;

                    //footer1.TotalWidth = 800f;
                    //footer1.WriteSelectedRows(0, -1, 20, 80 - 25, pcb);


                    if (bool.Parse(appSettings[17]) == false)
                    {
                        footer1.WriteSelectedRows(0, -1, doc.LeftMargin, doc.BottomMargin - 28, pcb);
                    }
                }


                doc.Add(Chunk.NEWLINE);
                doc.Add(para);
                doc.Add(para1);
                doc.Add(para2);
                doc.Add(SPACE);
                //doc.Add(tab);
                PdfPTable Morphchart = new PdfPTable(2);
                headercolor(Morphchart, new PdfPCell(new Phrase(LM.Translate("Analysing_charts", MessagesLabelsEN.Analysing_Charts).ToUpper(), headers2)), Element.ALIGN_CENTER, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                //if (refbydr.Text != "---" && refbydr2.Text != "---" && refbydr3.Text != "---")
                if (rapidProg.Text != "---" && npmbox.Text != "---" && immotbox.Text != "---")
                {
                    AddTextCell1(Morphchart, new PdfPCell(new Phrase(LM.Translate("MOTILITY", MessagesLabelsEN.Motility), Charttxt)), Element.ALIGN_CENTER, 1, 2, 0.6f, 0, 0.6f, 0.6f, 0, 0, 0, BaseColor.WHITE, 20, BaseColor.BLACK);

                    iTextSharp.text.Image motchart = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\Morphology_chart.png"); // ("Morphology_chart.png");

                    //motchart.ScaleToFit(190f, 200f);
                    AddTextCell1(Morphchart, new PdfPCell(motchart), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0.6f, 0.6f, 0, 40, 20, BaseColor.WHITE, 220, BaseColor.BLACK);

                    Morphchart.TotalWidth = 266f;
                    Morphchart.WriteSelectedRows(0, -1, 305.02f, 562f - space, pcb);

                }
                else
                {
                    AddTextCell1(Morphchart, new PdfPCell(new Phrase("", headers)), Element.ALIGN_CENTER, 1, 2, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, 5, BaseColor.WHITE, 200, BaseColor.BLACK);
                    Morphchart.TotalWidth = 266f;
                    Morphchart.WriteSelectedRows(0, -1, 305.02f, 562f - space, pcb);
                    PdfPTable NO_morph = new PdfPTable(2);
                    AddTextCell1(NO_morph, new PdfPCell(new Phrase(LM.Translate("No_Mot_Chart", MessagesLabelsEN.NO_MOTILITY_CHART).ToUpper(), headers2)), Element.ALIGN_CENTER, 1, 2, 3f, 3f, 3f, 3f, 5, 5, 5, BaseColor.WHITE, 100, new BaseColor(211, 211, 211));


                    NO_morph.TotalWidth = 200f;
                    NO_morph.WriteSelectedRows(0, -1, 336.01f, 490f - space, pcb);

                }




                iTextSharp.text.Image smichart = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\SMI_chart.png"); // ("Morphology_chart.png");

                PdfPTable Smichart = new PdfPTable(2);
                string Smi_head = "SMI";
                if (smibox.Text != "0" && smibox.Text != "---")
                {
                    AddCellheight(Smichart, new PdfPCell(new Phrase(LM.Translate("SMI", Smi_head), Charttxt)), Element.ALIGN_CENTER, 1, 2, 0, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE, 0);
                    //AddTextCell(Smichart, new PdfPCell(new Phrase("", Charttxt)), Element.ALIGN_CENTER, 1, 25, 0, 0, 0, 0.6f, 5, 0, BaseColor.WHITE);
                    //smichart.ScaleToFit(500, 100);

                    AddTextCell1(Smichart, new PdfPCell(smichart), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 10, 0, BaseColor.WHITE, 115, BaseColor.BLACK);
                    Smichart.TotalWidth = 266f;

                    Smichart.WriteSelectedRows(0, -1, 305.02f, 375f - space, pcb);

                }
                else
                {
                    AddTextCell1(Smichart, new PdfPCell(new Phrase("", headers)), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 5, 5, BaseColor.WHITE, 137, BaseColor.BLACK);
                    Smichart.TotalWidth = 266f;

                    Smichart.WriteSelectedRows(0, -1, 305.02f, 375f - space, pcb);

                    PdfPTable NO_smi = new PdfPTable(2);
                    AddTextCell1(NO_smi, new PdfPCell(new Phrase(LM.Translate("NO_SMI_CHART", MessagesLabelsEN.NO_SMI_CHART).ToUpper(), headers2)), Element.ALIGN_CENTER, 1, 2, 3f, 3f, 3f, 3f, 5, 5, 5, BaseColor.WHITE, 100, new BaseColor(211, 211, 211));
                    AddTextCell1(Smichart, new PdfPCell(new Phrase("", headers)), Element.ALIGN_TOP, 1, 2, 0, 0.6f, 0, 0.6f, 5, 5, 5, BaseColor.WHITE, 100, BaseColor.BLACK);


                    NO_smi.TotalWidth = 200f;
                    NO_smi.WriteSelectedRows(0, -1, 336.01f, 350f - space, pcb);

                }


                doc.Add(SPACE);
                // doc.Add(tab3);

                doc.Add(SPACE);
                // doc.Add(tab5);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);
                if (checkBox1.Checked == true)
                {
                    doc.NewPage();
                    //Paragraph pa = new Paragraph(" ", pfont1);
                    //pa.Alignment = Element.ALIGN_CENTER;
                    //doc.Add(Chunk.NEWLINE);
                    //pa.SpacingBefore = space;
                    //doc.Add(pa);
                    tab.WriteSelectedRows(0, -1, 20, 780 - space, pcb);
                    tab.HorizontalAlignment = 2;
                    tab2.WriteSelectedRows(0, -1, 305, 780 - space, pcb);
                    // doc.Add(tabfooter);

                    PdfPTable manualmorph = new PdfPTable(3);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("MANUAL PARAMETERS", headers2)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("RESULT", headers2)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("FRUCTOSE", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox16.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("VITALITY", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox7.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("RBC", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox6.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("ROUND CELLS/DEBRIS", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox3.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase("AGGLUTINATION", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox4.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);


                    AddTextCell(manualmorph, new PdfPCell(new Phrase("AGGREGATION", datafont)), Element.ALIGN_LEFT, 1, 2, 0, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox5.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    string optional = "";
                    if (opt.Text == "" || opt.Text == "Optional Field")
                    {
                        optional = "OPTIONAL";
                    }
                    else
                    {
                        optional = opt.Text;
                    }

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(optional, datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0.6F, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualmorph, new PdfPCell(new Phrase(textBox15.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6F, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    PdfPTable manualparam = new PdfPTable(3);

                    AddTextCell(manualparam, new PdfPCell(new Phrase("MANUAL MORPHOLOGY", headers2)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell(manualparam, new PdfPCell(new Phrase("RESULT", headers2)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                    AddTextCell(manualparam, new PdfPCell(new Phrase("NORMAL FORMS", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase(textBox12.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase("HEAD DEFECTIVES", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase(textBox11.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);



                    AddTextCell(manualparam, new PdfPCell(new Phrase("NECK/MIDPIECE", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase(textBox10.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase("TAIL DEFECTIVES", datafont)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase(textBox9.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    AddTextCell(manualparam, new PdfPCell(new Phrase("CYTOPLASMIC DROPLETS", datafont)), Element.ALIGN_LEFT, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);


                    AddTextCell(manualparam, new PdfPCell(new Phrase(textBox14.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(manualparam, new PdfPCell(new Phrase("ACROSOME/NUCLEUS DEFECTIVES", datafont)), Element.ALIGN_LEFT, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                    //AddTextCell(manualparam, new PdfPCell(new Phrase(textBox13.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);
                    if (textBox8.Text != "")
                    {
                        AddTextCell(manualparam, new PdfPCell(new Phrase("PINHEADS", datafont)), Element.ALIGN_LEFT, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                        AddTextCell(manualparam, new PdfPCell(new Phrase(textBox8.Text, datafont)), Element.ALIGN_CENTER, 1, 1, 0, 0.6f, 0, 0.6f, 5, 5, BaseColor.WHITE);

                    }

                    try
                    {
                        if (textBox12.Text == "")
                        {
                            textBox12.Text = 0.ToString();
                        }
                        if (textBox11.Text == "")
                        {
                            textBox11.Text = 0.ToString();
                        }
                        if (textBox8.Text == "")
                        {
                            textBox8.Text = 0.ToString();
                        }
                        if (textBox10.Text == "")
                        {
                            textBox10.Text = 0.ToString();
                        }
                        if (textBox9.Text == "")
                        {
                            textBox9.Text = 0.ToString();
                        }
                        if (textBox14.Text == "")
                        {
                            textBox14.Text = 0.ToString();
                        }
                        if (textBox13.Text == "")
                        {
                            textBox13.Text = 0.ToString();
                        }
                        morph_chart(int.Parse(textBox12.Text), int.Parse(textBox11.Text), int.Parse(textBox8.Text), int.Parse(textBox10.Text), int.Parse(textBox9.Text), int.Parse(textBox14.Text), 0);
                        PdfPTable halo_Chart_img = new PdfPTable(2);
                        iTextSharp.text.Image png1 = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\halo_Chart.png");
                        png1.ScaleToFit(280f, 280f);
                        AddTextCell0(halo_Chart_img, new PdfPCell(png1), Element.ALIGN_LEFT, 1, 2, 0, 0, 0, 0, 5, 80, BaseColor.WHITE, 135, 0, 450);
                        

                        PdfPTable addtionalTxt1 = new PdfPTable(2);
                        iTextSharp.text.Image sampletxt1 = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\assets\sample text1.png");
                        sampletxt1.ScaleToFit(280f, 280f);
                        AddTextCell0(addtionalTxt1, new PdfPCell(sampletxt1), Element.ALIGN_LEFT, 1, 2, 0, 0, 0, 0, 5, 80, BaseColor.WHITE, 135, 0, 450);


                        PdfPTable addtionalTxt2 = new PdfPTable(2);
                        iTextSharp.text.Image sampletxt2 = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\assets\sample text2.png");
                        sampletxt2.ScaleToFit(480f, 500f);
                        AddTextCell0(addtionalTxt2, new PdfPCell(sampletxt2), Element.ALIGN_CENTER, 1, 2, 0, 0, 0, 0, 0, 0, BaseColor.WHITE, 5, 0, 00);

                        

                        halo_Chart_img.TotalWidth = 560f;
                        halo_Chart_img.WriteSelectedRows(0, -1, 165, 565 - space, pcb);

                        addtionalTxt1.TotalWidth = 400f;
                        addtionalTxt1.WriteSelectedRows(0, -1, -120, 625 - space, pcb);

                        addtionalTxt2.TotalWidth = 560f;
                        addtionalTxt2.WriteSelectedRows(0, -1, 0, 325 - space, pcb);

                        manualmorph.TotalWidth = 266f;

                        manualmorph.WriteSelectedRows(0, -1, 20, 689 - space, pcb);

                        manualmorph.HorizontalAlignment = 0;

                        manualparam.TotalWidth = 266f;

                        manualparam.WriteSelectedRows(0, -1, 305, 689 - space, pcb);
                    }
                    catch { }

                    tabfooter.HorizontalAlignment = 0;

                    tabfooter.TotalWidth = 266f;
                    if (bool.Parse(appSettings[17]) == false)
                    {
                        tabfooter.WriteSelectedRows(0, -1, 20, 195 - space, pcb);
                        sign.WriteSelectedRows(0, -1, 305, 177 - space, pcb);
                        signImg.WriteSelectedRows(0, -1, 375, 200 - space, pcb);
                        footeraddress.WriteSelectedRows(0, -1, doc.LeftMargin, doc.BottomMargin - 30, pcb);
                    }

                    doc.Add(Chunk.NEWLINE);
                    //doc.Add(para);
                    //doc.Add(para1);
                    doc.Add(SPACE);
                }
            }
            catch (Exception Ex)
            {
                Logs.LogFile_Entries(Ex.ToString() + " \tAdvancedrpt in " + this.FindForm().Name, "Error");

            }





            doc.Close();
            System.Diagnostics.Process.Start(filePath);
        }





        public void placeholder()
        {

            opt.Text = "Optional Field";
            opt.ForeColor = Color.Silver;

        }




        public void AddTextCell0(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, BaseColor bgrndcolor, int paddingleft, int paddingright, int cellheight)
        {

            cell.HorizontalAlignment = cellHorizontalAlignment;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = ColSpan;
            cell.Rowspan = RowSpan;
            cell.BorderWidthTop = borderWidthTop;
            cell.BorderWidthBottom = borderWidthBottom;
            cell.BorderWidthLeft = borderWidthLeft;
            cell.BorderWidthRight = borderWidthRight;
            cell.PaddingTop = paddingtop;
            cell.PaddingBottom = paddingbottom;
            cell.PaddingLeft = paddingleft;
            cell.PaddingRight = paddingright;

            cell.FixedHeight = cellheight;
            cell.BackgroundColor = bgrndcolor;

            table.AddCell(cell);

        }




        public void LoadMorph()
        {
            Logs.LogFile_Entries("Generating Motility graph", "Info");
            byte[] newImage = new byte[0];
            try
            {
                //Store chart properties
                var chart = new Chart();
                chart.Height = 1000;
                chart.Width = 1100;



                var chartArea1 = new ChartArea();
                chart.ChartAreas.Add(chartArea1);
                chart.ChartAreas[0].AlignmentStyle = AreaAlignmentStyles.All;
                Series series1;


                string seriesName1 = " ";



                series1 = new Series();
                seriesName1 = "PIE Chart";
                series1.Name = seriesName1;
                series1.ChartType = SeriesChartType.Pie;
                chart.Series.Add(series1);

                chart.Legends.Add(new Legend("Legend2"));
                chart.Series[seriesName1].Legend = "Legend2";
                chart.Series[seriesName1].IsVisibleInLegend = true;

                chart.Legends["Legend2"].Position = new ElementPosition(-8, 65, 110, 10);
                chart.Legends["Legend2"].Font = new System.Drawing.Font("Arial", 18F, FontStyle.Bold);
                chart.ChartAreas[0].Position.Auto = false;
                chart.ChartAreas[0].Position.X = 5;
                chart.ChartAreas[0].Position.Y = 0;
                chart.ChartAreas[0].Position.Height = 65;
                chart.ChartAreas[0].Position.Width = 75;





                //immotile = int.Parse(refbydr.Text);
                //prog = int.Parse(refbydr2.Text);
                //nonprog = int.Parse(refbydr3.Text);


                series1.Points.AddXY(0, immotile);


                // chart.Series[seriesName1].Label = "#PERCENT\n#VALY";




                series1.Points[0].Color = Color.Red;

                chart.Series[seriesName1]["PieLabelStyle"] = "Outside";

                chart.Series[seriesName1].BorderColor = Color.Black;
                chart.Series[seriesName1].BorderWidth = 0;
                chart.Series[seriesName1].ShadowColor = Color.Gray;
                chart.Series[seriesName1].Points[0].Label = /*LM.Translate("IMMOTILITY", MessagesLabelsEN.IMMOTILITY_TEXT) +*/ "(#VALY%)";
                chart.Series[seriesName1].Points[0].Font = new System.Drawing.Font("Arial", 26F);

                chart.Series[seriesName1].Points[0].LegendText = LM.Translate("IMMOT_WU", MessagesLabelsEN.IMMOT_WU);
                //chart.Series[seriesName1].Points[0].sty = LegendImageStyle.Marker;




                series1.Points.AddXY(1, prog);

                // chart.Series[seriesName1].Label = "#VALX";

                chart.Series[seriesName1].Points[1].Color = Color.Green;
                chart.Series[seriesName1].Points[1].Label = /*LM.Translate("TOTALPROG_MOTILITY", MessagesLabelsEN.TOTALPROGRESSIVE_MOTILITY) +*/ "(#VALY%)";
                chart.Series[seriesName1].Points[1].LegendText = LM.Translate("Tot_Prg_Mot", MessagesLabelsEN.TOTALPROGRESSIVE_MOTILITY) + "(%)";
                chart.Series[seriesName1].Points[1].Font = new System.Drawing.Font("Arial", 26F);


                series1.Points.AddXY(2, nonprog);
                chart.Series[seriesName1].Points[2].Color = Color.Yellow;

                //chart.Series[seriesName1].Label = "#PERCENT";
                chart.Series[seriesName1].Points[2].Label = /*LM.Translate("NONPROG_MOTILITY", MessagesLabelsEN.NONPROGRESSIVE_MOTILITY) +*/ "(#VALY%)";
                chart.Series[seriesName1].Points[2].LegendText = LM.Translate("Prg_Mot", MessagesLabelsEN.NONPROGRESSIVE_MOTILITY) + "(%)";
                chart.Series[seriesName1].Points[2].Font = new System.Drawing.Font("Arial", 26F);
                series1.IsVisibleInLegend = true;


                chart.Series[seriesName1].Font = new System.Drawing.Font("Arial", 8F, FontStyle.Bold);


                chart.SaveImage(Application.StartupPath + @"\Morphology_chart.Png", ChartImageFormat.Png);
                //string fss = Application.StartupPath +  "Morphology_chart" + ".jpg";
                //System.Drawing.Image MorpImg = System.Drawing.Image.FromFile(fss)



            }
            catch (Exception ex)
            {

                Logs.LogFile_Entries(ex.ToString() + " \tLoadMorph in " + this.FindForm().Name, "Error");

            }

        }

        public void LoadSMI()
        {
            try
            {
                Logs.LogFile_Entries("Generating SMI graph", "Info");
                Chart barchart = new Chart();
                barchart.Size = new Size(810, 1500);



                ChartArea area = new ChartArea();
                barchart.ChartAreas.Add(area);

                barchart.BackColor = Color.Transparent;
                barchart.ChartAreas[0].BackColor = Color.Transparent;
                //barchart.ChartAreas[0].AxisY.MinorTickMark.Enabled = true;
                //barchart.ChartAreas[0].AxisX.MajorGrid.Enabled = false;
                barchart.ChartAreas[0].Area3DStyle.Inclination = 50;
                //barchart.ChartAreas[0].AxisY.MajorGrid.Enabled = false;
                //barchart.ChartAreas[0].AxisX.MinorGrid.Enabled = true;
                //barchart.ChartAreas[0].AxisY.MinorGrid.Enabled = true;
                barchart.ChartAreas[0].Area3DStyle.Enable3D = true;
                barchart.ChartAreas[0].Area3DStyle.Inclination = 5;

                // barchart.ChartAreas[0].AxisY.Interval = 60;



                Series series1 = new Series()
                {
                    Name = "seriesname",
                    ChartType = SeriesChartType.Bar,
                    IsVisibleInLegend = true,
                    BorderWidth = 3,
                    Font = new System.Drawing.Font("Helvitica", 26f),



                };
                barchart.Series.Add(series1);
                barchart.Series["seriesname"]["LabelStyle"] = "Right";
                //barchart.Series["seriesname"].SmartLabelStyle.IsMarkerOverlappingAllowed = false;
                barchart.Series["seriesname"]["PointWidth"] = "0.5";
                barchart.ChartAreas[0].AxisX.LabelStyle.Angle = 90;
                barchart.ChartAreas[0].AxisY.IsLabelAutoFit = false;
                barchart.ChartAreas[0].AxisY.LabelStyle.Angle = 90;


                barchart.Series["seriesname"].LabelAngle = 90;
                barchart.ChartAreas[0].Area3DStyle.Rotation = 5;
                barchart.ChartAreas[0].AxisX.LabelStyle.Font = new System.Drawing.Font("Ariel", 20F);
                barchart.ChartAreas[0].AxisY.LabelStyle.Font = new System.Drawing.Font("Ariel", 18F);
                DataPoint p1 = new DataPoint(0, 160);
                p1.AxisLabel = LM.Translate("guideline_Good", MessagesLabelsEN.good);
                p1.Color = Color.Yellow;
                p1.BorderColor = Color.Black;
                p1.LegendText = "SMI";


                series1.Points.Add(p1);

                DataPoint p2 = new DataPoint(0, 80);
                p2.BorderColor = Color.Black;
                p2.AxisLabel = LM.Translate("guideline_Med", MessagesLabelsEN.Medium);
                p2.Color = Color.Red;
                p2.LegendText = "SMI";
                series1.Points.Add(p2);

                DataPoint p3 = new DataPoint(0, SMI_val);
                p3.BorderColor = Color.Black;
                p3.AxisLabel = LM.Translate("smi_full", MessagesLabelsEN.SMI);
                p3.Color = Color.Blue;
                p3.LegendText = "SMI";
                series1.Points.Add(p3);
                barchart.Series["seriesname"].IsValueShownAsLabel = true;
                //this.Controls.Add(barchart);
                //barchart.AlignDataPointsByAxisLabel();


                barchart.SaveImage(Application.StartupPath + @"\SMI_chart.Png", ChartImageFormat.Png);
                string fileName = Application.StartupPath + @"\SMI_chart.Png";
                System.Drawing.Imaging.ImageFormat imageFormat = System.Drawing.Imaging.ImageFormat.Png;
                Bitmap bitmap = (Bitmap)Bitmap.FromFile(fileName);
                //this will rotate image to the left...
                bitmap.RotateFlip(RotateFlipType.Rotate270FlipNone);
                //lets save result back to file...

                bitmap.Save(fileName, imageFormat);
                bitmap.Dispose();
                //PointF firstLocation = new PointF(10f, 10f);
                //PointF secondLocation = new PointF(10f, 50f);
                //Bitmap bitmap1 = (Bitmap)System.Drawing.Image.FromFile(fileName);//load the image file

                //using (Graphics graphics = Graphics.FromImage(bitmap1))
                //{
                //    using (System.Drawing.Font arialFont = new System.Drawing.Font("Arial", 10))
                //    {
                //        graphics.DrawString("Hello", arialFont, Brushes.Blue, firstLocation);
                //        graphics.DrawString("World", arialFont, Brushes.Red, secondLocation);
                //    }
                //}
                //bitmap1.Save(fileName, imageFormat);
            }
            catch (Exception ex)
            {
                Logs.LogFile_Entries(ex.ToString() + " \tLoadSMI in " + this.FindForm().Name, "Error");

            }



        }


        public void AddTextCell(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, BaseColor bgrndcolor)
        {

            cell.HorizontalAlignment = cellHorizontalAlignment;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = ColSpan;
            cell.Rowspan = RowSpan;
            cell.BorderWidthTop = borderWidthTop;
            cell.BorderWidthBottom = borderWidthBottom;
            cell.BorderWidthLeft = borderWidthLeft;
            cell.BorderWidthRight = borderWidthRight;
            cell.PaddingTop = paddingtop;
            cell.PaddingBottom = paddingbottom;



            cell.BackgroundColor = bgrndcolor;

            table.AddCell(cell);

        }


        public void AddTextCell1(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, int paddingleft, BaseColor bgrndcolor, int cellheight, BaseColor bordercolor)
        {

            cell.HorizontalAlignment = cellHorizontalAlignment;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = ColSpan;
            cell.Rowspan = RowSpan;
            cell.BorderWidthTop = borderWidthTop;
            cell.BorderWidthBottom = borderWidthBottom;
            cell.BorderWidthLeft = borderWidthLeft;
            cell.BorderWidthRight = borderWidthRight;
            cell.PaddingTop = paddingtop;
            cell.PaddingLeft = paddingleft;
            cell.PaddingBottom = paddingbottom;
            cell.BorderColor = bordercolor;
            //cell.PaddingRight = paddingRight;
            cell.FixedHeight = cellheight;



            cell.BackgroundColor = bgrndcolor;

            table.AddCell(cell);

        }

        public void AddCellheight(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, BaseColor bgrndcolor, int paddingleft)
        {

            cell.HorizontalAlignment = cellHorizontalAlignment;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = ColSpan;
            cell.Rowspan = RowSpan;
            cell.BorderWidthTop = borderWidthTop;
            cell.BorderWidthBottom = borderWidthBottom;
            cell.BorderWidthLeft = borderWidthLeft;
            cell.BorderWidthRight = borderWidthRight;
            cell.PaddingTop = paddingtop;
            cell.PaddingRight = paddingleft;
            cell.PaddingBottom = paddingbottom;




            cell.BackgroundColor = bgrndcolor;

            table.AddCell(cell);

        }
        public void headercolor(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, BaseColor cellcolor)
        {

            cell.HorizontalAlignment = cellHorizontalAlignment;
            cell.VerticalAlignment = Element.ALIGN_MIDDLE;
            cell.Colspan = ColSpan;
            cell.Rowspan = RowSpan;
            cell.BorderWidthTop = borderWidthTop;
            cell.BorderWidthBottom = borderWidthBottom;
            cell.BorderWidthLeft = borderWidthLeft;
            cell.BorderWidthRight = borderWidthRight;
            cell.PaddingTop = paddingtop;
            cell.PaddingBottom = paddingbottom;
            cell.BackgroundColor = cellcolor;


            table.AddCell(cell);

        }

        public string RemoveSpaces(string str, char except)
        {
            try
            {
                for (int j = 0; j < str.Length; j++)
                {
                    try
                    {
                        if (str[j] == (char)(32) & (str[j - 1] != except))
                        {
                            str = str.Remove(j, 1);
                            j = j - 1;
                        }
                    }
                    catch
                    {
                        if (str[j] == (char)(32))
                        {
                            str = str.Remove(j, 1);
                            j = j - 1;
                        }
                    }

                }

            }
            catch (Exception e)
            {

            }
            return str;

        }

        // translates the test type and returns if type is frozen or not
        public bool translateType()
        {
            bool frozen = false;
            string defStr = "";
            string label = "";
            defStr = MessagesLabelsEN.WASHED;
            label = "WASHED";
            if (isFresh)
            {
                defStr = MessagesLabelsEN.FRESH;
                label = "FRESH";
            }

            else if (isFrozen)
            {
                defStr = MessagesLabelsEN.FROZEN;
                label = "FROZEN";
                frozen = true;
            }
            else if (isWashed)
            {
                defStr = MessagesLabelsEN.WASHED;
                label = "WASHED";
            }
            try
            {

                reportData[9] = LM.Translate(label, defStr);


            }
            catch (Exception ex)
            {

                Logs.LogFile_Entries(ex.ToString() + " \ttranslateType in " + this.FindForm().Name, "Error");


                //MessageBox.Show("There is no Report for this data");




            }
            return frozen;
        }




        private void SettingsButton_Click(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("Settings Button Clicked" + " \tSettingsButton_Click in " + this.FindForm().Name, "Info");
            SettingsFrm sf = new SettingsFrm(this);
            if (DEL_csv_records == true || archive_data == true)
            {
                if ((SaveButton.Enabled == true && bool.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 1) /*|| (SaveButton.Enabled == true && sf.ReportType.SelectedIndex == 0)*/)
                {
                    if (((MessageBox.Show(this, LM.Translate("SAVE_CHANGES", MessagesLabelsEN.SAVE_DATA), LM.Translate("CONFIRM_CHANGES", MessagesLabelsEN.CONFIRM_CHANGES), MessageBoxButtons.YesNo)) == DialogResult.Yes))
                    {
                        Logs.LogFile_Entries("Display Save data message to update csv", "Info");
                        updatecsv();
                        Logs.LogFile_Entries("CSV updated", "Info");
                        focus(null, null);
                        fructosetxt_Enter(null, null);
                        liquifictiontxt_Enter(null, null);
                        testerdesg_Enter(null, null);
                        testername_Enter(null, null);
                        refbydr_Enter(null, null);
                        refbydr2_Enter(null, null);
                        refbydr3_Enter(null, null);

                    }
                    else
                    {

                        patname.Text = old_patname;
                        fructosetxt.Text = old_Fru;
                        liquifictiontxt.Text = old_Liq;
                        testername.Text = old_tesname;
                        testerdesg.Text = old_tesdesg;
                        refbydr.Text = old_Refdr1;
                        refbydr2.Text = old_Refdr2;
                        refbydr3.Text = old_Refdr3;
                        SaveButton.Enabled = false;

                    }
                    //else
                    //{
                    //    ArchiveFrm archive_frm = new ArchiveFrm(mf);
                    //    archive_frm.display_csv_data();
                    //}

                }
            }

            FormCollection fc = Application.OpenForms;

            if ((Application.OpenForms["SettingsFrm"] as SettingsFrm) != null)
            {
                ///Form is already open
                MessageBox.Show(this, LM.Translate("Settings_window_open", MessagesLabelsEN.SETTINGS_WINDOW_ALREADY_OPENED), LM.Translate("SETTINGS", MessagesLabelsEN.SETTINGS));
                Logs.LogFile_Entries(MessagesLabelsEN.SETTINGS_WINDOW_ALREADY_OPENED + " \tSettingsButton_Click in " + this.FindForm().Name, "Info");
            }
            else
            {
                /// Form is not open
                sf.Show();
                Logs.LogFile_Entries("Settings Window opened" + " \tSettingsButton_Click in " + this.FindForm().Name, "Info");
            }

            sf.OnSettingsUpdate += new UpdateSettings(UpdateStatus);
            sf.OnUpdateIsSettingsOpenedFlag += new UpdateIsSettingsOpenedFlag(OnSettingsFrmClosed);

        }




        private void OnSettingsFrmClosed(object sender, EventArgs e)
        {
            // export.linecount();

            // here update appSettings from xml
            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);
            IsSettingsFrmOpened = false;


        }
        public bool translateType1()
        {
            bool frozen = false;
            string defStr = "";
            string label = "";
            defStr = MessagesLabelsEN.WASHED;
            label = "WASHED";
            if (isFresh)
            {
                defStr = MessagesLabelsEN.FRESH;
                label = "FRESH";
            }

            else if (isFrozen)
            {
                defStr = MessagesLabelsEN.FROZEN;
                label = "FROZEN";
                frozen = true;
            }
            else if (isWashed)
            {
                defStr = MessagesLabelsEN.WASHED;
                label = "WASHED";
            }
            try
            {
                if (!string.IsNullOrEmpty(patid.Text))
                    type.Text = LM.Translate(label, defStr);


            }
            catch (Exception ex)
            {


                Logs.LogFile_Entries(ex.ToString(), "Info");
                //MessageBox.Show("There is no Report for this data");




            }
            return frozen;
        }



        private void App_Start()
        {
            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);

            string[] settings = new string[15];

            try
            {
                if ((File.Exists(path) == false) || ((File.Exists(path) == true) && (File.ReadAllText(path) == String.Empty)))
                {
                    if (SerialPort.GetPortNames().Length == 0)
                        settings[0] = "None";
                    else
                        settings[0] = (SerialPort.GetPortNames())[0];
                    if (PrinterSettings.InstalledPrinters.Count == 0)
                        settings[1] = "None";
                    else
                        settings[1] = PrinterSettings.InstalledPrinters[0];
                    settings[2] = Application.StartupPath + @"\archive.csv";
                    settings[3] = "false";
                    settings[4] = "false";
                    settings[5] = "0";
                    settings[6] = "en-US"; //0
                    settings[7] = "true";
                    settings[8] = "false";
                    settings[9] = "";
                    settings[10] = "";
                    settings[11] = "1";
                    settings[12] = "false";
                    settings[13] = "";
                    settings[14] = "";
                    settings[15] = "";
                    settings[15] = "false";

                    settings[14] = Application.StartupPath + @"\settings.xml";
                    XmlUtility.WriteSettings(settings, path);
                    Logs.LogFile_Entries("Default settings created" + " \tApp_Start in " + this.FindForm().Name, "Info");
                }

                appSettings = XmlUtility.ReadSettings(path);


                switch (appSettings[6]) // save as locale string, get from file ending
                {
                    case "en-US":
                        currentLang = UniversalStrings.ENG_US;
                        break;
                    case "de-De":
                        currentLang = UniversalStrings.GER_GERMANY;
                        break;
                    case "it-IT":
                        currentLang = UniversalStrings.IT_ITALIA;
                        break;
                    case "fr-FR":
                        currentLang = UniversalStrings.FR_FRENCH;
                        break;
                    case "zh-CN":
                        currentLang = UniversalStrings.CH_CHINESE;
                        break;
                    default:
                        currentLang = CurrentCulture.ToString();
                        break;
                }
                LM.UpdateLanguage(currentLang);
                if (string.IsNullOrEmpty(appSettings[15]))
                {
                    optionallab.Text = "OPTIONAL1";
                }
                else
                {

                    optionallab.Text = appSettings[15].ToUpper();
                }
                if (string.IsNullOrEmpty(appSettings[16]))
                {
                    fructoselbl.Text = "OPTIONAL2";
                }
                else
                {

                    fructoselbl.Text = appSettings[16].ToUpper();
                }
                //LM.UpdateLanguage(appSettings[6]);

                portSelected.Text = LM.Translate("PORT", MessagesLabelsEN.PORT) + ": " + appSettings[0];

                if ((Boolean.Parse(appSettings[3]) == true && Boolean.Parse(appSettings[7]) == false) || (Boolean.Parse(appSettings[13]) == true && Boolean.Parse(appSettings[7]) == true))
                    autoprintSelected.Text = LM.Translate("AUTO_PRINT", MessagesLabelsEN.AUTO_PRINT) + ": " + LM.Translate("YES", MessagesLabelsEN.YES);
                else
                    autoprintSelected.Text = LM.Translate("AUTO_PRINT", MessagesLabelsEN.AUTO_PRINT) + ": " + LM.Translate("NO", MessagesLabelsEN.NO);

                /* if ((Boolean.Parse(appSettings[4])) == true)
                 {
                     exportSelected.Text = LM.Translate("EXPORT", MessagesLabelsEN.EXPORT) + ": " + LM.Translate("YES", MessagesLabelsEN.YES);
                     exportSelected.AutoToolTip = true;
                     exportSelected.ToolTipText = LM.Translate("EXPORT_FILE_NAME", MessagesLabelsEN.EXPORT_FILE_NAME) + appSettings[2];

                 }
                 else
                 {
                     exportSelected.Text = LM.Translate("EXPORT", MessagesLabelsEN.EXPORT) + ": " + LM.Translate("NO", MessagesLabelsEN.NO);
                     exportSelected.AutoToolTip = false;
                     exportSelected.ToolTipText = String.Empty;
                 }
                 */
                LM.GetSubControlsData(this);

                versionStatus.Text = LM.Translate("VERSION", MessagesLabelsEN.VERSION) + Application.ProductVersion;
            }

            catch (Exception ex)
            {

                MessageBox.Show(this, LM.Translate("SETTINGS_FILE_CORRUPTED_DEFAULT", MessagesLabelsEN.SETTINGS_FILE_CORRUPTED_DEFAULT), UniversalStrings.QC_GOLD_ARCHIVE);
                Logs.LogFile_Entries(MessagesLabelsEN.SETTINGS_FILE_CORRUPTED_DEFAULT + " \tApp_Start in " + this.FindForm().Name, "Error");
                File.Delete(path);
                App_Start();
            }


        }


        public void UpdateStatus()
        {
            App_Start();
            if (sp != null)
                sp.Close();
            GetData();
        }

        private void PortClosed(object sender, FormClosingEventArgs e)
        {

        }


        [DllImport("user32.dll", CharSet = CharSet.Auto, SetLastError = true)]
        private static extern IntPtr SendMessage(IntPtr hwnd, uint Msg, IntPtr wParam, IntPtr lParam);

        protected override void WndProc(ref Message message)
        {
            if (message.Msg == 0xA123)
                if (WindowState == FormWindowState.Minimized)
                    WindowState = FormWindowState.Normal;
            base.WndProc(ref message);
        }

        private void LangResxButton_Click(object sender, EventArgs e)
        {
            LM.CreateResxFromDB();
        }



        /*  private void Button1_Click(object sender, EventArgs e)
          {
              Browse openfrm = new Browse();
              openfrm.Show();
          }*/



        private void ArchiveButton_Click(object sender, EventArgs e)
        {
            export.Archive_Translate(appSettings[2]);

            Logs.LogFile_Entries("Archive Button clicked" + " \tArchiveButton_Click in " + this.FindForm().Name, "Info");
            if ((SaveButton.Enabled == true && bool.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 1) /*|| (SaveButton.Enabled == true && sf.ReportType.SelectedIndex == 0)*/)
            {
                if (((MessageBox.Show(this, LM.Translate("SAVE_CHANGES", MessagesLabelsEN.SAVE_DATA), LM.Translate("CONFIRM_CHANGES", MessagesLabelsEN.CONFIRM_CHANGES), MessageBoxButtons.YesNo)) == DialogResult.Yes))
                {
                    updatecsv();
                    focus(null, null);
                    fructosetxt_Enter(null, null);
                    liquifictiontxt_Enter(null, null);
                    testerdesg_Enter(null, null);
                    testername_Enter(null, null);
                    refbydr_Enter(null, null);
                    refbydr2_Enter(null, null);
                    refbydr3_Enter(null, null);

                }
                else
                {

                    patname.Text = old_patname;
                    fructosetxt.Text = old_Fru;
                    liquifictiontxt.Text = old_Liq;
                    testername.Text = old_tesname;
                    testerdesg.Text = old_tesdesg;
                    refbydr.Text = old_Refdr1;
                    refbydr2.Text = old_Refdr2;
                    refbydr3.Text = old_Refdr3;
                    SaveButton.Enabled = false;

                }
                //else
                //{
                //    ArchiveFrm archive_frm = new ArchiveFrm(mf);
                //    archive_frm.display_csv_data();
                //}

            }
            bool success = true;

            String filepath = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";

            ArchiveFrm openfrm = new ArchiveFrm(this);
            FileStream aFile = null;

            if (File.Exists(filepath))
            {
                try
                {
                    aFile = new FileStream(filepath, FileMode.Append, FileAccess.Read);
                    openfrm.ShowDialog();
                    Logs.LogFile_Entries("Archive Form Opened" + " \tArchiveButton_Click in " + this.FindForm().Name, "Info");


                }
                catch
                {
                    string filename;
                    filename = GetFileName(filepath);
                    bool cancelled = false;
                    //while (IsFileOpen(filename) && cancelled == false)
                    while (CanAccessFile(filepath) == false && cancelled == false)
                    {
                        this.Invoke(new Action(() =>
                        {
                            Logs.LogFile_Entries(MessagesLabelsEN.ACCESS_ERROR + " \tArchiveButton_Click in " + this.FindForm().Name, "Error");
                            if (MessageBox.Show(this, filepath + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                            {
                                Logs.LogFile_Entries(MessagesLabelsEN.Archive_Error + " \tArchiveButton_Click in " + this.FindForm().Name, "Error");
                                this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("Could_not_open", MessagesLabelsEN.Archive_Error), UniversalStrings.QC_GOLD_ARCHIVE); }));
                                cancelled = true;
                            };
                        }));
                    }

                    if (!cancelled)
                    {
                        Logs.LogFile_Entries("Archive Form Opened" + " \tArchiveButton_Click in " + this.FindForm().Name, "Info");
                        openfrm.ShowDialog();


                    }

                }
            }
            else
            {
                MessageBox.Show(LM.Translate("Archive_Empty", MessagesLabelsEN.Archive_Error_Missing), LM.Translate("", UniversalStrings.QC_GOLD_ARCHIVE));
                Logs.LogFile_Entries(MessagesLabelsEN.Archive_Error_Missing + " Received in " + this.FindForm().Name, "Error");

            }


        }





        private void Patinfoheader_OnValueChanged(object sender, EventArgs e)
        {

        }

        private void Refbtn_Click(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }
        String filePath;

        public string patid_save { get; private set; }
        public string abstinance_save { get; private set; }
        public string accession_save { get; private set; }
        public string collecteddate_save { get; private set; }
        public string received_field_save { get; private set; }
        public string type_save { get; private set; }
        public string volume_save { get; private set; }
        public string wbc_save { get; private set; }
        public string ph_save { get; private set; }
        public string conc_save { get; private set; }
        public string totalMotile_save { get; private set; }
        public string progressiveMot_save { get; private set; }
        public string nonProg_motility_save { get; private set; }
        public string immotile_save { get; private set; }
        public string mnf_save { get; private set; }
        public string msc_save { get; private set; }
        public string pmsc_save { get; private set; }
        public string fsc_save { get; private set; }
        public string velocity_save { get; private set; }
        public string smi_save { get; private set; }
        public string sperm_save { get; private set; }
        public string avg_save { get; private set; }
        public string cnt_str { get; private set; }
        public string odstr_save { get; private set; }
        public string aw_save { get; private set; }
        public string softwareV_save { get; private set; }
        public string deviceSN_save { get; private set; }
        public string motileSperm_save { get; private set; }
        public string progsperm_save { get; private set; }
        public string functionSperm_save { get; private set; }

        public void SaveButton_Click(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("Save Button clicked" + " \tSaveButton_Click in " + this.FindForm().Name, "Info");
            FileStream aFile = null;

            byte expType = 0;
            String filepath = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";

            try
            {
                aFile = new FileStream(filepath, FileMode.Append, FileAccess.Read);

            }
            catch
            {
                string filename;
                bool cancelled = false;
                filename = GetFileName(filepath);
                if (File.Exists(filepath))
                {
                    cancelled = false;

                }
                else
                {
                    cancelled = true;
                }
                //while (IsFileOpen(filename) && cancelled == false)
                exported = false;
                if (File.Exists(filename) == true)
                {


                    while (CanAccessFile(filepath) == false && cancelled == false)
                    {
                        this.Invoke(new Action(() =>
                        {
                            Logs.LogFile_Entries(MessagesLabelsEN.ACCESS_ERROR + " \tSaveButton_Click in " + this.FindForm().Name, "Error");
                            if (MessageBox.Show(this, filepath + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                            {
                                this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE); }));
                                cancelled = true;
                                Logs.LogFile_Entries(MessagesLabelsEN.IMPORT_FAILED + " \tSaveButton_Click in " + this.FindForm().Name, "Error");
                            };
                        }));
                    }
                }
                if (!cancelled && exported == false)
                {
                    if (File.Exists(filepath))
                        expType = 2;
                    else
                        expType = 1;
                    aFile = new FileStream(filepath, FileMode.Append, FileAccess.Write);

                    export.FillCsv(filepath, reportData, expType, aFile);
                }


            }

            if (!string.IsNullOrEmpty(patname.Text) || !string.IsNullOrEmpty(testername.Text) || !string.IsNullOrEmpty(testerdesg.Text) ||
         !string.IsNullOrEmpty(fructosetxt.Text) || !string.IsNullOrEmpty(liquifictiontxt.Text) || !string.IsNullOrEmpty(refbydr.Text) ||
         !string.IsNullOrEmpty(refbydr2.Text) || !string.IsNullOrEmpty(refbydr3.Text) || !string.IsNullOrEmpty(textBox2.Text) || !string.IsNullOrEmpty(accessionField.Text) || !string.IsNullOrEmpty(absField.Text)
         || !string.IsNullOrEmpty(volField.Text) || !string.IsNullOrEmpty(typeField.Text) || !string.IsNullOrEmpty(phField.Text) || !string.IsNullOrEmpty(wbcField.Text) || !string.IsNullOrEmpty(textBox1.Text)
         || !string.IsNullOrEmpty(richTextBox1.Text) || !string.IsNullOrEmpty(patidField.Text))
            {

                updatecsv();
                duplicateData = "";
                focus(null, null);
                fructosetxt_Enter(null, null);
                liquifictiontxt_Enter(null, null);
                testerdesg_Enter(null, null);
                testername_Enter(null, null);
                refbydr_Enter(null, null);
                refbydr2_Enter(null, null);
                refbydr3_Enter(null, null);

                check = true;
                //if (DEL_csv_records == true)
                //{
                //    updatecsv();
                //    check = true;
                //}

                //if (DEL_csv_records == false)
                //{
                //    updatecsv();
                //    check = true;
                //}
            }






        }


        public void getdata()
        {
            try
            {
                if (sp.IsOpen == false)
                {
                    sp = new SerialPort(appSettings[0]);
                    sp.BaudRate = 19200;
                    sp.DataBits = 8;
                    sp.StopBits = StopBits.One;
                    sp.Parity = Parity.None;
                    sp.Handshake = Handshake.None;
                    sp.Open();
                    sp.DataReceived += Recieved;
                }
            }
            catch (Exception ex)
            {

                /// MessageBox.Show(this, LM.Translate("PORT_ERROR", MessagesLabelsEN.PORT_ERROR), LM.Translate("PORT_ERROR_CAPTION", MessagesLabelsEN.PORT_ERROR_CAPTION));
            }
        }






        private void Fructose_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void Liquifiction_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void Testername_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void Testerdesg_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void MenuStrip_ItemClicked(object sender, ToolStripItemClickedEventArgs e)
        {

        }

        private void Rptname_Click(object sender, EventArgs e)
        {

        }


        public void MainForm_Load(object sender, EventArgs e)
        {
            placeholder();
            label63.Visible = false;
            textBox13.Visible = false;
            panel1.Enabled = false;
            Logs.LogFile();
            Logs.LogFile_Entries("App started", "Info");
            Logs.LogFile_Entries(Application.ProductVersion, "Info");
            Logs.LogFile_Entries("Current Language :" + currentLang + "", "Info");
            GetData();
            //Logs.LogFile();
            refreshfrm();
            //export.linecount();


            // Makes sure serial port is open before trying to write 


            stateOfFields(true);



        }
        public void stateOfFields(bool isErase)
        {
            if (!checkBox2.Checked)
            {
                patidField.Visible = false;
                testdateField1.Visible = false;
                collecteddateField.Visible = false;
                receivedDateField.Visible = false;
                absField.Visible = false;
                accessionField.Visible = false;
                volField.Visible = false;
                phField.Visible = false;
                wbcField.Visible = false;
                typeField.Visible = false;
                if (isErase)
                {
                    EmptyFields();
                }

                ArchiveButton.Enabled = true;
            }
            else
            {

                ArchiveButton.Enabled = false;
                //SaveButton.Enabled = false;
                patidField.Visible = true;
                testdateField1.Visible = true;
                collecteddateField.Visible = true;
                receivedDateField.Visible = true;
                absField.Visible = true;
                accessionField.Visible = true;
                volField.Visible = true;
                phField.Visible = true;
                wbcField.Visible = true;
                typeField.Visible = true;
                foreach (Control c in Controls)
                {
                    if (c is TextBox)
                    {
                        c.Enabled = true;
                    }
                }

            }
        }

        public void refreshfrm()
        {

            pictureBox2.Image = imageList1.Images[1];
            comments.Visible = false;
            Avg_desg.Visible = false;
            Avg_frclbl.Visible = false;
            Avg_liqlbl.Visible = false;
            Avg_patname.Visible = false;
            Avg_tesname.Visible = false;
            Avg_desg.Text = "";
            Avg_patname.Text = "";
            Avg_frclbl.Text = "";
            Avg_liqlbl.Text = "";
            Avg_tesname.Text = "";
            Avg_desg.Visible = false;
            refbydrlbl.Text = LM.Translate("REF_BY_DR", MessagesLabelsEN.REF_BY_DR).TrimEnd(':');
            SettingsButton.Enabled = true;
            Resultinfo.Text = LM.Translate("RESULT_INFORMATION", MessagesLabelsEN.RESULT_INFO);

            // List<System.Drawing.Image> images = new List<System.Drawing.Image>();

            // images.Add(System.Drawing.Image.FromFile(Application.StartupPath + "\report1.png"));
            //images.Add(System.Drawing.Image.FromFile(Application.StartupPath + "\report.png"));


            if (Boolean.Parse(appSettings[7]) == false)
            {
                Logs.LogFile_Entries("Print result strip UI displayed" + " \trefreshfrm in " + this.FindForm().Name, "Info");
                resultstrip.Text = LM.Translate("PRINT_RESULTS_STRIP", MessagesLabelsEN.PRINT_TEST_RESULTS);
                PrintButton.Image = System.Drawing.Image.FromFile(Application.StartupPath + "\\assets\\report1.png");


                ArchiveButton.Enabled = false;
                // pictureBox1.Location = new Point(152, 56)

                //pictureBox1.Image = imageList1.Images[2];
                foreach (Control c in Controls)
                {
                    if (c is TextBox || c is Button || c is Label || c is RichTextBox)
                    {
                        c.Visible = false;
                        resultstrip.Visible = true;
                    }

                }

            }
            else
            {

                ArchiveButton.Enabled = true;
                foreach (Control c in Controls)
                {
                    if (c is TextBox || c is Button || c is Label)
                    {
                        c.Visible = true;
                        resultstrip.Visible = false;
                    }

                }

            }
            if (appSettings[6] == "zh-CN")
            {
                devicesntxt.Location = new Point(80, 722);
                devicesn.Location = new Point(150, 722);
            }
            else
            {
                devicesntxt.Location = new Point(168, 816);
                devicesn.Location = new Point(256, 816);
            }
            if (appSettings[6] == "de-De" || appSettings[6] == "fr-FR")
            {
                testername.Location = new Point(178, 478);

            }
            if (Boolean.Parse(appSettings[7]) == true)
            {
                Plot.Visible = false;
                refbydr2.Visible = false;
                refbydr3.Visible = false;
                PrintButton.Text = "Export pdf";

            }

            if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 0)
            {
                Logs.LogFile_Entries("Basic Report UI displayed" + " \trefreshfrm in " + this.FindForm().Name, "Info");
                rptname.Text = LM.Translate("BASIC_REPORT", MessagesLabelsEN.BASIC_REPORT);
                pictureBox1.Image = imageList1.Images[1];
                // pictureBox1.Location = new Point(125, 58);
                fructoselbl.Visible = false;
                fructose.Visible = false;
                liquifiction.Visible = false;
                liquilbl.Visible = false;
                refbydr.Visible = false;
                refbtn.Visible = false;
                label40.Visible = false;
                refbydrlbl.Visible = false;
                patname.Visible = false;
                patnamelbl.Visible = false;
                fructosetxt.Visible = false;
                liquifictiontxt.Visible = false;
                refbydr2.Visible = false;
                refbydr3.Visible = false;
                Testerinfo.Visible = false;
                richTextBox1.Visible = false;
                optionallab.Visible = false;
                textBox1.Visible = false;
                textBox2.Visible = false;
                label15.Visible = false;
                label16.Visible = false;
                label1.Visible = false;
                /// Plot.Visible = false;
                //tesname.Visible = false;
                testernamelbl.Visible = false;
                tesdesg.Visible = false;
                testername.Visible = false;
                testerdesg.Visible = false;
                Resultinfo.Location = new Point(27, 415);
                label42.Location = new Point(27, 440);
                conc.Location = new Point(190, 455);
                label4.Location = new Point(27, 455);
                label5.Location = new Point(27, 480);
                totmot.Location = new Point(190, 480);
                rapidProglable.Location = new Point(27, 505);
                progmotility.Location = new Point(190, 505);


                label21.Location = new Point(27, 530);
                nonprgmot.Location = new Point(190, 530);

                label20.Location = new Point(27, 553);
                immotility.Location = new Point(190, 553);
                label19.Location = new Point(27, 575);
                morph.Location = new Point(190, 575);

                label27.Location = new Point(284, 455);
                msc.Location = new Point(455, 455);

                rapidpmscLable.Location = new Point(284, 480);
                pmsc.Location = new Point(455, 480);
                label25.Location = new Point(284, 505);
                fsc.Location = new Point(455, 505);

                label24.Location = new Point(284, 530);
                velocity.Location = new Point(455, 530);
                label23.Location = new Point(284, 553);
                smi.Location = new Point(455, 553);

                label31.Location = new Point(510, 455);
                sperm.Location = new Point(697, 455);
                label30.Location = new Point(510, 480);
                motsperm.Location = new Point(697, 480);

                label29.Location = new Point(510, 505);
                progsperm.Location = new Point(697, 505);
                label28.Location = new Point(510, 530);
                funcsperm.Location = new Point(697, 530);
                label22.Location = new Point(510, 555);
                label2.Location = new Point(697, 555);


            }

            else if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 1)
            {
                Logs.LogFile_Entries("Advanced Report UI displayed" + " \trefreshfrm in " + this.FindForm().Name, "Info");
                rptname.Text = LM.Translate("ADVANCED_REPORT", MessagesLabelsEN.ADVANCED_REPORT);
                // pictureBox1.Location = new Point(155, 58);
                pictureBox1.Image = imageList1.Images[0];



            }
            else if (Boolean.Parse(appSettings[7]) == false && resultstrip.Enabled == true)
            {
                pictureBox1.Image = imageList1.Images[2];
                // pictureBox1.Location = new Point(118, 57);
            }

            if (int.Parse(appSettings[11]) == 1 && string.IsNullOrEmpty(Testdate.Text))
            {
                foreach (Control c in Controls)
                {
                    if (c is TextBox)
                    {
                        c.Enabled = false;
                    }

                }
                //patname.Enabled = false;
                //fructosetxt.Enabled = false;
                //liquifictiontxt.Enabled = false;
                //testername.Enabled = false;
                //testerdesg.Enabled = false;
                //refbydr.Enabled = false;
                //refbydr2.Enabled = false;
                //refbydr3.Enabled = false;
                //textBox1.Enabled = false;
            }
            foreach (Control c in Controls)
            {
                if (c is TextBox)
                {
                    c.Enabled = true;

                }

            }


        }
        private void PictureBox1_Click(object sender, EventArgs e)
        {

        }

        private void Label15_Click(object sender, EventArgs e)
        {

        }



        private void OpenFileButton_Click(object sender, EventArgs e)
        {
            LoadMorph();
            //advancedrpt();

        }




        private void refbydr2_TextChanged(object sender, EventArgs e)
        {

        }

        private void Refbtn_MouseClick(object sender, MouseEventArgs e)
        {

        }

        private void PortClosed(object sender, FormClosedEventArgs e)
        {
            if (sp != null)
            {
                if (sp.IsOpen == true)
                    sp.Close();
            }
        }

        private void MainForm_FormClosing(object sender, FormClosingEventArgs e)
        {
            System.Windows.Forms.Application.Exit();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            string csvPath = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";
            ExportExcel(csvPath);
            //advancedrpt();
        }

        private void Resxdb_Click(object sender, EventArgs e)
        {
            LM.CreateResxFromDB();
        }

        private void Refbtn_Click_1(object sender, EventArgs e)
        {
            counter++;
            if (counter == 1)
            {
                refbydr2.Visible = true;

            }
            else if (counter == 2)
            {
                refbydr3.Visible = true;
                counter = 0;

                //refbtn.Text = "-";
            }
            //    else if (counter == 3)
            //    {
            //        refbydr3.Visible = false;

            //    }
            //    else if (counter == 4)
            //    {
            //        refbydr2.Visible = false;
            //        refbtn.Text = "+";
            //        counter = 0;
            //    }

        }

        private void Patname_TextChanged_1(object sender, EventArgs e)
        {
            patname.MaxLength = 20;
            savebtn();
            svbtndisable();
            if (refbydr.Visible == false)
            {
                SaveButton.Enabled = false;
            }
        }

        private void Fructosetxt_TextChanged_1(object sender, EventArgs e)
        {
            fructosetxt.MaxLength = 10;

            savebtn();
            svbtndisable();
        }

        private void Liquifictiontxt_TextChanged_1(object sender, EventArgs e)
        {
            liquifictiontxt.MaxLength = 10;
            savebtn();
            svbtndisable();
        }

        private void Testername_TextChanged_2(object sender, EventArgs e)
        {
            testername.MaxLength = 25;
            savebtn();
            svbtndisable();
        }

        private void Testerdesg_TextChanged_2(object sender, EventArgs e)
        {

            testerdesg.MaxLength = 25;
            savebtn();
            svbtndisable();
        }

        private void Refbydr_TextChanged_1(object sender, EventArgs e)
        {
            refbydr.MaxLength = 20;
            savebtn();
            svbtndisable();
        }

        private void Rptname_Click_1(object sender, EventArgs e)
        {

        }

        private void refbydr2_TextChanged_1(object sender, EventArgs e)
        {
            refbydr2.MaxLength = 20;
            savebtn();
        }

        private void refbydr3_TextChanged(object sender, EventArgs e)
        {
            refbydr3.MaxLength = 20;
            savebtn();
        }

        private void AutoprintSelected_Click(object sender, EventArgs e)
        {

        }


        private void patname_Enter(object sender, EventArgs e)
        {

        }

        public void focus(object sender, EventArgs e)
        {
            old_patname = string.Empty;
            old_patname = patname.Text;
        }

        public void fructosetxt_Enter(object sender, EventArgs e)
        {
            old_Fru = string.Empty;
            old_Fru = fructosetxt.Text;
        }

        public void liquifictiontxt_Enter(object sender, EventArgs e)
        {
            old_Liq = string.Empty;
            old_Liq = liquifictiontxt.Text;
        }

        public void testername_Enter(object sender, EventArgs e)
        {
            old_tesname = string.Empty;
            old_tesname = testername.Text;
        }

        public void testerdesg_Enter(object sender, EventArgs e)
        {
            old_tesdesg = string.Empty;
            old_tesdesg = testerdesg.Text;
        }

        public void refbydr_Enter(object sender, EventArgs e)
        {
            old_Refdr1 = string.Empty;
            old_Refdr1 = refbydr.Text;
        }

        public void refbydr2_Enter(object sender, EventArgs e)
        {
            old_Refdr2 = string.Empty;
            old_Refdr2 = refbydr2.Text;
        }

        public void refbydr3_Enter(object sender, EventArgs e)
        {
            old_Refdr3 = string.Empty;
            old_Refdr3 = refbydr3.Text;
        }

        private void MainForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            Logs.LogFile_Entries("App closed", "Info");
        }

        private void Patname_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
            //else if(ch == 08)
            //{
            //    e.Handled = false;
            //}
        }

        private void Fructosetxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void Liquifictiontxt_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void Testername_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void Testerdesg_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void Refbydr_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void Refbydr2_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void Refbydr3_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (ch == 44 || ch == 59)
            {
                e.Handled = true;
            }
        }

        private void comments_Leave(object sender, EventArgs e)
        {

        }

        private void comments_TextChanged(object sender, EventArgs e)
        {
            //Avg_data.Add(comments.Text);
            comments.MaxLength = 350;
        }

        private void comments_KeyUp(object sender, KeyEventArgs e)
        {

        }

        private void comments_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void comments_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyData & Keys.KeyCode)
            {

                case Keys.Enter:
                    max_lines++;
                    if (max_lines >= 7)
                    {
                        e.Handled = true;
                        e.SuppressKeyPress = true;
                    }
                    break;
            }



        }

        private void RichTextBox1_TextChanged(object sender, EventArgs e)
        {
            richTextBox1.MaxLength = 250;
        }

        private void TextBox1_TextChanged(object sender, EventArgs e)
        {
            textBox1.MaxLength = 20;
            savebtn();
            svbtndisable();
        }

        private void RichTextBox1_KeyDown(object sender, KeyEventArgs e)
        {
            switch (e.KeyData & Keys.KeyCode)
            {

                case Keys.Enter:
                    e.Handled = true;
                    e.SuppressKeyPress = true;

                    break;
            }
        }

        private void TextBox12_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void TextBox11_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void TextBox8_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void TextBox10_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void TextBox9_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void TextBox14_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }

        private void TextBox13_KeyPress(object sender, KeyPressEventArgs e)
        {
            char ch = e.KeyChar;
            if (!char.IsDigit(ch) && ch != 8)
            {
                Logs.LogFile_Entries(MessagesLabelsEN.Numerical_Error + "  \tHeaderSpace_KeyPress in" + this.FindForm().Name, "Error");
                MessageBox.Show(LM.Translate("Enter_numerical", MessagesLabelsEN.Numerical_Error), LM.Translate("Enter_value", MessagesLabelsEN.SETTINGS_ERROR));
                e.Handled = true;
            }
        }


        public void morph_chart(int nforms, int head_defects, int pinheads, int neck, int tail_defects, int cytoplasmicdroplets, int acrosome)
        {



            byte[] newImage = new byte[0];
            try
            {
                //Store chart properties
                var chart = new Chart();
                chart.Height = 1500;
                chart.Width = 1500;



                //string title = LM.Translate("MOTILITY", MessagesLabelsEN.Motility);
                //chart.Titles.Add(title);
                //chart.Titles[0].Font = new System.Drawing.Font("HELVITICA", 40f, FontStyle.Bold);
                //chart.Titles[0].Position.X = 40;
                //chart.Titles[0].Position.Y = 2;
                //chart.Titles[0].Position.Height = 45;
                //chart.Titles[0].Position.Width = 45;


                // iTextSharp.text.Font headers = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 11, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);


                var chartArea1 = new ChartArea();
                chart.ChartAreas.Add(chartArea1);
                chart.ChartAreas[0].AlignmentStyle = AreaAlignmentStyles.All;
                Series series1;


                string seriesName1 = " ";



                //if(immotility.Text == "---")
                //{
                //    int c = 0;

                //}

                //else
                //{
                //    int c = int.Parse(immotility.Text);
                //}

                series1 = new Series();
                seriesName1 = "PIE Chart";
                series1.Name = seriesName1;
                series1.ChartType = SeriesChartType.Pie;

                chart.Series.Add(series1);

                chart.Legends.Add(new Legend("Legend2"));
                chart.Series[seriesName1].Legend = "Legend2";
                chart.Series[seriesName1].IsVisibleInLegend = true;
                //chart.Legends["Legend2"].Position.Auto = false;
                //chart.Legends["Legend2"].Position = new ElementPosition(-5, 60, 90, 10);
                //chart.Legends["Legend2"].LegendItemOrder = LegendItemOrder.Auto;
                //chart.Series[seriesName1].ShadowColor = Color.Blue;
                //chart.Series[seriesName1].ShadowOffset = 5;

                chart.Legends["Legend2"].Position = new ElementPosition(65, 10, 40, 35);
                chart.Legends["Legend2"].Font = new System.Drawing.Font("Arial", 20F, FontStyle.Bold);
                chart.ChartAreas[0].Position.Auto = false;
                chart.ChartAreas[0].Position.X = 0;
                chart.ChartAreas[0].Position.Y = 0;
                chart.ChartAreas[0].Area3DStyle.Enable3D = true;
                chart.ChartAreas[0].Area3DStyle.Inclination = 45;
                chart.ChartAreas[0].Position.Height = 60;
                chart.ChartAreas[0].Position.Width = 70;
                chart.ChartAreas[0].ShadowColor = Color.Blue;







                //immotile = int.Parse(refbydr.Text);
                //prog = int.Parse(refbydr2.Text);
                //nonprog = int.Parse(refbydr3.Text);



                series1.Points.AddXY(0, nforms);
                series1.Points[0].Color = Color.FromArgb(11, 87, 189);
                //chart.Series[seriesName1]["PieLabelStyle"] = "Disabled";
                chart.Series[seriesName1].BorderColor = Color.Black;
                chart.Series[seriesName1].BorderWidth = 0;
                chart.Series[seriesName1].ShadowColor = Color.Gray;
                chart.Series[seriesName1].Points[0].Label = /*LM.Translate("IMMOTILITY", MessagesLabelsEN.IMMOTILITY_TEXT) +*/ "(#VALY%)";
                chart.Series[seriesName1].Points[0].Font = new System.Drawing.Font("Arial", 25F);

                chart.Series[seriesName1].Points[0].LegendText = "Normal Forms";
                //chart.Series[seriesName1].Points[0].sty = LegendImageStyle.Marker;




                series1.Points.AddXY(1, head_defects);
                // chart.Series[seriesName1].Label = "#PERCENT\n#VALY";
                series1.Points[1].Color = Color.FromArgb(203, 12, 12);
                //chart.Series[seriesName1]["PieLabelStyle"] = "Disabled";
                chart.Series[seriesName1].BorderColor = Color.Black;
                chart.Series[seriesName1].BorderWidth = 0;
                chart.Series[seriesName1].ShadowColor = Color.Gray;
                chart.Series[seriesName1].Points[1].Label = /*LM.Translate("IMMOTILITY", MessagesLabelsEN.IMMOTILITY_TEXT) +*/ "(#VALY%)";
                chart.Series[seriesName1].Points[1].Font = new System.Drawing.Font("Arial", 25F);
                chart.Series[seriesName1].Points[1].LegendText = "Head Defectives";

                //if (pinheads!=0)
                //{
                //    series1.Points.AddXY(2, pinheads);
                //    // chart.Series[seriesName1].Label = "#PERCENT\n#VALY";
                //    series1.Points[2].Color = Color.FromArgb(205, 20, 205);
                //    //chart.Series[seriesName1]["PieLabelStyle"] = "Disabled";
                //    chart.Series[seriesName1].BorderColor = Color.Black;
                //    chart.Series[seriesName1].BorderWidth = 0;
                //    chart.Series[seriesName1].ShadowColor = Color.Gray;
                //    chart.Series[seriesName1].Points[2].Label = /*LM.Translate("IMMOTILITY", MessagesLabelsEN.IMMOTILITY_TEXT) +*/ "(#VALY%)";
                //    chart.Series[seriesName1].Points[2].Font = new System.Drawing.Font("Arial", 25F);
                //    chart.Series[seriesName1].Points[2].LegendText = "Pin Heads";
                //}


                series1.Points.AddXY(2, neck);
                series1.Points[2].Color = Color.FromArgb(104, 205, 20);
                //chart.Series[seriesName1]["PieLabelStyle"] = "Disabled";
                chart.Series[seriesName1].BorderColor = Color.Black;
                chart.Series[seriesName1].BorderWidth = 0;
                chart.Series[seriesName1].ShadowColor = Color.Gray;
                chart.Series[seriesName1].Points[2].Label = "(#VALY%)";
                chart.Series[seriesName1].Points[2].Font = new System.Drawing.Font("Arial", 25F);
                chart.Series[seriesName1].Points[2].LegendText = "Neck/Mid Piece";

                series1.Points.AddXY(3, tail_defects);
                // chart.Series[seriesName1].Label = "#VALX";
                chart.Series[seriesName1].Points[3].Color = Color.FromArgb(238, 120, 9);
                chart.Series[seriesName1].Points[3].Label = "(#VALY%)";
                chart.Series[seriesName1].Points[3].LegendText = "Tail Defectives";
                chart.Series[seriesName1].Points[3].Font = new System.Drawing.Font("Arial", 25F);
                chart.Series[seriesName1].BorderColor = Color.Black;
                chart.Series[seriesName1].BorderWidth = 0;
                chart.Series[seriesName1].ShadowColor = Color.Gray;

                series1.Points.AddXY(4, cytoplasmicdroplets);
                // chart.Series[seriesName1].Label = "#VALX";
                chart.Series[seriesName1].Points[4].Color = Color.FromArgb(203, 55, 102);
                chart.Series[seriesName1].Points[4].Label = "(#VALY%)";
                chart.Series[seriesName1].Points[4].LegendText = "Cytoplasmic Droplets";
                chart.Series[seriesName1].Points[4].Font = new System.Drawing.Font("Arial", 25F);

                //series1.Points.AddXY(5, acrosome);

                //series1.Points[5].Color = Color.FromArgb(62, 226, 216);
                ////chart.Series[seriesName1]["PieLabelStyle"] = "Disabled";
                //chart.Series[seriesName1].BorderColor = Color.Black;
                //chart.Series[seriesName1].BorderWidth = 0;
                //chart.Series[seriesName1].ShadowColor = Color.Gray;
                //chart.Series[seriesName1].Points[5].Label = /*LM.Translate("IMMOTILITY", MessagesLabelsEN.IMMOTILITY_TEXT) +*/ "(#VALY%)";
                //chart.Series[seriesName1].Points[5].Font = new System.Drawing.Font("Arial", 25F);
                //chart.Series[seriesName1].Points[5].LegendText = "Acrosome";
                //chart.Series[seriesName1].Points[0].sty = LegendImageStyle.Marker;
                if (textBox12.Text == 0.ToString())
                {
                    textBox12.Text = "";
                }
                if (textBox11.Text == 0.ToString())
                {
                    textBox11.Text = "";
                }
                if (textBox8.Text == 0.ToString())
                {
                    textBox8.Text = "";
                }
                if (textBox10.Text == 0.ToString())
                {
                    textBox10.Text = "";
                }
                if (textBox9.Text == 0.ToString())
                {
                    textBox9.Text = "";
                }
                if (textBox14.Text == 0.ToString())
                {
                    textBox14.Text = "";
                }
                if (textBox13.Text == 0.ToString())
                {
                    textBox13.Text = "";
                }

                series1.IsVisibleInLegend = true;
                chart.Series[seriesName1].Font = new System.Drawing.Font("Arial", 8F, FontStyle.Bold);
                chart.SaveImage(Application.StartupPath + @"\halo_chart.Png", ChartImageFormat.Png);
                //string fss = Application.StartupPath +  "Morphology_chart" + ".jpg";
                //System.Drawing.Image MorpImg = System.Drawing.Image.FromFile(fss);

                ////////////////////////////////////////////////////////////////////////////////////



            }
            catch (Exception ex)
            {

                MessageBox.Show(ex.ToString());

            }



        }

        private void TextBox12_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void Opt_Enter(object sender, EventArgs e)
        {
            if (opt.Text == "Optional Field")
            {
                opt.Text = "";
                opt.ForeColor = Color.Black;
            }
        }

        private void CheckBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (checkBox1.Checked == true)
            {
                panel1.Enabled = true;
            }
            else
            {
                panel1.Enabled = false;
            }
        }

        private void CheckBox2_CheckedChanged(object sender, EventArgs e)
        {
            if (reportData.Count != 0)
            {
                stateOfFields(false);
            }
            else
            {
                stateOfFields(true);

            }

        }

        private void TextBox19_TextChanged(object sender, EventArgs e)
        {

        }

        private void TestdateField_ValueChanged(object sender, EventArgs e)
        {

        }

        private void concbox_TextChanged(object sender, EventArgs e)
        {

        }

        private void patidField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox2_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void accessionField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void volField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void typeField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void phField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void wbcField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void Plot_TextChanged(object sender, EventArgs e)
        {

        }

        private void absField_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox16_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox7_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox6_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox3_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox5_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox4_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox15_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox11_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox8_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox10_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox9_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox14_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox13_TextChanged(object sender, EventArgs e)
        {
            savebtn();
            svbtndisable();
        }

        private void textBox17_TextChanged(object sender, EventArgs e)
        {

        }

        private void label28_Click(object sender, EventArgs e)
        {

        }

        private void motspermbox_TextChanged(object sender, EventArgs e)
        {

        }
    }
}