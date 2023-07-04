using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Resources;
using System.IO;
using System.Windows.Forms;
using System.Globalization;
using System.Data;
using System.Diagnostics;
using System.ComponentModel;
using System.Data.SqlClient;
//using Microsoft.SqlServer.Management;
//using Microsoft.SqlServer.Management.Smo;
using System.Data.OleDb;


namespace QcGoldArchive
{
    class LanguageManagement
    {
        protected static LanguageManagement LM;
        ResourceManager resMan;
        
        private CultureInfo culture;
        private CultureInfo defCulture;
        private string[] appSettings;
        ToolTip ctl1;
        OleDbConnection conAccess;

        protected LanguageManagement()
        {
            resMan = ResourceManager.CreateFileBasedResourceManager("VResources", Application.StartupPath + "\\Language", null);
        }

        public static LanguageManagement CreateInstance()
        {
            if (LM == null)
                LM = new LanguageManagement();

            return LM;
        }

        public string Translate(string TextCode, string DefText)
        {
            string tempStr = null;
            //CultureInfo tmpCult = new CultureInfo(StrCult);
         //   tempStr = resMan.GetString(TextCode, culture);

            if (tempStr == "" || tempStr == null)
                tempStr = DefText;

            return tempStr;
        }







      
        public string[] Translate(string[] TextCodeArr)
        {
            string tempStr;
            string[] tempStrArr = new string[TextCodeArr.Length];
            int i = 0;
            foreach (string item in TextCodeArr)
            {
                tempStr = resMan.GetString(item, culture);
                if (tempStr == "" || tempStr == null)
                    tempStr = resMan.GetString(item, defCulture);
                tempStrArr[i] = tempStr;
                i++;
            }
                        
            return tempStrArr;
        }

        public void UpdateLanguage(string StrCult)
        {
            if (StrCult != null)
                culture = CultureInfo.CreateSpecificCulture(StrCult);
            else
                culture = CultureInfo.CreateSpecificCulture("en");
            defCulture = CultureInfo.CreateSpecificCulture("en");
        }

        public void GetSubControlsData(Form MainCtl)
        {
            SetControl(MainCtl);
            foreach (Control ctl in MainCtl.Controls)
            {
                //SetControl(ctl);
                if (ctl is Panel || ctl is GroupBox || ctl is StatusStrip)
                    GetSubControlsData(ctl);
                else if (ctl is ToolStrip)
                    GetToolStripItems((ToolStrip)ctl);
                else
                    SetControl(ctl);
            }
        }

        private void GetToolStripItems(ToolStrip ToolStr)
        {
            foreach (ToolStripItem tsi in ToolStr.Items)
            {
                SetControl(tsi);
            }
        }

        public void GetSubControlsData(Control MainCtl)
        {
            SetControl(MainCtl);
            foreach (Control ctl in MainCtl.Controls)
            {
                //SetControl(ctl);
                if (ctl is Panel || ctl is GroupBox || ctl is StatusStrip)
                    GetSubControlsData(ctl);
                ///else if (ctl is ToolStrip)
                // /   GetToolStripItems((ToolStrip)ctl);
                else
                    SetControl(ctl);
            }
        }
        


            private void SetControl(Control ctl)
             {
            
            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);
           
            switch (ctl.Name.ToString())
            {
                //case "generate":
                //    ctl1.SetToolTip(ctl, LM.Translate("ARCHIVE_HEADING", MessagesLabelsEN.ARCHIVE_HEADING)) ;
                //    break;
                case "Archivehead":
                    ctl.Text = LM.Translate("ARCHIVE_HEADING", MessagesLabelsEN.ARCHIVE_HEADING);
                    break;
                //case "generate":
                //    ctl1.ToolTipText = LM.Translate("CLEAR", MessagesLabelsEN.CLEAR);
                //    break;


                case "label122":
                    ctl.Text = LM.Translate("SEARCH_BY", MessagesLabelsEN.SEARCH_BY);
                    break;
                case "btnFilter":
                    ctl.Text = LM.Translate("SEARCH", MessagesLabelsEN.SEARCH);
                    break;
                case "btnShowAll":
                    ctl.Text = LM.Translate("CLEAR_FILTER", MessagesLabelsEN.CLEAR_FILTER);
                    break;
              
                case "Arc_cancel":
                    ctl.Text = LM.Translate("CANCEL", MessagesLabelsEN.CANCEL);
                    break;
                case "ArchiveFrm":
                    ctl.Text = LM.Translate("ARCHIVE_HEADING", MessagesLabelsEN.ARCHIVE_HEADING);
                    break;

                case "label17":
                     ctl.Text = LM.Translate("PATIENT_ID", MessagesLabelsEN.PATIENT_ID).TrimEnd(':'); 
                    break;

                case "patnamelbl":
                    ctl.Text = LM.Translate("PAT_NAME", MessagesLabelsEN.PATIENT_NAME).TrimEnd(':');
                    break;
                case "label36":
                    ctl.Text = LM.Translate("PATIENT_INFO", MessagesLabelsEN.PATIENT_INFO);
                    break;


                case "rpttype":
                    ctl.Text = LM.Translate("Report_type", MessagesLabelsEN.Report_type);
                    break;
                case "Testerinfo":
                    ctl.Text = LM.Translate("Tester_Information", MessagesLabelsEN.TESTER_INFO);
                    break;


                case "usedefault":
                    ctl.Text = LM.Translate("Use_Default", MessagesLabelsEN.Use_Default);
                    break;

                case "rpttitle":
                    ctl.Text = LM.Translate("Report_title", MessagesLabelsEN.Report_title);
                    break;

                case "sysinfo":
                    ctl.Text = LM.Translate("System_information", MessagesLabelsEN.sys_info_label);
                    break;

                case "devicesntxt":
                    ctl.Text = LM.Translate("DEVICE_SN", MessagesLabelsEN.DEVICE_SN);
                    break;

                case "refbydrlbl":
                    ctl.Text = LM.Translate("REF_BY_DR", MessagesLabelsEN.REF_BY_DR).TrimEnd(':');
                    break;

                case "label6":
                    ctl.Text = LM.Translate("TEST_DATE", MessagesLabelsEN.TEST_DATE);
                    break;

                case "label9":
                    ctl.Text = LM.Translate("ACCESSION", MessagesLabelsEN.ACCESSION);
                    break;

                case "rptname":
                   
                    if (bool.Parse(appSettings[7])==true && int.Parse(appSettings[11])==1 )
                    ctl.Text = LM.Translate("ADVANCED_REPORT", MessagesLabelsEN.ADVANCED_REPORT);
                    
                    else if(bool.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 0)
                        ctl.Text = LM.Translate("BASIC_REPORT", MessagesLabelsEN.BASIC_REPORT);
                    break;
                case "label10":
                    ctl.Text = LM.Translate("COLLECTED_DATE", MessagesLabelsEN.COLLECTED_DATE);
                    break;
                case "label11":
                    ctl.Text = LM.Translate("VOLUME_WU", MessagesLabelsEN.VOLUME);
                    break;
                case "label12":
                    ctl.Text = LM.Translate("TYPE", MessagesLabelsEN.TYPE);
                    break;

                case "label22":
                    ctl.Text = LM.Translate("Home_morph_sperm", MessagesLabelsEN.MORPH_NORM_FORMS);
                    break;


                case "titlelab":
                    ctl.Text = LM.Translate("Title", MessagesLabelsEN.Title);
                    break;

              

                case "subtitle":
                    ctl.Text = LM.Translate("Sub_Title", MessagesLabelsEN.Sub_Title);
                    break;

                case "label13":
                    ctl.Text = LM.Translate("PH", MessagesLabelsEN.PH);
                    break;

                case "addsign":
                    ctl.Text = LM.Translate("ADD_SIGNATURE", MessagesLabelsEN.ADD_SIGNATURE);
                    break;

                case "label14":
                    ctl.Text = LM.Translate("WBC_CONC", MessagesLabelsEN.WBC_CONC);
                    break;


                //case "fructoselbl":
                //    ctl.Text = LM.Translate("FRUCTOSE", MessagesLabelsEN.FRUCTOSE);
                //    break;

                case "label39":
                    ctl.Text = LM.Translate("SAMPLE_INFO", MessagesLabelsEN.SAMPLE_INFO);
                    break;

                case "label27":
                    ctl.Text = LM.Translate("MSC_WU", MessagesLabelsEN.MSC_WU);
                    break;

                case "label26":
                    ctl.Text = LM.Translate("PMSC_WU", MessagesLabelsEN.PMSC_WU);
                    break;

                case "label23":
                    ctl.Text = LM.Translate("SMI", MessagesLabelsEN.SMI);
                    break;

                case "label31":
                    ctl.Text = LM.Translate("NUM_SPERM_WU", MessagesLabelsEN.SPERMS);
                    break;

                case "label30":
                    ctl.Text = LM.Translate("MOTILE_SPERM_WU", MessagesLabelsEN.MOT_SPERMS);
                    break;

                case "label29":
                    ctl.Text = LM.Translate("PROG_SPERM_WU", MessagesLabelsEN.PROG_SPERMS);
                    break;

                case "label28":
                    ctl.Text = LM.Translate("FUNC_SPERM_WU", MessagesLabelsEN.FUNC_SPERMS);
                    break;

                case "DSpath":
                    ctl.Text = LM.Translate("CHOOSE", MessagesLabelsEN.CHOOSE);
                    break;

                


                case "label24":
                    ctl.Text = LM.Translate("VELOCITY_WU", MessagesLabelsEN.VELOCITY_WU);
                    break;

               case "label25":
                    ctl.Text = LM.Translate("Home_FSC", MessagesLabelsEN.Fsc);
                    break;

                case "testerinfo":
                    ctl.Text = LM.Translate("Tester_Information", MessagesLabelsEN.TESTER_INFO);
                    break;

                case "label18":
                    ctl.Text = LM.Translate("Home_Prog_mot", MessagesLabelsEN.PROG_MOT);
                    break;


                case "label19":
                    ctl.Text = LM.Translate("Home_morph", MessagesLabelsEN.MORPH_NORM);
                    break;

                case "label20":
                    ctl.Text = LM.Translate("Home_Immot", MessagesLabelsEN.IMOTILITY);
                    break;

                case "label21":
                    ctl.Text = LM.Translate("Home_Nonprog_mot", MessagesLabelsEN.NON_PROG_MOT);
                    break;

                case "label4":
                    ctl.Text = LM.Translate("CONC_WU", MessagesLabelsEN.CONC_WU);
                    break;


                case "label5":
                    ctl.Text = LM.Translate("Home_Tot_mot", MessagesLabelsEN.TOT_MOT);
                    break;

                case "Resultinfo":
                    ctl.Text = LM.Translate("RESULT_INFORMATION", MessagesLabelsEN.RESULT_INFO);
                    break;

                case "liquilbl":
                    ctl.Text = LM.Translate("LIQUIFACTION", MessagesLabelsEN.LIQUIFICTION);
                    break;

                case "testernamelbl":
                    ctl.Text = LM.Translate("TESTER_NAME", MessagesLabelsEN.TESTER_NAME).ToUpper();
                    break;


                case "tesdesg":
                    ctl.Text = LM.Translate("TESTER_DESIGNATION", MessagesLabelsEN.TESTER_DESG);
                    break;

                case "label7":
                    ctl.Text = LM.Translate("RECEIVED", MessagesLabelsEN.RECEIVED);
                    break;

                case "label8":
                    ctl.Text = LM.Translate("ABSTINENCE", MessagesLabelsEN.ABSTINENCE);
                    break;

                case "selectPortLabel":
                    ctl.Text = LM.Translate("SELECT_PORT", MessagesLabelsEN.SELECT_PORT);
                    break;
                case "PrinterNameLabel":
                    ctl.Text = LM.Translate("PRINTER_NAME", MessagesLabelsEN.PRINTER_NAME);
                    break;
                case "CancelButton":
                    ctl.Text = LM.Translate("CANCEL", MessagesLabelsEN.CANCEL);
                    break;
                case "ChoosePrinterButton":
                    ctl.Text = LM.Translate("CHOOSE", MessagesLabelsEN.CHOOSE);
                    break;
                case "FileNameLabel":
                    ctl.Text = LM.Translate("FILE_NAME", MessagesLabelsEN.FILE_NAME);
                    break;
                case "AutoPrintCheck":
                    ctl.Text = LM.Translate("AUTO_PRINT", MessagesLabelsEN.AUTO_PRINT);
                    break;
                case "pdfcheck": 
                    ctl.Text = LM.Translate("AUTO_PRINT", MessagesLabelsEN.AUTO_PRINT);
                    break;
                case "ExportCheck":
                    ctl.Text = LM.Translate("EXPORT_TO_EXCEL", MessagesLabelsEN.EXPORT_TO_EXCEL);
                    break;
                case "DateFormatLabel":
                    ctl.Text = LM.Translate("DATE_FORMAT_SET", MessagesLabelsEN.DATE_FORMAT_SET);
                    break;
                case "LanguageLabel":
                    ctl.Text = LM.Translate("LANGUAGE", MessagesLabelsEN.LANGUAGE) + ": ";
                    break;
                case "OpenFileButton":
                    ctl.Text = LM.Translate("CHOOSE", MessagesLabelsEN.CHOOSE);
                    break;
                case "SaveButton":
                    ctl.Text = LM.Translate("SAVE", MessagesLabelsEN.SAVE);
                    break;

                

                case "SettingsFrm":
                    ctl.Text = LM.Translate("SETTINGS", MessagesLabelsEN.SETTINGS);
                    break;
                case "ImportSet":
                    ctl.Text = LM.Translate("EXPORT_TO_EXCEL_SETTINGS", MessagesLabelsEN.EXPORT_TO_EXCEL_SETTINGS);
                    break;
                case "ReportSet":
                    ctl.Text = LM.Translate("REPORT_SETTINGS", MessagesLabelsEN.REPORT_SETTINGS);
                    break;
                case "MakePdfReportRadioButton":
                    ctl.Text = LM.Translate("MAKE_PDF_REPORT", MessagesLabelsEN.MAKE_PDF_REPORT);
                    break;
                case "PrintResultsRadioButton":
                    ctl.Text = LM.Translate("PRINT_TEST_RESULTS", MessagesLabelsEN.PRINT_TEST_RESULTS);
                    break;
                case "HeaderSpaceCheckBox":
                    ctl.Text = LM.Translate("ADD_HEADER_SPACE", MessagesLabelsEN.ADD_HEADER_SPACE);
                    break;
                case "HeaderSpaceLabel":
                    ctl.Text = LM.Translate("HEADER_SPACE_UNIT", MessagesLabelsEN.HEADER_SPACE_UNIT);
                    break;

              

                case "label3":
                   ctl.Text = LM.Translate("AGE", MessagesLabelsEN.AGE).TrimEnd(':'); 
                   
                    break;

                default:
                    break;
            }
        }
     
        
        
        private void SetControl(ToolStripItem ctl)
        {


            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);

            switch (ctl.Name.ToString())
            {

                case "csvselect":
                    ctl.Text = LM.Translate("CONFIRM", MessagesLabelsEN.CONFIRM);
                    break;

                case "Arc_cancel":
                    ctl.Text = LM.Translate("CANCEL", MessagesLabelsEN.CONFIRM);
                    break;


                case "generate":
                    ctl.Text = LM.Translate("Arc_Average", MessagesLabelsEN.CONFIRM);
                    break;
                case "ClearButton":
                    ctl.ToolTipText = LM.Translate("CLEAR", MessagesLabelsEN.CLEAR);
                    break;
                case "PrintButton":
                    if (Boolean.Parse(appSettings[7]) == true)
                    {
                        ctl.ToolTipText = LM.Translate("EXPORT_PDF", MessagesLabelsEN.PORT);
                    }
                    else
                    {
                        ctl.ToolTipText = LM.Translate("PRINT_RESULTS_STRIP", MessagesLabelsEN.PRINT_RESULTS_STRIP);
                    }

                    break;
                case "SettingsButton":
                    ctl.ToolTipText = LM.Translate("SETTINGS", MessagesLabelsEN.SETTINGS);
                    break;

                case "SaveButton":
                    ctl.Text = LM.Translate("SAVE", MessagesLabelsEN.SAVE);
                    //ctl.Text = string.Format(LM.Translate("Archive_tooltip", MessagesLabelsEN.SAVE), Environment.NewLine);
                    break;

                case "ArchiveButton":
                    ctl.ToolTipText = LM.Translate("Archive", MessagesLabelsEN.ARCHIVE);
                    break;
                
                    
                default:
                    break;
            }

        }

      

        public void CreateResxFromDB()
        {
            string filePathEN = Application.StartupPath + "\\Language" + @".\VResources.resx";
            ResXResourceWriter resxWriterEN = new ResXResourceWriter(filePathEN);
            string filePathDE = Application.StartupPath + "\\Language" + @".\VResources.de-De.resx";
            ResXResourceWriter resxWriterDE = new ResXResourceWriter(filePathDE);
            string filePathIT = Application.StartupPath + "\\Language" + @".\VResources.it-IT.resx";
            ResXResourceWriter resxWriterIT = new ResXResourceWriter(filePathIT);
            string filePathFR = Application.StartupPath + "\\Language" + @".\VResources.fr-FR.resx";
            ResXResourceWriter resxWriterFR = new ResXResourceWriter(filePathFR);
            string filePathZH = Application.StartupPath + "\\Language" + @".\VResources.zh-CN.resx";
            ResXResourceWriter resxWriterZH = new ResXResourceWriter(filePathZH);

            string LangDBPath = Application.StartupPath + "\\QcGold_Database" + @".\QcGoldLangDB.mdb";
            OpenAccessDatabase(LangDBPath);

            DataTable dtData = new DataTable();
            DataTable dtData2 = new DataTable();
            DataTable dtData3 = new DataTable();
            DataTable dtData4 = new DataTable();
            DataTable dtData5 = new DataTable();

            bool TestSql = OpenAccessTable("SELECT LabelCode, LabelText FROM T_translate", ref dtData);
            for (int i = 0; i < dtData.Rows.Count; i++)
            {
                resxWriterEN.AddResource(dtData.Rows[i]["LabelCode"].ToString(), dtData.Rows[i]["LabelText"].ToString());
            }
            resxWriterEN.Close();

            TestSql = OpenAccessTable("SELECT T_translate.LabelCode, T_Translated_Labels.TranslatedText FROM T_translate INNER JOIN T_Translated_Labels ON T_translate.ID = T_Translated_Labels.LabelID WHERE LanguageID = 1", ref dtData2);
            for (int i = 0; i < dtData2.Rows.Count; i++)
            {
                resxWriterDE.AddResource(dtData2.Rows[i]["LabelCode"].ToString(), dtData2.Rows[i]["TranslatedText"].ToString());
            }
            resxWriterDE.Close();

            TestSql = OpenAccessTable("SELECT T_translate.LabelCode, T_Translated_Labels.TranslatedText FROM T_translate INNER JOIN T_Translated_Labels ON T_translate.ID = T_Translated_Labels.LabelID WHERE LanguageID = 2", ref dtData3);
            for (int i = 0; i < dtData3.Rows.Count; i++)
            {
                resxWriterIT.AddResource(dtData3.Rows[i]["LabelCode"].ToString(), dtData3.Rows[i]["TranslatedText"].ToString());
            }
            resxWriterIT.Close();

            TestSql = OpenAccessTable("SELECT LabelCode, TranslatedText FROM French", ref dtData4);
            for (int i = 0; i < dtData4.Rows.Count; i++)
            {
                resxWriterFR.AddResource(dtData4.Rows[i]["LabelCode"].ToString(), dtData4.Rows[i]["TranslatedText"].ToString());
            }
            resxWriterFR.Close();
            
            TestSql = OpenAccessTable("SELECT LabelCode, TranslatedText FROM Chinese", ref dtData5);
            for (int i = 0; i < dtData5.Rows.Count; i++)
            {
                resxWriterZH.AddResource(dtData5.Rows[i]["LabelCode"].ToString(), dtData5.Rows[i]["TranslatedText"].ToString());
            }
            resxWriterZH.Close();

            CreateBinary(filePathEN);
            CreateBinary(filePathDE);
            CreateBinary(filePathIT);
            CreateBinary(filePathFR);
            CreateBinary(filePathZH);

        }

        private bool OpenAccessDatabase(string PathDB)
        {
            conAccess = new OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0; Data Source = " + PathDB /*+ " ; Persist Security Info=False;Mode= Share Deny None"*/);
            conAccess.Open();
            return true;
        }

        private bool OpenAccessTable(string SqlStr, ref DataTable table)
        {
            try
            {
                OleDbCommand comm = new OleDbCommand();
                comm.CommandType = CommandType.Text;
                comm.Connection = conAccess;
                comm.CommandText = SqlStr;
                OleDbDataAdapter dataAdapter = new OleDbDataAdapter(comm);
                new OleDbDataAdapter(comm).Fill(table);
            }
            catch
            {
                return false;
            }

            return true;
        }

        public void CreateResxEN()
        {
            string filePath = Application.StartupPath + "\\Language" + @".\VResources.resx";

            ResXResourceWriter resxWriter = new ResXResourceWriter(filePath);

            resxWriter.AddResource("PORT_ERROR", MessagesLabelsEN.PORT_ERROR);
            resxWriter.AddResource("PORT_ERROR_CAPTION", MessagesLabelsEN.PORT_ERROR_CAPTION);
            resxWriter.AddResource("EXPORT_ERROR", MessagesLabelsEN.EXPORT_ERROR);
            resxWriter.AddResource("END_OF_REPORT", MessagesLabelsEN.END_OF_REPORT);
            resxWriter.AddResource("REPORT_ERROR", MessagesLabelsEN.REPORT_ERROR);
            resxWriter.AddResource("REPORT_ERROR_CAPTION", MessagesLabelsEN.REPORT_ERROR_CAPTION);
            resxWriter.AddResource("REPORT_CREATION_ERROR", MessagesLabelsEN.REPORT_CREATION_ERROR);
            resxWriter.AddResource("P_Name", MessagesLabelsEN.P_Name);

            resxWriter.AddResource("DEVICE_SN", MessagesLabelsEN.DEVICE_SN);
            resxWriter.AddResource("SW_VERSION", MessagesLabelsEN.SW_VERSION);
            resxWriter.AddResource("TEST_DATE", MessagesLabelsEN.TEST_DATE);
            resxWriter.AddResource("PATIENT_ID", MessagesLabelsEN.PATIENT_ID);
            resxWriter.AddResource("VOID", MessagesLabelsEN.VOID);
            resxWriter.AddResource("BIRTH_DATE", MessagesLabelsEN.BIRTH_DATE);
            resxWriter.AddResource("ABSTINENCE", MessagesLabelsEN.ABSTINENCE);
            resxWriter.AddResource("ACCESSION", MessagesLabelsEN.ACCESSION);
            resxWriter.AddResource("COLLECTED_DATE", MessagesLabelsEN.COLLECTED_DATE);
            resxWriter.AddResource("RECIEVED", MessagesLabelsEN.RECEIVED);
            resxWriter.AddResource("TYPE", MessagesLabelsEN.TYPE);
            resxWriter.AddResource("VOLUME", MessagesLabelsEN.VOLUME);
            resxWriter.AddResource("WBC_CONC", MessagesLabelsEN.WBC_CONC);
            resxWriter.AddResource("PH", MessagesLabelsEN.PH);
            resxWriter.AddResource("CONC", MessagesLabelsEN.CONC);
            resxWriter.AddResource("PR_NP", MessagesLabelsEN.PR_NP);
            resxWriter.AddResource("PR_NP_REM", MessagesLabelsEN.PR_NP_REM);
            resxWriter.AddResource("PROG", MessagesLabelsEN.PROG);
            resxWriter.AddResource("PROG_REM", MessagesLabelsEN.PROG_REM);
            resxWriter.AddResource("NONPROG", MessagesLabelsEN.NONPROG);
            resxWriter.AddResource("IMMOT", MessagesLabelsEN.IMMOT);
            resxWriter.AddResource("WHO_5", MessagesLabelsEN.WHO_5);
            resxWriter.AddResource("WHO_5_REM", MessagesLabelsEN.WHO_5_REM);
            resxWriter.AddResource("MSC", MessagesLabelsEN.MSC);
            resxWriter.AddResource("PMSC", MessagesLabelsEN.PMSC);
            resxWriter.AddResource("FSC", MessagesLabelsEN.FSC);
            resxWriter.AddResource("VELOCITY", MessagesLabelsEN.VELOCITY);
            resxWriter.AddResource("SMI", MessagesLabelsEN.SMI);
            resxWriter.AddResource("NUM_SPERM", MessagesLabelsEN.NUM_SPERM);
            resxWriter.AddResource("MOTILE_SPERM", MessagesLabelsEN.MOTILE_SPERM);
            resxWriter.AddResource("PROG_SPERM", MessagesLabelsEN.PROG_SPERM);
            resxWriter.AddResource("FUNC_SPERM", MessagesLabelsEN.FUNC_SPERM);
            resxWriter.AddResource("SPERM", MessagesLabelsEN.SPERM);
            resxWriter.AddResource("SPERM_REM", MessagesLabelsEN.SPERM_REM);
            resxWriter.AddResource("SYSTEM_DATA", MessagesLabelsEN.SYSTEM_DATA);

            resxWriter.AddResource("BIRTH", MessagesLabelsEN.BIRTH);
            resxWriter.AddResource("COLLECTED", MessagesLabelsEN.COLLECTED);

            resxWriter.AddResource("MOTILITY_RESULTS", MessagesLabelsEN.MOTILITY_RESULTS);
            
            resxWriter.AddResource("DAYS", MessagesLabelsEN.DAYS);
            resxWriter.AddResource("ML", MessagesLabelsEN.ML);
            resxWriter.AddResource("M_PER_ML", MessagesLabelsEN.M_PER_ML);
            resxWriter.AddResource("PERCENT", MessagesLabelsEN.PERCENT);
            resxWriter.AddResource("M", MessagesLabelsEN.M);
            resxWriter.AddResource("MIC_PER_SEC", MessagesLabelsEN.MIC_PER_SEC);
            resxWriter.AddResource("FRESH", MessagesLabelsEN.FRESH);
            resxWriter.AddResource("FROZEN", MessagesLabelsEN.FROZEN);
            resxWriter.AddResource("WASHED", MessagesLabelsEN.WASHED);

            resxWriter.AddResource("TEST_DATE_WU", MessagesLabelsEN.TEST_DATE_WU);
            resxWriter.AddResource("COLLECTED_DATE_WU", MessagesLabelsEN.COLLECTED_DATE_WU);
            resxWriter.AddResource("RECIEVED_WU", MessagesLabelsEN.RECEIVED_WU);
            resxWriter.AddResource("ABSTINENCE_WU", MessagesLabelsEN.ABSTINENCE_WU);
            resxWriter.AddResource("VOLUME_WU", MessagesLabelsEN.VOLUME_WU);
            resxWriter.AddResource("WBC_CONC_WU", MessagesLabelsEN.WBC_CONC_WU);
            resxWriter.AddResource("CONC_WU", MessagesLabelsEN.CONC_WU);
            resxWriter.AddResource("PR_NP_WU", MessagesLabelsEN.PR_NP_WU);
            resxWriter.AddResource("PROG_WU", MessagesLabelsEN.PROG_WU);
            resxWriter.AddResource("NONPROG_WU", MessagesLabelsEN.NONPROG_WU);
            resxWriter.AddResource("IMMOT_WU", MessagesLabelsEN.IMMOT_WU);
            resxWriter.AddResource("WHO_5_WU", MessagesLabelsEN.WHO_5_WU);
            resxWriter.AddResource("MSC_WU", MessagesLabelsEN.MSC_WU);
            resxWriter.AddResource("PMSC_WU", MessagesLabelsEN.PMSC_WU);
            resxWriter.AddResource("FSC_WU", MessagesLabelsEN.FSC_WU);
            resxWriter.AddResource("VELOCITY_WU", MessagesLabelsEN.VELOCITY_WU);
            resxWriter.AddResource("NUM_SPERM_WU", MessagesLabelsEN.NUM_SPERM_WU);
            resxWriter.AddResource("MOTILE_SPERM_WU", MessagesLabelsEN.MOTILE_SPERM_WU);
            resxWriter.AddResource("PROG_SPERM_WU", MessagesLabelsEN.PROG_SPERM_WU);
            resxWriter.AddResource("FUNC_SPERM_WU", MessagesLabelsEN.FUNC_SPERM_WU);
            resxWriter.AddResource("SPERM_WU", MessagesLabelsEN.SPERM_WU);
            resxWriter.AddResource("DEVICE_SN_WU", MessagesLabelsEN.DEVICE_SN_WU);

            resxWriter.AddResource("TITLE_TEST_RESULTS", MessagesLabelsEN.TITLE_TEST_RESULTS);
            resxWriter.AddResource("PATIENT_DATA", MessagesLabelsEN.PATIENT_DATA);
            resxWriter.AddResource("SAMPLE_DATA", MessagesLabelsEN.SAMPLE_DATA);
            resxWriter.AddResource("TEST_RESULTS", MessagesLabelsEN.TEST_RESULTS);
            resxWriter.AddResource("TOTALS_PER_VOL", MessagesLabelsEN.TOTALS_PER_VOL);
            resxWriter.AddResource("OTHER", MessagesLabelsEN.OTHER);
            resxWriter.AddResource("EXCEL_FILE_ERROR_CAPTION", MessagesLabelsEN.EXCEL_FILE_ERROR_CAPTION);
            resxWriter.AddResource("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR);
            resxWriter.AddResource("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED);
            resxWriter.AddResource("EXPORT_FILE_ERROR", MessagesLabelsEN.EXPORT_FILE_ERROR);
            resxWriter.AddResource("PRINTER_ERROR", MessagesLabelsEN.PRINTER_ERROR);
            resxWriter.AddResource("SETTINGS_WINDOW_ALREADY_OPENED", MessagesLabelsEN.SETTINGS_WINDOW_ALREADY_OPENED);
            resxWriter.AddResource("SETTINGS", MessagesLabelsEN.SETTINGS);
            resxWriter.AddResource("PRINT", MessagesLabelsEN.PRINT);
            resxWriter.AddResource("CLEAR", MessagesLabelsEN.CLEAR);
            resxWriter.AddResource("YES", MessagesLabelsEN.YES);
            resxWriter.AddResource("NO", MessagesLabelsEN.NO);
            resxWriter.AddResource("EXPORT_FILE_NAME", MessagesLabelsEN.EXPORT_FILE_NAME);
            resxWriter.AddResource("VERSION", MessagesLabelsEN.VERSION);
            resxWriter.AddResource("SETTINGS_FILE_CORRUPTED_DEFAULT", MessagesLabelsEN.SETTINGS_FILE_CORRUPTED_DEFAULT);
            resxWriter.AddResource("SETTINGS_FILE_CORRUPTED_RELOAD", MessagesLabelsEN.SETTINGS_FILE_CORRUPTED_RELOAD);
            resxWriter.AddResource("NONE", MessagesLabelsEN.NONE);
            resxWriter.AddResource("SELECT_FILE_TO_EXPORT", MessagesLabelsEN.SELECT_FILE_TO_EXPORT);
            resxWriter.AddResource("EXCEL_NOT_EXIST", MessagesLabelsEN.EXCEL_NOT_EXIST);
            resxWriter.AddResource("SAVE_CHANGES", MessagesLabelsEN.SAVE_CHANGES);
            resxWriter.AddResource("CONFIRM_CHANGES", MessagesLabelsEN.CONFIRM_CHANGES);
            resxWriter.AddResource("SPECIFY_FILE_NAME", MessagesLabelsEN.SPECIFY_FILE_NAME);
            resxWriter.AddResource("EXPORT_FILENAME_MISSING", MessagesLabelsEN.EXPORT_FILENAME_MISSING);
            resxWriter.AddResource("INVALID_FILEPATH", MessagesLabelsEN.INVALID_FILEPATH);
            resxWriter.AddResource("INVALID_FILENAME", MessagesLabelsEN.INVALID_FILENAME);
            resxWriter.AddResource("SETTINGS_FILE_ERROR", MessagesLabelsEN.SETTINGS_FILE_ERROR);
            resxWriter.AddResource("FILE_OPEN_ERROR", MessagesLabelsEN.FILE_OPEN_ERROR);
            resxWriter.AddResource("EXPORT_ERROR_CAPTION", MessagesLabelsEN.EXPORT_ERROR_CAPTION);


            resxWriter.AddResource("PORT", MessagesLabelsEN.PORT);
            resxWriter.AddResource("PRINTER", MessagesLabelsEN.PRINTER);
            resxWriter.AddResource("EXPORT_FILE", MessagesLabelsEN.EXPORT_FILE);
            resxWriter.AddResource("AUTO_PRINT", MessagesLabelsEN.AUTO_PRINT);
            resxWriter.AddResource("EXPORT", MessagesLabelsEN.EXPORT);
            resxWriter.AddResource("DATE_FORMAT", MessagesLabelsEN.DATE_FORMAT_SET);

            resxWriter.AddResource("SELECT_PORT", MessagesLabelsEN.SELECT_PORT);
            resxWriter.AddResource("PRINTER_NAME", MessagesLabelsEN.PRINTER_NAME);
            resxWriter.AddResource("LANGUAGE", MessagesLabelsEN.LANGUAGE);
            resxWriter.AddResource("CHOOSE", MessagesLabelsEN.CHOOSE);
            resxWriter.AddResource("EXPORT_TO_EXCEL", MessagesLabelsEN.EXPORT_TO_EXCEL);
            resxWriter.AddResource("EXPORT_TO_EXCEL_SETTINGS", MessagesLabelsEN.EXPORT_TO_EXCEL_SETTINGS);
            resxWriter.AddResource("FILE_NAME", MessagesLabelsEN.FILE_NAME);
            resxWriter.AddResource("DATE_FORMAT_SET", MessagesLabelsEN.DATE_FORMAT_SET);
            resxWriter.AddResource("OPEN", MessagesLabelsEN.OPEN);
            resxWriter.AddResource("SAVE", MessagesLabelsEN.SAVE);
            resxWriter.AddResource("CANCEL", MessagesLabelsEN.CANCEL);
            resxWriter.AddResource("PRINTED_FROM", MessagesLabelsEN.PRINTED_FROM);
            resxWriter.AddResource("REPORT_SETTINGS", MessagesLabelsEN.REPORT_SETTINGS);
            resxWriter.AddResource("MAKE_PDF_REPORT", MessagesLabelsEN.MAKE_PDF_REPORT);
            resxWriter.AddResource("PRINT_TEST_RESULTS", MessagesLabelsEN.PRINT_TEST_RESULTS);
            resxWriter.AddResource("ADD_HEADER_SPACE", MessagesLabelsEN.ADD_HEADER_SPACE);
            resxWriter.AddResource("HEADER_SPACE_UNIT", MessagesLabelsEN.HEADER_SPACE_UNIT);

            resxWriter.AddResource("REPORT_FILE_NAME", MessagesLabelsEN.REPORT_FILE_NAME);
            resxWriter.AddResource("TEST_REPORT", MessagesLabelsEN.TEST_REPORT);
            resxWriter.AddResource("REPORT_CONTINUED", MessagesLabelsEN.REPORT_CONTINUED);
            resxWriter.AddResource("DEVICE_INFO", MessagesLabelsEN.DEVICE_INFO);
            resxWriter.AddResource("PATIENT_INFO", MessagesLabelsEN.PATIENT_INFO);
            resxWriter.AddResource("SAMPLE_INFO", MessagesLabelsEN.SAMPLE_INFO);
            resxWriter.AddResource("PARAMETER", MessagesLabelsEN.PARAMETER);
            resxWriter.AddResource("RESULT", MessagesLabelsEN.RESULT);

            resxWriter.Close();

            CreateBinary(filePath);

        }

        public void CreateResxDE()
        {
            string filePath = Application.StartupPath + "\\Language" + @".\VResources.de-De.resx";

            ResXResourceWriter resxWriter = new ResXResourceWriter(filePath);

            resxWriter.AddResource("PORT_ERROR", MessagesLabelsDE.PORT_ERROR);
            resxWriter.AddResource("PORT_ERROR_CAPTION", MessagesLabelsDE.PORT_ERROR_CAPTION);
            resxWriter.AddResource("EXPORT_ERROR", MessagesLabelsDE.EXPORT_ERROR);
            resxWriter.AddResource("END_OF_REPORT", MessagesLabelsDE.END_OF_REPORT);
            resxWriter.AddResource("REPORT_ERROR", MessagesLabelsDE.REPORT_ERROR);
            resxWriter.AddResource("REPORT_ERROR_CAPTION", MessagesLabelsDE.REPORT_ERROR_CAPTION);
            resxWriter.AddResource("REPORT_CREATION_ERROR", MessagesLabelsDE.REPORT_CREATION_ERROR);

            resxWriter.AddResource("DEVICE_SN", MessagesLabelsDE.DEVICE_SN);
            resxWriter.AddResource("SW_VERSION", MessagesLabelsDE.SW_VERSION);
            resxWriter.AddResource("TEST_DATE", MessagesLabelsDE.TEST_DATE);
            resxWriter.AddResource("PATIENT_ID", MessagesLabelsDE.PATIENT_ID);
            resxWriter.AddResource("VOID", MessagesLabelsDE.VOID);
            resxWriter.AddResource("BIRTH_DATE", MessagesLabelsDE.BIRTH_DATE);
            resxWriter.AddResource("ABSTINENCE", MessagesLabelsDE.ABSTINENCE);
            resxWriter.AddResource("ACCESSION", MessagesLabelsDE.ACCESSION);
            resxWriter.AddResource("COLLECTED_DATE", MessagesLabelsDE.COLLECTED_DATE);
            resxWriter.AddResource("RECIEVED", MessagesLabelsDE.RECEIVED);
            resxWriter.AddResource("TYPE", MessagesLabelsDE.TYPE);
            resxWriter.AddResource("VOLUME", MessagesLabelsDE.VOLUME);
            resxWriter.AddResource("WBC_CONC", MessagesLabelsDE.WBC_CONC);
            resxWriter.AddResource("PH", MessagesLabelsDE.PH);
            resxWriter.AddResource("CONC", MessagesLabelsDE.CONC);
            resxWriter.AddResource("PR_NP", MessagesLabelsDE.PR_NP);
            resxWriter.AddResource("PROG", MessagesLabelsDE.PROG);
            resxWriter.AddResource("NONPROG", MessagesLabelsDE.NONPROG);
            resxWriter.AddResource("IMMOT", MessagesLabelsDE.IMMOT);
            resxWriter.AddResource("WHO_5", MessagesLabelsDE.WHO_5);
            resxWriter.AddResource("MSC", MessagesLabelsDE.MSC);
            resxWriter.AddResource("PMSC", MessagesLabelsDE.PMSC);
            resxWriter.AddResource("FSC", MessagesLabelsDE.FSC);
            resxWriter.AddResource("VELOCITY", MessagesLabelsDE.VELOCITY);
            resxWriter.AddResource("SMI", MessagesLabelsDE.SMI);
            resxWriter.AddResource("NUM_SPERM", MessagesLabelsDE.NUM_SPERM);
            resxWriter.AddResource("MOTILE_SPERM", MessagesLabelsDE.MOTILE_SPERM);
            resxWriter.AddResource("PROG_SPERM", MessagesLabelsDE.PROG_SPERM);
            resxWriter.AddResource("FUNC_SPERM", MessagesLabelsDE.FUNC_SPERM);
            resxWriter.AddResource("SPERM", MessagesLabelsDE.SPERM);
            resxWriter.AddResource("SYSTEM_DATA", MessagesLabelsDE.SYSTEM_DATA);

            resxWriter.AddResource("BIRTH", MessagesLabelsDE.BIRTH);
            resxWriter.AddResource("COLLECTED", MessagesLabelsDE.COLLECTED);

            resxWriter.AddResource("MOTILITY_RESULTS",MessagesLabelsDE.MOTILITY_RESULTS);

            resxWriter.AddResource("DAYS", MessagesLabelsDE.DAYS);
            resxWriter.AddResource("ML", MessagesLabelsDE.ML);
            resxWriter.AddResource("M_PER_ML", MessagesLabelsDE.M_PER_ML);
            resxWriter.AddResource("PERCENT", MessagesLabelsDE.PERCENT);
            resxWriter.AddResource("M", MessagesLabelsDE.M);
            resxWriter.AddResource("MIC_PER_SEC", MessagesLabelsDE.MIC_PER_SEC);
            resxWriter.AddResource("FRESH", MessagesLabelsDE.FRESH);
            resxWriter.AddResource("FROZEN", MessagesLabelsDE.FROZEN);
            resxWriter.AddResource("WASHED", MessagesLabelsDE.WASHED);

            resxWriter.AddResource("TEST_DATE_WU", MessagesLabelsDE.TEST_DATE_WU);
            resxWriter.AddResource("COLLECTED_DATE_WU", MessagesLabelsDE.COLLECTED_DATE_WU);
            resxWriter.AddResource("RECIEVED_WU", MessagesLabelsDE.RECEIVED_WU);
            resxWriter.AddResource("ABSTINENCE_WU", MessagesLabelsDE.ABSTINENCE_WU);
            resxWriter.AddResource("VOLUME_WU", MessagesLabelsDE.VOLUME_WU);
            resxWriter.AddResource("WBC_CONC_WU", MessagesLabelsDE.WBC_CONC_WU);
            resxWriter.AddResource("CONC_WU", MessagesLabelsDE.CONC_WU);
            resxWriter.AddResource("PR_NP_WU", MessagesLabelsDE.PR_NP_WU);
            resxWriter.AddResource("PROG_WU", MessagesLabelsDE.PROG_WU);
            resxWriter.AddResource("NONPROG_WU", MessagesLabelsDE.NONPROG_WU);
            resxWriter.AddResource("IMMOT_WU", MessagesLabelsDE.IMMOT_WU);
            resxWriter.AddResource("WHO_5_WU", MessagesLabelsDE.WHO_5_WU);
            resxWriter.AddResource("MSC_WU", MessagesLabelsDE.MSC_WU);
            resxWriter.AddResource("PMSC_WU", MessagesLabelsDE.PMSC_WU);
            resxWriter.AddResource("FSC_WU", MessagesLabelsDE.FSC_WU);
            resxWriter.AddResource("VELOCITY_WU", MessagesLabelsDE.VELOCITY_WU);
            resxWriter.AddResource("NUM_SPERM_WU", MessagesLabelsDE.NUM_SPERM_WU);
            resxWriter.AddResource("MOTILE_SPERM_WU", MessagesLabelsDE.MOTILE_SPERM_WU);
            resxWriter.AddResource("PROG_SPERM_WU", MessagesLabelsDE.PROG_SPERM_WU);
            resxWriter.AddResource("FUNC_SPERM_WU", MessagesLabelsDE.FUNC_SPERM_WU);
            resxWriter.AddResource("SPERM_WU", MessagesLabelsDE.SPERM_WU);
            resxWriter.AddResource("DEVICE_SN_WU", MessagesLabelsDE.DEVICE_SN_WU);

            resxWriter.AddResource("TITLE_TEST_RESULTS", MessagesLabelsDE.TITLE_TEST_RESULTS);
            resxWriter.AddResource("PATIENT_DATA", MessagesLabelsDE.PATIENT_DATA);
            resxWriter.AddResource("SAMPLE_DATA", MessagesLabelsDE.SAMPLE_DATA);
            resxWriter.AddResource("TEST_RESULTS", MessagesLabelsDE.TEST_RESULTS);
            resxWriter.AddResource("TOTALS_PER_VOL", MessagesLabelsDE.TOTALS_PER_VOL);
            resxWriter.AddResource("OTHER", MessagesLabelsDE.OTHER);
            resxWriter.AddResource("EXCEL_FILE_ERROR_CAPTION", MessagesLabelsDE.EXCEL_FILE_ERROR_CAPTION);
            resxWriter.AddResource("ACCESS_ERROR", MessagesLabelsDE.ACCESS_ERROR);
            resxWriter.AddResource("IMPORT_FAILED", MessagesLabelsDE.IMPORT_FAILED);
            resxWriter.AddResource("EXPORT_FILE_ERROR",MessagesLabelsDE.EXPORT_FILE_ERROR);
            resxWriter.AddResource("PRINTER_ERROR", MessagesLabelsDE.PRINTER_ERROR);
            resxWriter.AddResource("SETTINGS_WINDOW_ALREADY_OPENED", MessagesLabelsDE.SETTINGS_WINDOW_ALREADY_OPENED);
            resxWriter.AddResource("SETTINGS", MessagesLabelsDE.SETTINGS);
            resxWriter.AddResource("PRINT", MessagesLabelsDE.PRINT);
            resxWriter.AddResource("CLEAR", MessagesLabelsDE.CLEAR);
            resxWriter.AddResource("YES", MessagesLabelsDE.YES);
            resxWriter.AddResource("NO", MessagesLabelsDE.NO);
            resxWriter.AddResource("EXPORT_FILE_NAME", MessagesLabelsDE.EXPORT_FILE_NAME);
            resxWriter.AddResource("VERSION", MessagesLabelsDE.VERSION);
            resxWriter.AddResource("SETTINGS_FILE_CORRUPTED_DEFAULT", MessagesLabelsDE.SETTINGS_FILE_CORRUPTED_DEFAULT);
            resxWriter.AddResource("SETTINGS_FILE_CORRUPTED_RELOAD", MessagesLabelsDE.SETTINGS_FILE_CORRUPTED_RELOAD);
            resxWriter.AddResource("NONE", MessagesLabelsDE.NONE);
            resxWriter.AddResource("SELECT_FILE_TO_EXPORT", MessagesLabelsDE.SELECT_FILE_TO_EXPORT);
            resxWriter.AddResource("EXCEL_NOT_EXIST", MessagesLabelsDE.EXCEL_NOT_EXIST);
            resxWriter.AddResource("SAVE_CHANGES", MessagesLabelsDE.SAVE_CHANGES);
            resxWriter.AddResource("CONFIRM_CHANGES", MessagesLabelsDE.CONFIRM_CHANGES);
            resxWriter.AddResource("SPECIFY_FILE_NAME", MessagesLabelsDE.SPECIFY_FILE_NAME);
            resxWriter.AddResource("EXPORT_FILENAME_MISSING", MessagesLabelsDE.EXPORT_FILENAME_MISSING);
            resxWriter.AddResource("INVALID_FILEPATH", MessagesLabelsDE.INVALID_FILEPATH);
            resxWriter.AddResource("INVALID_FILENAME",MessagesLabelsDE.INVALID_FILENAME);
            resxWriter.AddResource("SETTINGS_FILE_ERROR", MessagesLabelsDE.SETTINGS_FILE_ERROR);
            resxWriter.AddResource("FILE_OPEN_ERROR", MessagesLabelsDE.FILE_OPEN_ERROR);
            resxWriter.AddResource("EXPORT_ERROR_CAPTION", MessagesLabelsDE.EXPORT_ERROR_CAPTION);

            resxWriter.AddResource("PORT", MessagesLabelsDE.PORT);
            resxWriter.AddResource("PRINTER", MessagesLabelsDE.PRINTER);
            resxWriter.AddResource("EXPORT_FILE", MessagesLabelsDE.EXPORT_FILE);
            resxWriter.AddResource("AUTO_PRINT", MessagesLabelsDE.AUTO_PRINT);
            resxWriter.AddResource("EXPORT", MessagesLabelsDE.EXPORT);
            resxWriter.AddResource("DATE_FORMAT", MessagesLabelsDE.DATE_FORMAT);

            resxWriter.AddResource("SELECT_PORT", MessagesLabelsDE.SELECT_PORT);
            resxWriter.AddResource("PRINTER_NAME", MessagesLabelsDE.PRINTER_NAME);
            resxWriter.AddResource("LANGUAGE", MessagesLabelsDE.LANGUAGE);
            resxWriter.AddResource("CHOOSE", MessagesLabelsDE.CHOOSE);
            resxWriter.AddResource("EXPORT_TO_EXCEL", MessagesLabelsDE.EXPORT_TO_EXCEL);
            resxWriter.AddResource("EXPORT_TO_EXCEL_SETTINGS", MessagesLabelsDE.EXPORT_TO_EXCEL_SETTINGS);
            resxWriter.AddResource("FILE_NAME", MessagesLabelsDE.FILE_NAME);
            resxWriter.AddResource("DATE_FORMAT_SET", MessagesLabelsDE.DATE_FORMAT_SET);
            resxWriter.AddResource("OPEN", MessagesLabelsDE.OPEN);
            resxWriter.AddResource("SAVE", MessagesLabelsDE.SAVE);
            resxWriter.AddResource("CANCEL", MessagesLabelsDE.CANCEL);
            resxWriter.AddResource("PRINTED_FROM", MessagesLabelsDE.PRINTED_FROM);
            resxWriter.AddResource("REPORT_SETTINGS", MessagesLabelsDE.REPORT_SETTINGS);
            resxWriter.AddResource("MAKE_PDF_REPORT", MessagesLabelsDE.MAKE_PDF_REPORT);
            resxWriter.AddResource("PRINT_TEST_RESULTS", MessagesLabelsDE.PRINT_TEST_RESULTS);
            resxWriter.AddResource("ADD_HEADER_SPACE", MessagesLabelsDE.ADD_HEADER_SPACE);
            resxWriter.AddResource("HEADER_SPACE_UNIT", MessagesLabelsDE.HEADER_SPACE_UNIT);

            resxWriter.AddResource("REPORT_FILE_NAME", MessagesLabelsDE.REPORT_FILE_NAME);
            resxWriter.AddResource("TEST_REPORT", MessagesLabelsDE.TEST_REPORT);
            resxWriter.AddResource("REPORT_CONTINUED", MessagesLabelsDE.REPORT_CONTINUED);
            resxWriter.AddResource("DEVICE_INFO", MessagesLabelsDE.DEVICE_INFO);
            resxWriter.AddResource("PATIENT_INFO", MessagesLabelsDE.PATIENT_INFO);
            resxWriter.AddResource("SAMPLE_INFO", MessagesLabelsDE.SAMPLE_INFO);
            resxWriter.AddResource("PARAMETER", MessagesLabelsDE.PARAMETER);
            resxWriter.AddResource("RESULT", MessagesLabelsDE.RESULT);

            resxWriter.Close();

            CreateBinary(filePath);

        }

        public void CreateBinary(string filePath)
        {
            string CurrentPath = Application.StartupPath;

            int ExitCode;
            ProcessStartInfo ProcessInfo;
            Process Process = new Process();

            ProcessInfo = new ProcessStartInfo("resgen.exe", filePath.Substring(CurrentPath.Length + 1));

            ProcessInfo.UseShellExecute = false;
            ProcessInfo.RedirectStandardOutput = true;

            Process = Process.Start(ProcessInfo);
            string output = Process.StandardOutput.ReadToEnd();
            Process.WaitForExit();
            ExitCode = Process.ExitCode;
            Process.Close();

        }
    }
}
