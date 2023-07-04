using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Globalization;
using System.Diagnostics;
using System.IO;
using System.Data;
using System.Windows;
using System.Windows.Forms;
using QcGoldArchive;
using Application = System.Windows.Forms.Application;

namespace QcGoldArchive

{



    public class ReportTableBuilder
    {

       






    LanguageManagement LM = LanguageManagement.CreateInstance();

        bool frozen = false;

        public void addheader(DataTable DT, List<string> val)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("value", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = LM.Translate("DEVICE_SN", MessagesLabelsEN.DEVICE_SN);
            row0["value"] = val[0];
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = LM.Translate("SW_VERSION", MessagesLabelsEN.SW_VERSION);
            row1["value"] = val[1];
            DT.Rows.Add(row1);

        }


        public void CreateDeviceInformation(DataTable DT, List<string> val)
        {
            try
            {
                DT.Columns.Add("parameter", typeof(string));
                DT.Columns.Add("value", typeof(string));

                DataRow row0 = DT.NewRow();
                row0["parameter"] = LM.Translate("DEVICE_SN", MessagesLabelsEN.DEVICE_SN);
                row0["value"] = val[0];
                DT.Rows.Add(row0);

                DataRow row1 = DT.NewRow();
                row1["parameter"] = LM.Translate("SW_VERSION", MessagesLabelsEN.SW_VERSION);
                row1["value"] = val[1];
                DT.Rows.Add(row1);

            }
            catch (Exception ex)
            {

            }
        }
        public void CreatePatientInformation(DataTable DT, List<string> val)
        {
            try
            {
                DT.Columns.Add("parameter", typeof(string));
                DT.Columns.Add("value", typeof(string));

                DataRow row0 = DT.NewRow();
                row0["parameter"] = LM.Translate("PATIENT_ID", MessagesLabelsEN.PATIENT_ID);
                row0["value"] = val[0];
                DT.Rows.Add(row0);

                DataRow row1 = DT.NewRow();
                row1["parameter"] = LM.Translate("BIRTH_DATE", MessagesLabelsEN.BIRTH_DATE);
                row1["value"] = val[1];
                DT.Rows.Add(row1);
            }
            catch (Exception ex)
            {

            }


        }





        public void CreateSampleInformationTable(DataTable DT, List<string> val)
        {
            try
            {
                DT.Columns.Add("parameter", typeof(string));
                DT.Columns.Add("value", typeof(string));

                DataRow row0 = DT.NewRow();
                row0["parameter"] = LM.Translate("TEST_DATE_WU", MessagesLabelsEN.TEST_DATE_WU);
                row0["value"] = val[0];
                DT.Rows.Add(row0);

                DataRow row1 = DT.NewRow();
                row1["parameter"] = LM.Translate("ABSTINENCE_WU", MessagesLabelsEN.ABSTINENCE_WU);
                row1["value"] = val[1];
                DT.Rows.Add(row1);

                DataRow row2 = DT.NewRow();
                row2["parameter"] = LM.Translate("ACCESSION", MessagesLabelsEN.ACCESSION);
                row2["value"] = val[2];
                DT.Rows.Add(row2);

                DataRow row3 = DT.NewRow();
                row3["parameter"] = LM.Translate("COLLECTED_DATE_WU", MessagesLabelsEN.COLLECTED_DATE_WU);
                row3["value"] = val[3];
                DT.Rows.Add(row3);

                DataRow row4 = DT.NewRow();
                row4["parameter"] = LM.Translate("RECEIVED_WU", MessagesLabelsEN.RECEIVED_WU);
                row4["value"] = val[4];
                DT.Rows.Add(row4);

                DataRow row5 = DT.NewRow();
                row5["parameter"] = LM.Translate("TYPE", MessagesLabelsEN.TYPE);
                row5["value"] = val[5];
                DT.Rows.Add(row5);

                DataRow row6 = DT.NewRow();
                row6["parameter"] = LM.Translate("VOLUME_WU", MessagesLabelsEN.VOLUME_WU);
                row6["value"] = val[6];
                DT.Rows.Add(row6);

                DataRow row7 = DT.NewRow();
                row7["parameter"] = LM.Translate("WBC_CONC_WU", MessagesLabelsEN.WBC_CONC_WU);
                row7["value"] = val[7];
                DT.Rows.Add(row7);

                DataRow row8 = DT.NewRow();
                row8["parameter"] = LM.Translate("PH", MessagesLabelsEN.PH);
                row8["value"] = val[8];
                DT.Rows.Add(row8);

                //if (val[5].Equals(LM.Translate("FROZEN", MessagesLabelsEN.FROZEN)))
                //{
                //    frozen = true;
                //}
            }
            catch (Exception ex)
            {

            }
        }

        public void CreateSemAnalTestResultsTable(DataTable DT, List<string> val, bool shortReport)
        {

            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("result", typeof(string));
            if (shortReport)
            {
                DataRow row1 = DT.NewRow();
                row1["parameter"] = LM.Translate("MSC_WU", MessagesLabelsEN.MSC_WU);
                row1["result"] = val[6];
                DT.Rows.Add(row1);

                DataRow row2 = DT.NewRow();
                row2["parameter"] = LM.Translate("PMSC_WU", MessagesLabelsEN.PMSC_WU);
                row2["result"] = val[7];
                DT.Rows.Add(row2);

                DataRow row3 = DT.NewRow();
                row3["parameter"] = LM.Translate("VELOCITY_WU", MessagesLabelsEN.VELOCITY_WU);
                row3["result"] = val[9];
                DT.Rows.Add(row3);

                DataRow row4 = DT.NewRow();
                row4["parameter"] = LM.Translate("SMI", MessagesLabelsEN.SMI);
                row4["result"] = val[10];
                DT.Rows.Add(row4);

                DataRow row5 = DT.NewRow();
                row5["parameter"] = LM.Translate("MOTILE_SPERM_WU", MessagesLabelsEN.MOTILE_SPERM_WU);
                row5["result"] = val[12];
                DT.Rows.Add(row5);

                DataRow row6 = DT.NewRow();
                row6["parameter"] = LM.Translate("PROG_SPERM_WU", MessagesLabelsEN.PROG_SPERM_WU);
                row6["result"] = val[13];
                DT.Rows.Add(row6);

            }
            else
            {
                try
                {

                    DataRow row0 = DT.NewRow();
                    row0["parameter"] = LM.Translate("CONC_WU", MessagesLabelsEN.CONC_WU);
                    row0["result"] = val[0];
                    DT.Rows.Add(row0);

                    DataRow row1 = DT.NewRow();
                    row1["parameter"] = LM.Translate("PR_NP_WU", MessagesLabelsEN.PR_NP_WU);
                    row1["result"] = val[1];
                    DT.Rows.Add(row1);

                    DataRow row2 = DT.NewRow();
                    row2["parameter"] = LM.Translate("PROG_WU", MessagesLabelsEN.PROG_WU);
                    row2["result"] = val[2];
                    DT.Rows.Add(row2);

                    DataRow row3 = DT.NewRow();
                    row3["parameter"] = LM.Translate("NONPROG_WU", MessagesLabelsEN.NONPROG_WU);
                    row3["result"] = val[3];
                    DT.Rows.Add(row3);

                    DataRow row4 = DT.NewRow();
                    row4["parameter"] = LM.Translate("IMMOT_WU", MessagesLabelsEN.IMMOT_WU);
                    row4["result"] = val[4];
                    DT.Rows.Add(row4);

                    DataRow row5 = DT.NewRow();
                    row5["parameter"] = LM.Translate("WHO_5_WU", MessagesLabelsEN.WHO_5_WU);
                    if (val[5] == "N.A.")
                    { val[5] = "N/A"; }
                    row5["result"] = val[5];
                    DT.Rows.Add(row5);

                    DataRow row6 = DT.NewRow();
                    row6["parameter"] = LM.Translate("MSC_WU", MessagesLabelsEN.MSC_WU);
                    if (val[6] == "N.A.")
                    { val[6] = "N/A"; }
                    row6["result"] = val[6];
                    DT.Rows.Add(row6);

                    DataRow row7 = DT.NewRow();
                    row7["parameter"] = LM.Translate("PMSC_WU", MessagesLabelsEN.PMSC_WU);
                    if (val[7] == "N.A.")
                    { val[7] = "N/A"; }
                    row7["result"] = val[7];
                    DT.Rows.Add(row7);

                    DataRow row8 = DT.NewRow();
                    row8["parameter"] = LM.Translate("FSC_WU", MessagesLabelsEN.FSC_WU);
                    if (val[8] == "N.A.")
                    { val[8] = "N/A"; }
                    row8["result"] = val[8];
                    DT.Rows.Add(row8);

                    DataRow row9 = DT.NewRow();
                    row9["parameter"] = LM.Translate("VELOCITY_WU", MessagesLabelsEN.VELOCITY_WU);
                    if (val[9] == "N.A.")
                    { val[9] = "N/A"; }
                    row9["result"] = val[9];
                    DT.Rows.Add(row9);

                    DataRow row10 = DT.NewRow();
                    row10["parameter"] = LM.Translate("SMI", MessagesLabelsEN.SMI);
                    if (val[10] == "N.A.")
                    { val[10] = "N/A"; }
                    row10["result"] = val[10];
                    DT.Rows.Add(row10);

                    DataRow row11 = DT.NewRow();
                    row11["parameter"] = LM.Translate("NUM_SPERM_WU", MessagesLabelsEN.NUM_SPERM_WU);
                    if (val[11] == "N.A.")
                    { val[11] = "N/A"; }
                    row11["result"] = val[11];
                    DT.Rows.Add(row11);

                    DataRow row12 = DT.NewRow();
                    row12["parameter"] = LM.Translate("MOTILE_SPERM_WU", MessagesLabelsEN.MOTILE_SPERM_WU);
                    if (val[12] == "N.A.")
                    { val[12] = "N/A"; }
                    row12["result"] = val[12];
                    DT.Rows.Add(row12);

                    DataRow row13 = DT.NewRow();
                    row13["parameter"] = LM.Translate("PROG_SPERM_WU", MessagesLabelsEN.PROG_SPERM_WU);
                    if (val[13] == "N.A.")
                    { val[13] = "N/A"; }
                    row13["result"] = val[13];
                    DT.Rows.Add(row13);

                    DataRow row14 = DT.NewRow();
                    row14["parameter"] = LM.Translate("FUNC_SPERM_WU", MessagesLabelsEN.FUNC_SPERM_WU);
                    if (val[14] == "N.A.")
                    { val[14] = "N/A"; }
                    row14["result"] = val[14];
                    DT.Rows.Add(row14);

                    DataRow row15 = DT.NewRow();
                    row15["parameter"] = LM.Translate("SPERM_WU", MessagesLabelsEN.SPERM_WU);
                    if (val[15] == "N.A.")
                    { val[15] = "N/A"; }
                    row15["result"] = val[15];
                    DT.Rows.Add(row15);

                }
                catch (Exception ex)
                {

                }
            }

        }
    }

    public class Average_report
    {




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
        public void AddTextCells(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, BaseColor bgrndcolor,int cellheight)
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

        public void AddTextCell1(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, int paddingleft, BaseColor bgrndcolor)
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
            cell.PaddingLeft = paddingleft;
            cell.PaddingBottom = paddingbottom;




            cell.BackgroundColor = bgrndcolor;

            table.AddCell(cell);

        }
        public void headercolor(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom)
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
            cell.BackgroundColor = new BaseColor(176, 224, 230);


            table.AddCell(cell);

        }


        public void averagerpt(List<string> reportData, List<string> row1_data, List<string> row2_data)
        {
            Logclass Logs = new Logclass();
            MainForm mf = new MainForm();
            Logs.LogFile_Entries("Generating Averaging Report", "Info");
            char temp = '\0';
            char prev = '\0';
            char befPrev = '\0';
            int pntCnt = 0; // count for points in string
            bool AVGflag = false;
            bool ODflag = false;
            bool CNTflag = false;
            bool CNTpreflag = false;
            iTextSharp.text.Image Arrow;
            string AVGStr = "AVG ";
            string ODStr = "OD ";
            string CNTStr = "CNT ";



            mf.emptylines();
            LanguageManagement LM = LanguageManagement.CreateInstance();
            //string path = Application.StartupPath + @"\settings.xml";
            //appSettings = XmlUtility.ReadSettings(path);
            string path = Application.StartupPath + @"\settings.xml";
            string[] appSettings = XmlUtility.ReadSettings(path);
            var dateTime = DateTime.Now;
            try
            {
                mf.immotile = float.Parse(reportData[16]);
                mf.prog = float.Parse(reportData[14]);
                mf.nonprog = float.Parse(reportData[15]);
            }
            catch { }
            try { mf.SMI_val = float.Parse(reportData[22]); } catch { }
            mf.LoadMorph();
            mf.LoadSMI();





            MyPdfPageEventHelpPageNo rptgenerator = new MyPdfPageEventHelpPageNo();

            var stringDate = dateTime.ToString().Replace('/', '_');
            stringDate = stringDate.Replace(':', '_');

            string filePath = Path.GetTempPath() + "Semen_Analysis_Report" + stringDate + ".pdf";


            Document doc = new Document(PageSize.A4, 5F, 5F, 0F, 45F);
            PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream(filePath, FileMode.Create));
            try
            {

               

                doc.Open();
                string head;
                string subhead;
                int headerspace;
                int age = 0;
                //string dob = "15/11/1986";
                float space;
                string trimtext = reportData[3].Substring(6);
                int year = int.Parse(trimtext);
                int fourDigitYear = CultureInfo.CurrentCulture.Calendar.ToFourDigitYear(year);
                
                if (reportData[3] != "00/00/00")
                {
                    try
                    {
                        DateTime userdate = DateTime.Parse(reportData[3]);
                        int usermonth = int.Parse(reportData[3].Substring(3, 2));
                        if ((DateTime.Now.Month >= usermonth) && DateTime.Now.Date >= userdate)
                        { age = DateTime.Now.Year - fourDigitYear; }
                        else
                        {
                            age = (DateTime.Now.Year - fourDigitYear) - 1;
                        }
                    }
                    catch (Exception ex)
                    {
                        age = (DateTime.Now.Year - fourDigitYear) - 1;
                    }
                }

                    head = LM.Translate("AVG_HEAD", MessagesLabelsEN.Average_Report_head);
                    subhead =  LM.Translate("AVG_SUB_HEAD", MessagesLabelsEN.Average_Report_subhead);





                
                
                //FontFactory.Register(fontPath);
                //BaseFont bf = BaseFont.CreateFont(fontPath, BaseFont.IDENTITY_H,BaseFont.EMBEDDED);
           
                iTextSharp.text.Font pfont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(),  13, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font headers = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 11, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font headers2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font headers3 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 9, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                iTextSharp.text.Font Charttxt = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                iTextSharp.text.Font pfont2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 11, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font datafont = FontFactory.GetFont(Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                iTextSharp.text.Font datafooter = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.GRAY);
                iTextSharp.text.Font datafont2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 10, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                if(appSettings[6]== "zh-CN")
                {
                    var fontPath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
                    FontFactory.Register(fontPath);
                    pfont1 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 13, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    headers = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 11, Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    headers2 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, Font.BOLD, iTextSharp.text.BaseColor.BLACK);
                    headers3 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9, Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                    Charttxt = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 12, Font.BOLD, iTextSharp.text.BaseColor.BLACK);

                    pfont2 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 11, Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                    datafont = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
                    datafooter = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8, Font.NORMAL, iTextSharp.text.BaseColor.GRAY);
                    datafont2 = FontFactory.GetFont(fontPath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10, Font.NORMAL, iTextSharp.text.BaseColor.BLACK);


                }
              
                Paragraph para = new Paragraph(head, pfont1);
                Paragraph para1 = new Paragraph(subhead, pfont2);
                para.Alignment = Element.ALIGN_CENTER;
                para1.Alignment = Element.ALIGN_CENTER;
                PdfPTable tab = new PdfPTable(2);
                Paragraph patinfo = new Paragraph(LM.Translate("PATIENT_INFO", MessagesLabelsEN.PATIENT_INFO), headers2);
                Paragraph SPACE = new Paragraph("\n");
                PdfContentByte pcb = wri.DirectContent;

                PdfPCell header1 = new PdfPCell(new Phrase(patinfo));

                header1.Colspan = 2;

                header1.BorderWidthBottom = 0.6f;
                header1.BorderWidthLeft = 0;
                header1.BorderWidthRight = 0;
                header1.BorderWidthTop = 0;
                header1.PaddingBottom = 8;

                header1.HorizontalAlignment = 0;
                if (reportData[2] == "N/A"|| reportData[2] == "---") { reportData[2] = " "; }
                PdfPCell cell1 = new PdfPCell(new Phrase(LM.Translate("PATIENT_ID", MessagesLabelsEN.PATIENT_ID) + " : ", datafont));
                PdfPCell cell1value = new PdfPCell(new Phrase(reportData[1], datafont));
                PdfPCell cell2 = new PdfPCell(new Phrase(LM.Translate("PAT_NAME", MessagesLabelsEN.PATIENT_NAME) + " : ", datafont));
                PdfPCell cell2value = new PdfPCell(new Phrase(reportData[2], datafont));
                tab.AddCell(header1);

                AddTextCell1(tab, cell1, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5,3, BaseColor.WHITE);
                AddTextCell1(tab, cell1value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5,30, BaseColor.WHITE);

                AddTextCell1(tab, cell2, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5,3, BaseColor.WHITE);
                AddTextCell1(tab, cell2value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5,30, BaseColor.WHITE);

           

                tab.TotalWidth = 266f;
                tab.WriteSelectedRows(0, -1, 20, 755  , pcb);
                tab.HorizontalAlignment = 2;


                //reference by doctor
                PdfPTable tab2 = new PdfPTable(2);
                PdfPCell header2 = new PdfPCell(new Phrase(" ", datafont));
                header2.Colspan = 2;

                header2.BorderWidthBottom = 0.6f;
                header2.BorderWidthLeft = 0;
                header2.BorderWidthRight = 0;
                header2.BorderWidthTop = 0;
                PdfPCell nill = new PdfPCell(new Phrase("  ", headers));

                PdfPCell fill = new PdfPCell(new Phrase(" ", datafont));
                //AddTextCell(tab2, cell3, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                //AddTextCell(tab2, cell3value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                PdfPCell cell4 = new PdfPCell(new Phrase(LM.Translate("BIRTH_DATE", MessagesLabelsEN.BIRTH_DATE) + " / " + LM.Translate("AGE", MessagesLabelsEN.AGE) + " : ", datafont));
                PdfPCell cell4value = new PdfPCell(new Phrase(reportData[3] + " " + "|" + " " + age.ToString(), datafont));

                PdfPCell cell5 = new PdfPCell(new Phrase(LM.Translate("TEST_DATE_WU", MessagesLabelsEN.TEST_DATE_WU) + " :", datafont));

               PdfPCell cell5value = new PdfPCell(new Phrase(reportData[0].Insert(reportData[0].Length-6,"|"), datafont));
                header1.HorizontalAlignment = 0;
                //tab2.HorizontalAlignment = 2;

                AddTextCell(tab2, nill, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(tab2, nill, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);

                AddTextCell1(tab2, cell4, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5,3, BaseColor.WHITE);
                AddTextCell1(tab2, cell4value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5,30, BaseColor.WHITE);

                //AddTextCell1(tab2, new PdfPCell(new Phrase(LM.Translate("TEST_DATE_WU", MessagesLabelsEN.TEST_DATE_WU)+" :", datafont)), Element.ALIGN_LEFT, 1, 2, 0, 0.6f, 0, 0, 5, 5,3, BaseColor.WHITE);
                AddTextCell1(tab2, cell5, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, 3, BaseColor.WHITE);
                AddTextCell1(tab2, cell5value, Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, 30, BaseColor.WHITE);




                tab2.TotalWidth = 266f;
                tab2.WriteSelectedRows(0, -1, 305, 758  , pcb);


                //sample information table
                PdfPTable tab3 = new PdfPTable(2);

                AddTextCell(tab3, new PdfPCell(new Phrase(LM.Translate("SAMPLE_INFO", MessagesLabelsEN.SAMPLE_INFO), headers2)), Element.ALIGN_LEFT, 1, 2, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);

                //headercolor(tab3, nill, Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0.6f, 0.6f, 5, 5);

                AddTextCell1(tab3, new PdfPCell(new Phrase(LM.Translate("SAMPLE_TYPE", MessagesLabelsEN.SAMPLE_TYPE) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(reportData[8], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(LM.Translate("ACCESSION", MessagesLabelsEN.ACCESSION) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(reportData[5], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);
            
                AddTextCell1(tab3, new PdfPCell(new Phrase(LM.Translate("LIQUIFACTION", MessagesLabelsEN.LIQUIFICTION) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(reportData[30], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);
           
                AddTextCell1(tab3, new PdfPCell(new Phrase(LM.Translate("FRUCTOSE", MessagesLabelsEN.LIQUIFICTION) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(reportData[31], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(LM.Translate("WBC_CONC_WU", MessagesLabelsEN.WBC_CONC_WU) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0.6f, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab3, new PdfPCell(new Phrase(reportData[10], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0.6f, 0, 0, 5, 5,30, BaseColor.WHITE);


               


                tab3.TotalWidth = 266f;

                tab3.WriteSelectedRows(0, -1, 20, 684  , pcb);

                tab3.HorizontalAlignment = 0;



                //Semen examination table



                PdfPTable tab4 = new PdfPTable(2);





                AddTextCell(tab4, new PdfPCell(new Phrase(" ", headers2)), Element.ALIGN_LEFT, 1, 2, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);

                //headercolor(tab4, new PdfPCell(new Phrase("RESULT", headers)), Element.ALIGN_CENTER, 1, 1, 0.6f, 0, 0.6f, 0.6f, 5, 5);

                AddTextCell1(tab4, new PdfPCell(new Phrase(LM.Translate("COLLECTED_DATE_WU", MessagesLabelsEN.COLLECTED_DATE_WU) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5, 03, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(reportData[6].Insert(reportData[6].Length - 6, "|"), datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5, 30, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(LM.Translate("RECEIVED_WU", MessagesLabelsEN.RECEIVED_WU) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(reportData[7].Insert(reportData[7].Length - 6, "|"), datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);


                AddTextCell1(tab4, new PdfPCell(new Phrase(LM.Translate("ABSTINENCE_WU", MessagesLabelsEN.ABSTINENCE_WU) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(reportData[4], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);

               
                AddTextCell1(tab4, new PdfPCell(new Phrase(LM.Translate("VOLUME_WU", MessagesLabelsEN.VOLUME) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(reportData[9], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0, 0, 0, 5, 5,30, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(LM.Translate("PH", MessagesLabelsEN.PH) + " :", datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0.6f, 0, 0, 5, 5,03, BaseColor.WHITE);

                AddTextCell1(tab4, new PdfPCell(new Phrase(reportData[11], datafont)), Element.ALIGN_LEFT, 1, 1, 0.6f, 0.6f, 0, 0, 5, 5,30, BaseColor.WHITE);

                tab4.TotalWidth = 266f;

                tab4.WriteSelectedRows(0, -1, 305, 684, pcb);



                //parameters table
               


                PdfPTable tab5 = new PdfPTable(65);



                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("PARAMETER", MessagesLabelsEN.PARAMETER), headers2)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                AddTextCell(tab5, new PdfPCell(new Phrase("1", headers2)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(176, 224, 230));
                AddTextCell(tab5, new PdfPCell(new Phrase("2", headers2)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(176, 224, 230));
                AddTextCell1(tab5, new PdfPCell(new Phrase(LM.Translate("Ref_val", MessagesLabelsEN.Ref_Val), headers2)), Element.ALIGN_CENTER, 1, 11, 0.6f, 0, 0, 0, 5, 5, 0, new BaseColor(176, 224, 230), 18, BaseColor.BLACK);
                AddTextCell1(tab5, new PdfPCell(new Phrase(LM.Translate("AVRG", MessagesLabelsEN.AVG_Val), headers2)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, 0, new BaseColor(176, 224, 230), 18, BaseColor.BLACK);
                
                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("CONC_WU", MessagesLabelsEN.CONC_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row1_data[0], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row2_data[0], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);


                

                AddTextCell(tab5, new PdfPCell(new Phrase(">= 15", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                string red_arrow_path = Application.StartupPath + @"\assets\RED.png";
                string green_arrow_path = Application.StartupPath + @"\assets\GREEN.png";

                try
                {
                    if (float.Parse(reportData[12]) >= 15)
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

                AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, BaseColor.WHITE);
                
                    headercolor(tab5, new PdfPCell(new Phrase(reportData[12], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                
                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("PR_NP_WU", MessagesLabelsEN.PR_NP_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[1], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[1], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                
                try
                {
                    if (float.Parse(reportData[13]) >= 40)
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

                headercolor(tab5, new PdfPCell(new Phrase(">= 40", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[13], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("PROG_WU", MessagesLabelsEN.PROG_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row1_data[2], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row2_data[2], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
              

                AddTextCell(tab5, new PdfPCell(new Phrase(">= 32", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                try
                {
                    if (float.Parse(reportData[14]) >= 32)
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


                AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(reportData[14], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("NONPROG_WU", MessagesLabelsEN.NONPROG_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[3], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[3], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
               
                headercolor(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[15], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("IMMOT_WU", MessagesLabelsEN.RESULT), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row1_data[4], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row2_data[4], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(reportData[16], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("WHO_5_WU", MessagesLabelsEN.WHO_5_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[5], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[5], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
               
                headercolor(tab5, new PdfPCell(new Phrase(">= 4", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                try
                {
                    if (float.Parse(reportData[17]) >= 4)
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


                AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[17], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("MSC_WU", MessagesLabelsEN.MSC_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(row1_data[6], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(row2_data[6], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

              

                AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(reportData[18], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("PMSC_WU", MessagesLabelsEN.PMSC_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[7], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[7], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                headercolor(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[19], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("FSC_WU", MessagesLabelsEN.FSC_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(row1_data[8], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(row2_data[8], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

        

                AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(reportData[20], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("VELOCITY_WU", MessagesLabelsEN.VELOCITY_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[9], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[9], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                headercolor(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[21], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("SMI", MessagesLabelsEN.SMI), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                if (reportData[22] == "N.A.")
                { reportData[22] = "N/A"; } 
                if (row1_data[10] == "N.A.")
                { row1_data[10] = "N/A"; } 
                if (row2_data[10] == "N.A.")
                { row2_data[10] = "N/A"; }


                headercolor(tab5, new PdfPCell(new Phrase(row1_data[10], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row2_data[10], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(reportData[22], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("NUM_SPERM_WU", MessagesLabelsEN.NUM_SPERM_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                if (reportData[23] == "N.A.")
                { reportData[23] = "N/A"; }
                if (row1_data[11] == "N.A.")
                { row1_data[11] = "N/A"; }
                if (row2_data[11] == "N.A.")
                { row2_data[11] = "N/A"; }
                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[11], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[11], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                headercolor(tab5, new PdfPCell(new Phrase(">= 39", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                try
                {
                    if (float.Parse(reportData[23]) >= 39)
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


                AddTextCell(tab5, new PdfPCell(Arrow), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 2, 2, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[23], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("MOTILE_SPERM_WU", MessagesLabelsEN.MOTILE_SPERM_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                if (reportData[24] == "N.A.")
                { reportData[24] = "N/A"; }
                if (row1_data[12] == "N.A.")
                { row1_data[12] = "N/A"; }
                if (row2_data[12] == "N.A.")
                { row2_data[12] = "N/A"; }
                headercolor(tab5, new PdfPCell(new Phrase(row1_data[12], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row2_data[12], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);

                headercolor(tab5, new PdfPCell(new Phrase(reportData[24], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("PROG_SPERM_WU", MessagesLabelsEN.PROG_SPERM_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                if (reportData[25] == "N.A.")
                { reportData[25] = "N/A"; }
                if (row1_data[13] == "N.A.")
                { row1_data[13] = "N/A"; }
                if (row2_data[13] == "N.A.")
                { row2_data[13] = "N/A"; }
                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[13], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[13], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                headercolor(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[25], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));
                AddTextCell(tab5, new PdfPCell(new Phrase(LM.Translate("FUNC_SPERM_WU", MessagesLabelsEN.FUNC_SPERM_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                if (reportData[26] == "N.A.")
                { reportData[26] = "N/A"; }
                if (row1_data[14] == "N.A.")
                { row1_data[14] = "N/A"; }
                if (row2_data[14] == "N.A.")
                { row2_data[14] = "N/A"; }
                headercolor(tab5, new PdfPCell(new Phrase(row1_data[14], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(row2_data[14], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0, 0, 0.6f, 5, 5, BaseColor.WHITE);

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0, 0, 0, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(reportData[26], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE);
                headercolor(tab5, new PdfPCell(new Phrase(LM.Translate("SPERM_WU", MessagesLabelsEN.SPERM_WU), datafont)), Element.ALIGN_LEFT, 1, 34, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                if (reportData[27] == "N.A.")
                { reportData[27] = "N/A"; }
                if (row1_data[15] == "N.A.")
                { row1_data[15] = "N/A"; }
                if (row2_data[15] == "N.A.")
                { row2_data[15] = "N/A"; }

                AddTextCell(tab5, new PdfPCell(new Phrase(row1_data[15], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0.6f, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(row2_data[15], datafont)), Element.ALIGN_CENTER, 1, 6, 0.6f, 0.6f, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                headercolor(tab5, new PdfPCell(new Phrase("---", datafont)), Element.ALIGN_CENTER, 1, 7, 0.6f, 0.6f, 0, 0.6f, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase("")), Element.ALIGN_CENTER, 1, 4, 0.6f, 0.6f, 0, 0, 5, 5, new BaseColor(243, 243, 243));

                AddTextCell(tab5, new PdfPCell(new Phrase(reportData[27], datafont)), Element.ALIGN_CENTER, 1, 8, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, new BaseColor(243, 243, 243));


                tab5.TotalWidth = 326f;


                tab5.WriteSelectedRows(0, -1, 20, 552  , pcb);

                tab5.HorizontalAlignment = 0;





                PdfPTable tabfooter = new PdfPTable(1);

                tabfooter.WidthPercentage = 45;
                if (reportData[28] == "N/A") { reportData[28] = " "; }
                if (reportData[29] == "N/A") { reportData[29] = " "; }
                AddTextCell(tabfooter, new PdfPCell(new Phrase(LM.Translate("TESTER_NAME", null) + ": " + reportData[28], datafont)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);

                AddTextCell(tabfooter, new PdfPCell(new Phrase(LM.Translate("DESIGNATION", MessagesLabelsEN.DESIGNATION) + ": " + reportData[29], datafont)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                tabfooter.HorizontalAlignment = 0;

                tabfooter.TotalWidth = 266f;

                tabfooter.WriteSelectedRows(0, -1, 20, 102  , pcb);

                
                try
                {
                    
                   
                    var phrase = new Phrase();
                  
                    
                    var boldFont = FontFactory.GetFont(FontFactory.HELVETICA, 8, Font.BOLD);
                    PdfPCell hcell3 = new PdfPCell(new Phrase(phrase));
                    var p1 =new Chunk(reportData[32], datafont);
                    
                   var p2 = new Chunk(LM.Translate("COMMENTS", MessagesLabelsEN.COMMENTS).ToUpper()+" :", headers2);
                    Paragraph comb = new Paragraph();
                    comb.Add(new Chunk(p2));
                    comb.Add(new Chunk(p1));
                    comb.Leading = 17;
                    hcell3.AddElement(comb);
                  
                    //phrase.Add();



                    PdfPTable comments = new PdfPTable(1);
                    AddTextCells(comments, hcell3, Element.ALIGN_LEFT, 1, 1, 0, 0, 0, 0, 0, 0, BaseColor.WHITE,50);
                    //hcell3.MultipliedLeading = 20;

                    comments.TotalWidth = 556f;
                    comments.WriteSelectedRows(0, -1, 20, 237, pcb);
                }
                catch (Exception ex)
                { 
                   

                }







                PdfPTable sign = new PdfPTable(1);

                sign.TotalWidth = 266f;

                AddTextCell(sign, new PdfPCell(new Phrase(LM.Translate("SIGNATURE", MessagesLabelsEN.SIGNATURE) + ": ", datafont)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);

                //AddCellheight(sign, new PdfPCell(png), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE, 1);

                sign.HorizontalAlignment = 5;

                sign.WriteSelectedRows(0, -1, 305, 84  , pcb);

                //footer Data

                string p = appSettings[10];





               
               


               

                doc.Add(SPACE);
                doc.Add(SPACE);
                doc.Add(para);
                doc.Add(para1);
                doc.Add(SPACE);
                //doc.Add(tab);
                PdfPTable Morphchart = new PdfPTable(2);
                headercolor(Morphchart, new PdfPCell(new Phrase(LM.Translate("Analysing_charts", MessagesLabelsEN.Analysing_Charts), headers2)), Element.ALIGN_CENTER, 1, 2, 0.6f, 0, 0.6f, 0.6f, 5, 5, new BaseColor(176, 224, 230));

                //if (refbydr.Text != "---" && refbydr2.Text != "---" && refbydr3.Text != "---")
                if (reportData[14] != "---" && reportData[15] != "---" && reportData[16] != "---"&& reportData[14] != "N/A" && reportData[15] != "N/A" && reportData[16] != "N/A")
                {
                    AddTextCell1(Morphchart, new PdfPCell(new Phrase(LM.Translate("MOTILITY", MessagesLabelsEN.Motility), Charttxt)), Element.ALIGN_LEFT, 1, 2, 0.6f, 0, 0.6f, 0.6f, 0, 0, 85, BaseColor.WHITE, 20, BaseColor.BLACK);

                    iTextSharp.text.Image motchart = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\Morphology_chart.png"); // ("Morphology_chart.png");

                    //motchart.ScaleToFit(190f, 200f);
                    AddTextCell1(Morphchart, new PdfPCell(motchart), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0.6f, 0.6f, 0, 40, 15, BaseColor.WHITE, 220, BaseColor.BLACK);

                    //LM.Translate("MOTILITY", MessagesLabelsEN.Motility)
                    //motchart.SetAbsolutePosition(doc.PageSize.Width - 250f, doc.PageSize.Height - 520f  );
                    //doc.Add(motchart);

                    Morphchart.TotalWidth = 227f;
                    Morphchart.WriteSelectedRows(0, -1, 345.02f, 552f, pcb);


                }
                else
                {
                    AddTextCell1(Morphchart, new PdfPCell(new Phrase("", headers)), Element.ALIGN_CENTER, 1, 2, 0.6f, 0.6f, 0.6f, 0.6f, 5, 5, 5, BaseColor.WHITE, 200, BaseColor.BLACK);
                    Morphchart.TotalWidth = 227f;
                    Morphchart.WriteSelectedRows(0, -1, 345.02f, 552f  , pcb);
                    PdfPTable NO_morph = new PdfPTable(2);
                    AddTextCell1(NO_morph, new PdfPCell(new Phrase(LM.Translate("No_Mot_Chart", MessagesLabelsEN.NO_MOTILITY_CHART).ToUpper(), headers2)), Element.ALIGN_CENTER, 1, 2, 3f, 3f, 3f, 3f, 5, 5, 5, BaseColor.WHITE, 100, new BaseColor(211, 211, 211));


                    NO_morph.TotalWidth = 170f;
                    NO_morph.WriteSelectedRows(0, -1, 369.01f, 515f, pcb);

                }




                iTextSharp.text.Image smichart = iTextSharp.text.Image.GetInstance(Application.StartupPath + @"\SMI_chart.png"); // ("Morphology_chart.png");

                PdfPTable Smichart = new PdfPTable(2);
                string Smi_head = "SMI";
                if (reportData[22] != "0" && reportData[22] != "---" && reportData[22] != "N/A")
                {
                    AddCellheight(Smichart, new PdfPCell(new Phrase(LM.Translate("SMI", MessagesLabelsEN.SMI), Charttxt)), Element.ALIGN_CENTER, 1, 2, 0, 0, 0.6f, 0.6f, 5, 5, BaseColor.WHITE, 0);
                    //AddTextCell(Smichart, new PdfPCell(new Phrase("", Charttxt)), Element.ALIGN_CENTER, 1, 25, 0, 0, 0, 0.6f, 5, 0, BaseColor.WHITE);
                    //smichart.ScaleToFit(500, 100);

                    AddTextCell1(Smichart, new PdfPCell(smichart), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 0, 0, BaseColor.WHITE, 115, BaseColor.BLACK);
                    Smichart.TotalWidth = 227f;

                    Smichart.WriteSelectedRows(0, -1, 345.02f, 383f, pcb);

                }
                else
                {
                    AddTextCell1(Smichart, new PdfPCell(new Phrase("", headers)), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0.6f, 0.6f, 5, 5, 5, BaseColor.WHITE, 137, BaseColor.BLACK);
                    Smichart.TotalWidth = 227f;

                    Smichart.WriteSelectedRows(0, -1, 345.02f, 383f  , pcb);

                    PdfPTable NO_smi = new PdfPTable(2);
                    AddTextCell1(NO_smi, new PdfPCell(new Phrase(LM.Translate("NO_SMI_CHART", MessagesLabelsEN.NO_SMI_CHART).ToUpper(), headers2)), Element.ALIGN_CENTER, 1, 2, 3f, 3f, 3f, 3f, 5, 5, 5, BaseColor.WHITE, 100, new BaseColor(211, 211, 211));
                    AddTextCell1(Smichart, new PdfPCell(new Phrase("", headers)), Element.ALIGN_CENTER, 1, 2, 0, 0.6f, 0, 0.6f, 5, 5, 5, BaseColor.WHITE, 137, BaseColor.BLACK);


                    NO_smi.TotalWidth = 170f;
                    NO_smi.WriteSelectedRows(0, -1, 369.01f, 375f, pcb);

                }



                //footer

                try
                {




                    string footerStr = LM.Translate("PRINTED_FROM", MessagesLabelsEN.PRINTED_FROM) + " " + UniversalStrings.QC_GOLD_ARCHIVE + " | " + LM.Translate("DEVICE_SN_WU", MessagesLabelsEN.DEVICE_SN_WU) + " " + reportData[33] + " | " + dateTime ;
         

                    PdfPTable footer = new PdfPTable(55);
                    footer.WidthPercentage = 45;
                    AddTextCell(footer, new PdfPCell(new Phrase(footerStr, datafooter)), Element.ALIGN_CENTER, 1, 55, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                    AddTextCell(footer, new PdfPCell(new Phrase(" ", datafont)), Element.ALIGN_CENTER, 1, 0, 0, 0, 0, 0, 5, 5, BaseColor.WHITE);
                    footer.HorizontalAlignment = 0;
                    Rectangle page = doc.PageSize;
                    footer.TotalWidth = page.Width - doc.LeftMargin - doc.RightMargin;
                    footer.WriteSelectedRows(0, -1, doc.LeftMargin, footer.TotalHeight + 25, pcb);
                    //footer.TotalWidth = 500f;
                    //footer.WriteSelectedRows(0, -1, 230, 80 - 25, pcb);


                }
                catch { }








                doc.Add(SPACE);
                // doc.Add(tab3);

                doc.Add(SPACE);
                // doc.Add(tab5);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);
                doc.Add(Chunk.NEWLINE);
                // doc.Add(tabfooter);



            }
            
            catch (Exception Ex)
            {
                Logs.LogFile_Entries(Ex.ToString() , "Error");

            }


            doc.Close();
            System.Diagnostics.Process.Start(filePath);
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

    }




}

public class Logclass
{
    StreamWriter log;
    static string stringDate;
    string path = System.Windows.Forms.Application.StartupPath;
    string dateTime = DateTime.Now.ToString();
    public void LogFile()
    {

        var dateTime = DateTime.Now;

        stringDate = dateTime.ToString().Replace('/', '_');
        stringDate = stringDate.Replace(':', '_');
        log = new StreamWriter(path + @"\LogFiles\logfile" + stringDate + ".txt");
        log.Close();
    }
    public void LogFile_Entries(string Logdata, string Error_info)
    {

        StreamWriter SR = File.AppendText(path + @"\LogFiles\logfile" + stringDate + ".txt");

        SR.WriteLine(DateTime.Now + "\t" + "[" + Error_info + "]:" + "\t" + Logdata);

        SR.WriteLine("\n");
        // Close the stream:
        SR.Close();
    }
}

public class PdfReportGenerator
{
    public string ResourcesStr = "", PageLabelM = "", OfLabelM = "";
    private const int TEXT_LINE_LENGTH = 52;
    private const int SPACE_BEFORE_TABLE = 15;
    private Stopwatch stopWatch = new Stopwatch();
    public String ElapsedTime = "";
    public BaseFont FontType;
    public string de;

    internal string pa;
    Document doc = new Document(PageSize.A4, 5F, 5F, 0F, 45F);



    iTextSharp.text.Font pfont1 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 13, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
    iTextSharp.text.Font headers = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 11, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
    iTextSharp.text.Font headers2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);
    iTextSharp.text.Font headers3 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 9, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

    iTextSharp.text.Font Charttxt = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 12, iTextSharp.text.Font.BOLD, iTextSharp.text.BaseColor.BLACK);

    iTextSharp.text.Font pfont2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 11, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
    iTextSharp.text.Font datafont = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
    iTextSharp.text.Font datafooter = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 8, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.GRAY);
    iTextSharp.text.Font datafont2 = FontFactory.GetFont(iTextSharp.text.Font.FontFamily.HELVETICA.ToString(), 10, iTextSharp.text.Font.NORMAL, iTextSharp.text.BaseColor.BLACK);
           


public bool GenerateSEMENANALYSISReport(float headerSpace, string ReportFileNamePath, string ReportFileName, String reportname, String reportnamecontinued, string footerStr, String DeviceInformationHeader, DataTable DeviceInformationTable,  /*String logoUrl, String facilityTableName, DataTable headertable, String footerStr,*/ String PatientInformationHeader, DataTable PatientInformationTable, String SampleInformationHeader, DataTable SampleInformationTable, DataRow SemAnalParametersHeader, DataTable SemAnalParametersTable/*, DataRow ImageVideoAvaliableHeader, DataTable ImageVideoAvaliableTable, DataTable CommentsRow, DataRow SystemParametersHeader, DataTable SystemParametersTable, String AdditionalDataHeader, DataTable AdditionalDataTable, String MorphTitle, DataTable MorphHeaders, DataTable MorphTable, DataTable MorphImgTable, String VitalityTitle, DataTable VitalityHeaders, DataTable VitalityTable, DataTable VitalityImgTable, String DNAFragTitle, DataTable MainDNAFragHeaders, DataTable MainDNAFragTable, DataTable SubDNAFragHeaders, DataTable SubDNAFragTable, DataTable DNAFragImgTable, String CaptureImgTitle, DataTable CaptureImgTable, int PicturesInPage, bool IsSampleInformationOnAllPages*/)
    {
        //PdfPTable table;
        stopWatch.Start();

        float spacer = iTextSharp.text.Utilities.MillimetersToPoints(headerSpace);
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        //CultureInfo currentCulture = new CultureInfo("he-IL");
        Font Arial = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10F, Font.NORMAL, BaseColor.BLACK);
        Font ArialHead = new Font(FontType, 12F, Font.BOLD, BaseColor.BLACK);
        Font ArialSmall = new Font(FontType, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(FontType, 8F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(FontType, 8F, Font.NORMAL, BaseColor.RED);
        Font ArialBold = new Font(FontType, 10F, Font.BOLD, BaseColor.BLACK);

        ///CloseIfPdfOpen(ReportFileName, 25);

        Document doc = new Document(PageSize.A4, 5F, 5F, 0F, 45F);
        PdfWriter pdfWriter = PdfWriter.GetInstance(doc, new FileStream(ReportFileNamePath, FileMode.Create));

        MyPdfPageEventHelpPageNo e = new MyPdfPageEventHelpPageNo()
        {

            ISLowerHeader = false,
            ReportName = reportname,
            ReportNameContinued = reportnamecontinued,
            FooterStr = footerStr, // added
            PageLabel = PageLabelM,
            fontType = FontType,
            OfLable = OfLabelM,
            HeaderSpace = spacer // in points not in mm (headerSpace)

        };


        pdfWriter.PageEvent = e;





        doc.Open();












        /// 
        /// 
        // create Tables
        ///


        CreateDeviceInformationTable(doc, DeviceInformationTable, DeviceInformationHeader);
        e.IsPageExceed = true;

        CreatePatientAndQCInformationTable(doc, PatientInformationTable, PatientInformationHeader);
        e.IsPageExceed = true;

        CreateSampleInformationLongevitySemenAnalysisTable(doc, SampleInformationTable, SampleInformationHeader);
        e.IsPageExceed = true;

        CreateParameterLongevitySemenAnalysisTable(doc, SemAnalParametersHeader, SemAnalParametersTable, false);
        e.IsPageExceed = true;

        ///
        doc.Close();
        ///
        saveElapsedTime();
        return true;



    }


    private void CloseIfPdfOpen(string PDFName, byte ReportLen)
    {
        int i = 0;
        foreach (Process proc in Process.GetProcessesByName("AcroRd32"))
        {
            string tmpStr = proc.MainWindowTitle;

            if (tmpStr != "")
            {
                if (tmpStr.Length >= ReportLen)
                {
                    if (tmpStr.Substring(0, ReportLen) == PDFName)
                    {
                        proc.Kill();
                        proc.Dispose();
                    }
                }
            }
            i++;
        }
    }

    private void addheader(Document mdoc, DataTable Table, String TableHeader)
    {
        //CultureInfo currentCulture = new CultureInfo("he-IL");

        Font ArialSmall = new Font(FontType, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallRowSpan = new Font(FontType, 6F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(FontType, 6F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(FontType, 6F, Font.NORMAL, BaseColor.RED);
        Font ArialSmallwhite = new Font(FontType, 6F, Font.NORMAL, BaseColor.WHITE);
        Font ArialBold = new Font(FontType, 10F, Font.BOLD, BaseColor.BLACK);
        Font ArialBoldBlue = new Font(FontType, 10F, Font.BOLD, new BaseColor(0, 68, 136));

        PdfPTable table1 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table1.WidthPercentage = 35;

        PdfPTable table2 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table2.WidthPercentage = 35;

        int count = 0;
        foreach (DataRow row in Table.Rows)
        {

            if (count % 2 == 0)
            {
                AddTextCell(table1, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table1, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }
            else
            {
                AddTextCell(table2, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table2, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }

            count++;
        }
        if (count % 2 == 1)
        {
            AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
        }
        PdfPTable outer = new PdfPTable(new float[] { 0.4F, 0.02F, 0.4F });

        PdfPCell outercell = new PdfPCell(new Phrase(TableHeader, ArialBold));
        outercell.PaddingTop = 5;
        outercell.PaddingBottom = 5;
        outercell.Colspan = 3;
        outercell.HorizontalAlignment = Element.ALIGN_LEFT;
        outercell.VerticalAlignment = Element.ALIGN_MIDDLE;
        outercell.BorderWidth = 0f;
        outercell.BorderWidthBottom = 0.6f;
        outer.AddCell(outercell);

        outer.DefaultCell.BorderWidth = 0f;
        outer.AddCell(table1);

        PdfPCell GAPcell = new PdfPCell(new Phrase("", ArialBold));
        GAPcell.BorderWidth = 0f;
        outer.AddCell(GAPcell);


        outer.AddCell(table2);
        outer.SpacingBefore = SPACE_BEFORE_TABLE;
        outer.WidthPercentage = 90;

        mdoc.Add(outer);

    }

    private void CreateDeviceInformationTable(Document mdoc, DataTable Table, String TableHeader)
    {
        //CultureInfo currentCulture = new CultureInfo("he-IL");
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        Font ArialSmall = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallRowSpan = new Font(FontType, 6F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(FontType, 6F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(FontType, 6F, Font.NORMAL, BaseColor.RED);
        Font ArialSmallwhite = new Font(FontType, 6F, Font.NORMAL, BaseColor.WHITE);
        Font ArialBold = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10F, Font.BOLD, BaseColor.BLACK);
        Font ArialBoldBlue = new Font(FontType, 10F, Font.BOLD, new BaseColor(0, 68, 136));

        PdfPTable table1 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table1.WidthPercentage = 35;

        PdfPTable table2 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table2.WidthPercentage = 35;

        int count = 0;
        foreach (DataRow row in Table.Rows)
        {

            if (count % 2 == 0)
            {
                AddTextCell(table1, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table1, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }
            else
            {
                AddTextCell(table2, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table2, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }

            count++;
        }
        if (count % 2 == 1)
        {
            AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
        }
        PdfPTable outer = new PdfPTable(new float[] { 0.4F, 0.02F, 0.4F });

        PdfPCell outercell = new PdfPCell(new Phrase(TableHeader, ArialBold));
        outercell.PaddingTop = 5;
        outercell.PaddingBottom = 5;
        outercell.Colspan = 3;
        outercell.HorizontalAlignment = Element.ALIGN_LEFT;
        outercell.VerticalAlignment = Element.ALIGN_MIDDLE;
        outercell.BorderWidth = 0f;
        outercell.BorderWidthBottom = 0.6f;
        outer.AddCell(outercell);

        outer.DefaultCell.BorderWidth = 0f;
        outer.AddCell(table1);

        PdfPCell GAPcell = new PdfPCell(new Phrase("", ArialBold));
        GAPcell.BorderWidth = 0f;
        outer.AddCell(GAPcell);


        outer.AddCell(table2);
        outer.SpacingBefore = SPACE_BEFORE_TABLE;
        outer.WidthPercentage = 90;

        mdoc.Add(outer);

    }
    private void CreatePatientAndQCInformationTable(Document mdoc, DataTable Table, String TableHeader)
    {

        //CultureInfo currentCulture = new CultureInfo("he-IL");
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        Font ArialSmall = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallRowSpan = new Font(FontType, 6F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(FontType, 6F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(FontType, 6F, Font.NORMAL, BaseColor.RED);
        Font ArialSmallwhite = new Font(FontType, 6F, Font.NORMAL, BaseColor.WHITE);
        Font ArialBold = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10F, Font.BOLD, BaseColor.BLACK);
        Font ArialBoldBlue = new Font(FontType, 10F, Font.BOLD, new BaseColor(0, 68, 136));

        PdfPTable table1 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table1.WidthPercentage = 35;

        PdfPTable table2 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table2.WidthPercentage = 35;

        int count = 0;
        foreach (DataRow row in Table.Rows)
        {

            if (count % 2 == 0)
            {
                AddTextCell(table1, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table1, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }
            else
            {
                AddTextCell(table2, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table2, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }

            count++;
        }
        if (count % 2 == 1)
        {
            AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
        }
        PdfPTable outer = new PdfPTable(new float[] { 0.4F, 0.02F, 0.4F });

        PdfPCell outercell = new PdfPCell(new Phrase(TableHeader, ArialBold));
        outercell.PaddingTop = 5;
        outercell.PaddingBottom = 5;
        outercell.Colspan = 3;
        outercell.HorizontalAlignment = Element.ALIGN_LEFT;
        outercell.VerticalAlignment = Element.ALIGN_MIDDLE;
        outercell.BorderWidth = 0f;
        outercell.BorderWidthBottom = 0.6f;
        outer.AddCell(outercell);

        outer.DefaultCell.BorderWidth = 0f;
        outer.AddCell(table1);

        PdfPCell GAPcell = new PdfPCell(new Phrase("", ArialBold));
        GAPcell.BorderWidth = 0f;
        outer.AddCell(GAPcell);


        outer.AddCell(table2);
        outer.SpacingBefore = SPACE_BEFORE_TABLE;
        outer.WidthPercentage = 90;

        mdoc.Add(outer);

    }

    private void CreateSampleInformationLongevitySemenAnalysisTable(Document mdoc, DataTable Table, String TableHeader)
    {

        //CultureInfo currentCulture = new CultureInfo("he-IL");
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        Font ArialSmall =FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallRowSpan = new Font(FontType, 6F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(FontType, 6F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(FontType, 6F, Font.NORMAL, BaseColor.RED);
        Font ArialSmallwhite = new Font(FontType, 6F, Font.NORMAL, BaseColor.WHITE);
        Font ArialBold = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10F, Font.BOLD, BaseColor.BLACK);
        Font ArialBoldBlue = new Font(FontType, 8F, Font.BOLD, new BaseColor(0, 68, 136));

        PdfPTable table1 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table1.WidthPercentage = 35;

        PdfPTable table2 = new PdfPTable(new float[] { 0.5F, 0.5F });
        table2.WidthPercentage = 35;



        int count = 0;
        foreach (DataRow row in Table.Rows)
        {
            if (count % 2 == 0)
            {
                AddTextCell(table1, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table1, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }
            else
            {
                AddTextCell(table2, new PdfPCell(new Phrase(row["parameter"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
                AddTextCell(table2, new PdfPCell(new Phrase(row["value"].ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            }

            count++;
        }
        if (Table.Rows.Count % 2 != 0)
        {
            AddTextCell(table1, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
            AddTextCell(table1, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6f, 0, 0, 5, 5, BaseColor.WHITE);
        }
        AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0, 0, 0, 0, 0, BaseColor.WHITE);
        AddTextCell(table2, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0, 0, 0, 0, 0, BaseColor.WHITE);

        PdfPTable outer = new PdfPTable(new float[] { 0.4F, 0.02F, 0.4F });

        PdfPCell outercell = new PdfPCell(new Phrase(TableHeader, ArialBold));
        outercell.PaddingTop = 5;
        outercell.PaddingBottom = 5;
        outercell.Colspan = 3;
        outercell.HorizontalAlignment = Element.ALIGN_LEFT;
        outercell.VerticalAlignment = Element.ALIGN_MIDDLE;
        outercell.BorderWidth = 0f;
        outercell.BorderWidthBottom = 0.6f;
        outer.AddCell(outercell);

        outer.DefaultCell.BorderWidth = 0f;
        outer.AddCell(table1);

        PdfPCell GAPcell = new PdfPCell(new Phrase("", ArialBold));
        GAPcell.BorderWidth = 0f;
        outer.AddCell(GAPcell);

        outer.AddCell(table2);
        outer.SpacingBefore = SPACE_BEFORE_TABLE;
        outer.WidthPercentage = 90;

        mdoc.Add(outer);
    }

    private void CreateParameterLongevitySemenAnalysisTable(Document mdoc, DataRow Headers, DataTable Table, bool IsLongevity)
    {
        PdfPTable table;
        PdfPCell cell;
        //CultureInfo currentCulture = new CultureInfo("he-IL");
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        Font ArialSmall = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallRowSpan = new Font(FontType, 7F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(FontType, 7F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(FontType, 7F, Font.NORMAL, BaseColor.RED);
        Font ArialSmallBlue = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 7F, Font.NORMAL, new BaseColor(0, 68, 136));
        Font ArialBold = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10F, Font.BOLD, BaseColor.BLACK);
        Font ArialBoldBlue = new Font(FontType, 8F, Font.BOLD, new BaseColor(0, 68, 136));

        float[] headersparam = new float[Headers.ItemArray.Length];
        int headerscounter = 0;
        for (int j = 0; j < Headers.ItemArray.Length; j++)
        {
            Array.Resize(ref headersparam, headerscounter + 1);
            if (headerscounter == Headers.ItemArray.Length - 1)
                headersparam[headerscounter] = 0.1f;
            else if (headerscounter == 0)
            {
                if (IsLongevity)
                    headersparam[headerscounter] = 0.3f;
                else
                    headersparam[headerscounter] = 0.15f;
            }
            else
                headersparam[headerscounter] = 0.1f;
            headerscounter++;
        }
        table = new PdfPTable(headersparam);
        table.WidthPercentage = 90;
        table.SpacingBefore = SPACE_BEFORE_TABLE;
        int headercount = 0;
        foreach (object obj in Headers.ItemArray)
        {
            if (headercount == 0)
            {
                cell = new PdfPCell(new Phrase(obj.ToString(), ArialBold));
                cell.PaddingTop = 5;
                cell.PaddingBottom = 5;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_LEFT;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = 0.6f;
                table.AddCell(cell);
            }
            else
            {
                cell = new PdfPCell(new Phrase(obj.ToString(), ArialBold));
                cell.PaddingTop = 5;
                cell.PaddingBottom = 5;
                cell.Colspan = 1;
                cell.HorizontalAlignment = Element.ALIGN_CENTER;
                cell.VerticalAlignment = Element.ALIGN_MIDDLE;
                cell.BorderWidth = 0f;
                cell.BorderWidthBottom = 0.6f;
                table.AddCell(cell);
            }
            headercount++;
        }


        int i = 0;
        int ColCounter = 0;
        while (i < Table.Rows.Count)
        {
            ColCounter = 0;
            foreach (object obj in Table.Rows[i].ItemArray)
            {
                cell = new PdfPCell(new Phrase("", ArialBold));
                switch (obj.ToString().ToLower())
                {
                    case "0":
                        if (Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "status" || Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "resultsflag")
                            AddTextCell(table, new PdfPCell(new Phrase("", ArialSmallBlue)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        else
                            AddTextCell(table, new PdfPCell(new Phrase(obj.ToString(), ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                    case "1":
                        if (Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "status" || Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "resultsflag")
                        {
                            //table.AddCell(ImagecellUp);
                        }
                        else
                            AddTextCell(table, new PdfPCell(new Phrase(obj.ToString(), ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                    case "2":
                        if (Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "status" || Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "resultsflag")
                        {

                            //table.AddCell(ImagecellDown);
                        }
                        else
                            AddTextCell(table, new PdfPCell(new Phrase(obj.ToString(), ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                    case "3":
                        if (Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "status" || Table.Rows[i].Table.Columns[ColCounter].ColumnName.ToLower() == "resultsflag")
                        {
                            //table.AddCell(ImagecellExclamtionMark);
                        }
                        else
                            AddTextCell(table, new PdfPCell(new Phrase(obj.ToString(), ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                    case "empty":
                        AddTextCell(table, new PdfPCell(new Phrase(" ", ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                    case "":
                        AddTextCell(table, new PdfPCell(new Phrase("", ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                    default:
                        if (ColCounter == 0)
                            AddTextCell(table, new PdfPCell(new Phrase(obj.ToString(), ArialSmall)), Element.ALIGN_LEFT, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        else
                            AddTextCell(table, new PdfPCell(new Phrase(obj.ToString(), ArialSmall)), Element.ALIGN_CENTER, 1, 1, 0, 0.6F, 0, 0, 5, 5, BaseColor.WHITE);
                        break;
                }
                ColCounter++;
            }
            i++;
        }
        //table.KeepTogether = true;
        table.HeaderRows = 1;

        mdoc.Add(table);
    }

    //Horizontal alignment 0=Left, 1=Center, 2=Right
    // Adds cell to table
    private void AddTextCell(PdfPTable table, PdfPCell cell, int cellHorizontalAlignment, int RowSpan, int ColSpan, float borderWidthTop, float borderWidthBottom, float borderWidthLeft, float borderWidthRight, int paddingtop, int paddingbottom, BaseColor bgrndcolor)
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


    private void saveElapsedTime()
    {
        stopWatch.Stop();
        ElapsedTime = stopWatch.ElapsedMilliseconds.ToString() + " mSeconds";
    }

    public bool GenerateSEMENANALYSISReport(string v1, string v2, string v3, string v4, string v5, DataTable deviceInformationTable, string v6, DataTable patientInformationTable, string v7, DataTable sampleInformationTable, DataRow semAnalParametersDR, DataTable semAnalParametersTable)
    {
        throw new NotImplementedException();
    }
}

public class MyPdfPageEventHelpPageNo : iTextSharp.text.pdf.PdfPageEventHelper
{
    protected PdfTemplate total;
    List<PdfTemplate> pageNumberTemplates = new List<PdfTemplate>();
    protected BaseFont cour;
    public bool IsPageExceed = false;
    public String FooterStr { get; set; }
    public bool ISLowerHeader { get; set; }
    //public String LogoURL { get; set; }
    public String ReportName { get; set; }
    public String ReportNameContinued { get; set; }
    //public DataTable HeaderTable { get; set; }
    //public String HeaderTableName { get; set; }
    public String Device { get; set; }
    public String SWVersion { get; set; }
    public String PageLabel { get; set; }
    public String OfLable { get; set; }
    public BaseFont fontType { get; set; }
    public float HeaderSpace { get; set; }
    public override void OnOpenDocument(PdfWriter writer, Document document)
    {
        total = writer.DirectContent.CreateTemplate(PageSize.A4.Width, PageSize.A4.Height);
        total.BoundingBox = new Rectangle(0, 0, PageSize.A4.Width, PageSize.A4.Height);

        cour = BaseFont.CreateFont(BaseFont.HELVETICA, BaseFont.WINANSI, BaseFont.NOT_EMBEDDED);
    }

    public override void OnStartPage(PdfWriter writer, Document document)
    {

        base.OnStartPage(writer, document);
        ///
        PdfPTable table;
        PdfPCell cell;
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        //CultureInfo currentCulture = new CultureInfo("he-IL");
        //Font Arial = FontFactory.GetFont(fontpath, 10F, Font.NORMAL, BaseColor.BLACK);
        Font ArialHead = FontFactory.GetFont(fontpath, 1F, Font.NORMAL, BaseColor.BLACK);

        table = new PdfPTable(new float[] { 0.1f });
        table.WidthPercentage = 90;
        if (ISLowerHeader)
            table.SpacingAfter = 40;
        else
            table.SpacingAfter = 20;

        cell = new PdfPCell(new Phrase("    ", ArialHead));
        cell.Colspan = 1;
        cell.BackgroundColor = BaseColor.WHITE;
        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
        cell.BorderWidth = 0f;
        table.AddCell(cell);
        document.Add(table);
        ///
 
        if (IsPageExceed)
        {
            CreateReportHeader(document, /*logo,*/ ReportNameContinued, /*"", null,*/ Device, SWVersion, writer);
        }
        else
        {
            CreateReportHeader(document, /*logo,*/ ReportName, /*HeaderTableName, HeaderTable,*/ Device, SWVersion, writer);
        }
        IsPageExceed = true;
    }
    private void CreateReportHeader(Document mdoc, /*byte[] logo,*/ String ReportName, /*string FacilityTableName, DataTable HeaderTable,*/ String Device, String SWVersion, PdfWriter writer)
    {
        PdfPTable table;
        PdfPCell cell;
        //CultureInfo currentCulture = new CultureInfo("he-IL");
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
        Font Arial = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 10F, Font.NORMAL, BaseColor.BLACK);
        Font ArialHead = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 11F, Font.BOLD, BaseColor.BLACK);
        Font ArialHeadPage = new Font(fontType, 12F, Font.NORMAL, BaseColor.BLACK);
        Font ArialHeadPageBold = new Font(fontType, 12F, Font.BOLD, BaseColor.BLACK);
        Font ArialSmall = new Font(fontType, 9F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallRowSpan = new Font(fontType, 7F, Font.NORMAL, BaseColor.BLACK);
        Font ArialSmallGreen = new Font(fontType, 7F, Font.NORMAL, new BaseColor(0, 100, 0));
        Font ArialSmallRed = new Font(fontType, 7F, Font.NORMAL, BaseColor.RED);
        Font ArialSmallBlue = new Font(fontType, 8F, Font.UNDERLINE, BaseColor.BLUE/*new BaseColor(0, 68, 136)*/);
        Font ArialBold = new Font(fontType, 10F, Font.BOLD, BaseColor.BLACK);


        //if (HeaderTable != null)
        //{
        table = new PdfPTable(new float[] { 0.1f });
        table.WidthPercentage = 90;
        table.SpacingBefore = HeaderSpace;

        cell = new PdfPCell(new Phrase(ReportName, ArialHead));
        cell.Colspan = 3;
        cell.BackgroundColor = BaseColor.WHITE;
        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
        cell.BorderWidth = 0f;
        table.AddCell(cell);
        //mdoc.Add(table);

        // ADD DEVICE AND SW VERSION


        //}

        cell = new PdfPCell(new Phrase(Device, ArialHead));
        cell.Colspan = 3;
        cell.HorizontalAlignment = Element.ALIGN_LEFT;
        cell.VerticalAlignment = Element.ALIGN_TOP;
        cell.BorderWidth = 0.0f;
        table.AddCell(cell);

  

        mdoc.Add(table);
    }




    // OnCloseDocumet
    public override void OnEndPage(PdfWriter writer, Document document)
    {
        // holds today's date in dd/MM/yyyy format
        float adjust;
        float textBase;
        float textSize;
        string text;
        PdfContentByte cb = writer.DirectContent;
        cb.SaveState();
        text = FooterStr;
        textBase = document.Bottom - 20;//-20;
        textSize = FooterStr.Length; //helv.GetWidthPoint(text, 12);
        cb.BeginText();
        cb.SetFontAndSize(cour, 8);

        adjust = cour.GetWidthPoint(text, 6);
        cb.SetTextMatrix((document.Right - textSize - adjust) / 2, textBase);
        //cb.SetTextMatrix(document.Left, textBase);
        //cb.ShowText(text);
        cb.EndText();
        cb.AddTemplate(total, document.Left + 107, textBase);

        PdfTemplate t = writer.DirectContent.CreateTemplate(180, 50);
        pageNumberTemplates.Add(t);
        writer.DirectContent.AddTemplate(t, document.LeftMargin + 520, document.PageSize.GetTop(document.BottomMargin) - 15);

        cb.RestoreState();
       
        PdfPTable table = new PdfPTable(new float[] { 1.0f });
        var fontpath = Application.StartupPath + "\\assets\\arial-unicode-ms.ttf";
    
        Font ArialHeadPage = FontFactory.GetFont(fontpath, BaseFont.IDENTITY_H, BaseFont.EMBEDDED, 8F, Font.NORMAL, new BaseColor(128, 128, 128));
        table.WidthPercentage = 0;
        table.SpacingBefore = 0;
        table.SpacingAfter = 0;

        table.HorizontalAlignment = Element.ALIGN_RIGHT;
        PdfPCell cell = new PdfPCell(new Phrase(text, ArialHeadPage));
        cell.Colspan = 1;
        cell.BackgroundColor = BaseColor.WHITE;
        cell.HorizontalAlignment = Element.ALIGN_CENTER;
        cell.VerticalAlignment = Element.ALIGN_MIDDLE;
        cell.PaddingBottom = 30;
        cell.BorderWidth = 0f;
        cell.PaddingTop = 0;
        table.AddCell(cell);
        Rectangle page = document.PageSize;
        table.TotalWidth = page.Width - document.LeftMargin - document.RightMargin;
        table.WriteSelectedRows(0, -1, document.LeftMargin, table.TotalHeight + 7, writer.DirectContent);
        //document.Add(table);
        //cb.RestoreState();
    }



   



}













