using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Globalization;
using System.IO;
using System.Diagnostics;
using System.Threading;
using System.Windows.Forms;
using QcGoldArchive;

namespace QcGoldArchive
{
    public delegate void ProcessError(string exceptionMessage, string title);
    public class Export
    {
        LanguageManagement LM = LanguageManagement.CreateInstance();

        StreamWriter streamWriter = null;
        public event ProcessError OnProcessError;
        string LineStr = "";
        private char dsChar;
        string path;
        int totalcols;
        FileStream aFile1 = null;
        //string NextColStr=",";
        private CultureInfo ExcelCultureInfo;
        private string dateFormat;
        string DateTimeFormat;
        string DateFormat;
        string[] appSettings;
        Logclass Logs = new Logclass();
        public Export(char mdsChar, CultureInfo mExcelCultureInfo, string mdateFormat)
        {
            dsChar = mdsChar;
            ExcelCultureInfo = mExcelCultureInfo;
            dateFormat = mdateFormat;
            SetDateFormat();
        }

        private void SetDateFormat()
        {
            if (dateFormat == "0")            // Parse dates as mm/dd or dd/mm
            {
                DateTimeFormat = UniversalStrings.DateTimeFormatEU;
                DateFormat = UniversalStrings.DateFormatEU;

            }
            else
            {
                DateTimeFormat = UniversalStrings.DateTimeFormatUS;
                DateFormat = UniversalStrings.DateFormatUS;
            }

        }

        /// <summary>
        /// CSV
        /// </summary>
        /// <param name="filepath"></param>
        /// <param name="ExportType"></param>
        /// 
        public void Archive_Translate(string path)
        {
            int count = 0;
            try
            {
                FileStream aFile = null;


                string LineStr;
                if (dsChar == char.Parse(","))
                    LineStr = ";";
                else
                    LineStr = ",";

                var retainedLines = File.ReadAllLines(path)
                       .Skip(1) // if you have an header in your file and don't want it
                       .Where(x => x.Split(char.Parse(LineStr))[1] != "clear");

                File.Delete(path);
                streamWriter = new StreamWriter(path, false);//filename, true);
                string[] labels = UniversalStrings.NEW_FILE_INDIAN_LABELS_WHO6;

                foreach (string label in labels)
                {
                    WriteToFile(label.ToUpper());
                    NextColumn();
                }
                NextRow();



                CloseFile();
                File.AppendAllLines(path, retainedLines);
                ///// updating type in csv
                ///
                List<String> lines = new List<String>();
                using (StreamReader reader = new StreamReader(path))
                {
                    String line;



                    while ((line = reader.ReadLine()) != null)
                    {

                        if (line.Contains(LineStr))
                        {

                            String[] split = line.Split(char.Parse(LineStr));

                            if (split[8].Contains("FRESH") || split[8].Contains("FRISCH") || split[8].Contains("FRESCO") || split[8].Contains("FRAIS") || split[8].Contains("新鲜"))
                            {

                                split[8] = LM.Translate("FRESH", MessagesLabelsEN.FRESH);
                                line = String.Join(LineStr, split);
                            }
                            else if (split[8].Contains("FROZEN") || split[8].Contains("CONGELATO") || split[8].Contains("GEFROREN") || split[8].Contains("CONGELE") || split[8].Contains("冷冻"))
                            {

                                split[8] = LM.Translate("FROZEN", MessagesLabelsEN.FRESH);
                                line = String.Join(LineStr, split);
                            }
                            else if (split[8].Contains("WASHED") || split[8].Contains("GEWASCHEN") || split[8].Contains("LAVATO") || split[8].Contains("LAVE") || split[8].Contains("洗涤"))
                            {

                                split[8] = LM.Translate("WASHED", MessagesLabelsEN.FRESH);
                                line = String.Join(LineStr, split);

                            }


                        }

                        lines.Add(line);
                        count++;
                    }
                }

                using (StreamWriter writer = new StreamWriter(path, false))
                {
                    foreach (String line in lines)
                        writer.WriteLine(line);
                }





            }
            catch (Exception ex)
            {
                //MessageBox.Show(LM.Translate("Archive_Empty", MessagesLabelsEN.Archive_Error_Missing), LM.Translate("", UniversalStrings.QC_GOLD_ARCHIVE));
            }
        }

        public void linecount()
        {

            string path1 = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path1);
            var path = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";



            var lines = File.ReadAllLines(path);
            if (File.Exists(path))
            {

                char LineStr;
                if (dsChar == char.Parse(","))
                    LineStr = ';';
                else
                    LineStr = ',';

                totalcols = lines[0].Split(LineStr).Length;
                if (totalcols < 40)
                {
                    File.Delete(path);
                }
            }

        }
        public void Archive_Translate1(string path)
        {
            try
            {


                FileStream aFile = null;




                var retainedLines = File.ReadAllLines(path)
                       .Skip(1) // if you have an header in your file and don't want it
                       .Where(x => x.Split(',')[1] != "clear");
                File.Delete(path);
                streamWriter = new StreamWriter(path, false);//filename, true);
                string[] labels = UniversalStrings.NEW_FILE_INDIAN_LABELS_WHO6;

                foreach (string label in labels)
                {
                    WriteToFile(label.ToUpper());
                    NextColumn();
                }
                NextRow();



                CloseFile();
                File.AppendAllLines(path, retainedLines);
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.ToString());
            }
        }



        public void FillCsv(string path, List<string> data, byte expType, FileStream aFile, bool isManual = false)
        {
            string logdata = "";
            for (int i = 0; i <= data.Count - 1; i++)
            {
                logdata += data[i] + "\n";
            }
            Logs.LogFile_Entries("Data before saving to archive " + logdata, "info");
            logdata = "";
            //Export expCsv = null;
            //DateTime temp;
            //byte expType;
            double num;

            if (path != "")
            {

                //    if (File.Exists(path))
                //    {
                //        expType = 2;
                //        StartCsvExport(path, expType);
                //    }
                //    else
                //    {
                //        expType = 1;
                //        StartCsvExport(path, expType);

                //        
                //    }
                //    else
                //        




                if (expType == 2)
                {



                    streamWriter = new StreamWriter(aFile);

                }
                else if (expType == 1)
                {
                    streamWriter = new StreamWriter(path, false);//filename, true);
                    string[] labels = UniversalStrings.NEW_FILE_INDIAN_LABELS_WHO6;

                    foreach (string label in labels)
                    {
                        WriteToFile(label.ToUpper());
                        NextColumn();
                    }
                    NextRow();
                }
                if (isManual == true)
                {
                    for (int i = 0; i < data.Count; i++)
                    {
                        //if ((i == 2) || (i == 4) || (i == 7) || (i == 8))
                        //    {

                        //        if (DateTime.TryParseExact(data[i], DateTimeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp)) // DateTimeFormat scope of Export class
                        //        {
                        //            data[i] = temp.ToString();   //Parsing to DateTime
                        //        }
                        //        else if (DateTime.TryParseExact(data[i], DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp))
                        //            data[i] = temp.ToString();   //Parsing to Date
                        //        //else
                        //            //data[i] = UniversalStrings.EMPTY_VALUE;
                        //    }

                            //data[i] = num.ToString();
                            if (dsChar != char.Parse("."))
                                data[i] = data[i].Replace('.', ',');
  

                        WriteToFile(data[i]);
                        NextColumn();
                    }
                }
                else
                {
                    for (int i = 2; i < data.Count; i++)
                    {
                        //if ((i == 2) || (i == 4) || (i == 7) || (i == 8))
                        //    {

                        //        if (DateTime.TryParseExact(data[i], DateTimeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp)) // DateTimeFormat scope of Export class
                        //        {
                        //            data[i] = temp.ToString();   //Parsing to DateTime
                        //        }
                        //        else if (DateTime.TryParseExact(data[i], DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp))
                        //            data[i] = temp.ToString();   //Parsing to Date
                        //        //else
                        //            //data[i] = UniversalStrings.EMPTY_VALUE;
                        //    }

                        if (i != 2 & i != 3 & i != 4 & i != 7 & i != 8 & data[i] != "N.A."/*& (Double.TryParse(data[i], out num))*/)
                        {
                            //data[i] = num.ToString();
                            if (dsChar != char.Parse("."))
                                data[i] = data[i].Replace('.', ',');
                        }

                        WriteToFile(data[i]);
                        NextColumn();
                    }
                }
                WriteToFile(data[1]);
                NextColumn();
                WriteToFile(data[0]);
                NextRow();

                CloseFile();
                Logs.LogFile_Entries("Data filled in csv", "Info");
                //Process.Start(path);
            }
        }
        //private void StartCsvExport(string filepath,byte ExportType)
        //{

        //    if (ExportType == 2)
        //    {
        //        FileStream aFile;
        //        try
        //        {
        //            aFile = new FileStream(filepath, FileMode.Append, FileAccess.Write);
        //        }
        //        catch
        //        {
        //            string filename;
        //            filename = GetFileName(filepath);
        //            while (IsFileOpen(filename))
        //            {
        //                if (MessageBox.Show(new Form() { TopMost = true }, filepath + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
        //                {
        //                    //IsValidFile = false;
        //                    MessageBox.Show(new Form() { TopMost = true }, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE);
        //                    return;
        //                }
        //                //throw new FieldAccessException();
        //                //MessageBox.Show( LM.Translate("FILE_OPEN_ERROR", MessagesLabelsEN.FILE_OPEN_ERROR), LM.Translate("EXPORT_ERROR_CAPTION", MessagesLabelsEN.EXPORT_ERROR_CAPTION));
        //            }

        //            //CloseIfReportOpen(filename, filename.Length);
        //            aFile = new FileStream(filepath, FileMode.Append, FileAccess.Write);
        //        }

        //    //    sw = new  StreamWriter(filename, true); //File.AppendText(filename);//
        //        streamWriter = new StreamWriter(aFile);//filename, true);
        //    }
        //    else
        //        streamWriter = new StreamWriter(filepath, false);
        //}

        private string GetFileName(string FilePath)
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

        private void CloseIfReportOpen(string FileName, int ReportLen)
        {
            foreach (Process proc in Process.GetProcessesByName("Excel"))
            {
                string tmpStr = proc.MainWindowTitle;

                if (tmpStr != "")
                {
                    if (tmpStr.Length >= ReportLen)
                    {
                        if (tmpStr.Substring(0, ReportLen) == FileName)
                        {
                            proc.Kill();
                            proc.Dispose();
                            //proc.CloseMainWindow();
                            //proc.WaitForExit(); //  the request to the operating system to terminate the associated process might not be handled if the process is written to never enter its message loop.

                        }
                    }
                }
            }
        }

        public void WriteToFile(string ValueData)
        {

            if (dsChar == char.Parse("."))
                ValueData = ValueData.Replace(",", ";");
            else
                ValueData = ValueData.Replace(";", ",");

            //  sw.Write(ValueData.Replace("\r\n", " "));
            LineStr = LineStr + ValueData.Replace("\r\n", " ");
        }
        public void NextColumn()
        {
            //sw.Write(dsChar.ToString());
            if (dsChar == char.Parse(","))
                LineStr = LineStr + ";";
            else
                LineStr = LineStr + ",";
        }
        public void NextRow()
        {
            //   sw.Write(sw.NewLine);

            streamWriter.WriteLine(LineStr);
            LineStr = "";
        }

        public void CloseFile()
        {
            streamWriter.Close();
        }

        /// <summary>
        /// Excel
        /// </summary>
        /// <param name="path"></param>

        public bool MakeNewExcelFile(string path)
        {
            string[] labels = UniversalStrings.NEW_FILE_INDIAN_LABELS_WHO6;
            bool success = false;

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook;

            try
            {
                workbook = (Microsoft.Office.Interop.Excel.Workbook)app.Workbooks.Add(System.Reflection.Missing.Value);
            }

            catch
            {
                Thread.CurrentThread.CurrentCulture = ExcelCultureInfo;
                workbook = (Microsoft.Office.Interop.Excel.Workbook)app.Workbooks.Add(System.Reflection.Missing.Value);
            }

            Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];
            Microsoft.Office.Interop.Excel.Range workSheet_range = null;

            app.ErrorCheckingOptions.NumberAsText = false;
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["A", Type.Missing]).ColumnWidth = 15;
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["B", Type.Missing]).ColumnWidth = 20;
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["F", Type.Missing]).ColumnWidth = 15;
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["G", Type.Missing]).ColumnWidth = 15;
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["Q", Type.Missing]).ColumnWidth = 14;
            ((Microsoft.Office.Interop.Excel.Range)worksheet.Columns["AA", Type.Missing]).ColumnWidth = 10;

            workSheet_range = worksheet.get_Range("A1", "AC1");
            workSheet_range.Merge(28);
            workSheet_range.WrapText = true;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.HorizontalAlignment = -4108;
            workSheet_range.VerticalAlignment = -4108;
            workSheet_range.Font.Bold = true;
            workSheet_range.Font.Size = 12;
            worksheet.Cells[1, 1] = "RESULTS";


            workSheet_range = worksheet.get_Range("A2", "D2");
            workSheet_range.Merge(5);
            workSheet_range.HorizontalAlignment = -4108;
            workSheet_range.Font.Bold = true;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 192, 192));
            worksheet.Cells[2, 1] = LM.Translate("PATIENT_DATA", MessagesLabelsEN.PATIENT_DATA);

            workSheet_range = worksheet.get_Range("E2", "K2");
            workSheet_range.Merge(5);
            workSheet_range.HorizontalAlignment = -4108;
            workSheet_range.Font.Bold = true;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(150, 150, 150));
            worksheet.Cells[2, 5] = LM.Translate("SAMPLE_DATA", MessagesLabelsEN.SAMPLE_DATA);

            workSheet_range = worksheet.get_Range("L2", "V2");
            workSheet_range.Merge(11);
            workSheet_range.HorizontalAlignment = -4108;
            workSheet_range.Font.Bold = true;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 192, 192));
            worksheet.Cells[2, 12] = LM.Translate("TEST_RESULTS", MessagesLabelsEN.TEST_RESULTS);

            workSheet_range = worksheet.get_Range("W2", "AA2");
            workSheet_range.Merge(5);
            workSheet_range.HorizontalAlignment = -4108;
            workSheet_range.Font.Bold = true;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(150, 150, 150));
            worksheet.Cells[2, 23] = LM.Translate("TOTALS_PER_VOL", MessagesLabelsEN.TOTALS_PER_VOL);

            workSheet_range = worksheet.get_Range("AB2", "AC2");
            workSheet_range.Merge(2);
            workSheet_range.Font.Bold = true;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 192, 192));
            workSheet_range.HorizontalAlignment = -4108;
            worksheet.Cells[2, 28] = LM.Translate("OTHER", MessagesLabelsEN.OTHER);

            workSheet_range = worksheet.get_Range("A3", "D3");    // Patient data
            workSheet_range.HorizontalAlignment = -4108;          //alignment - center
            workSheet_range.VerticalAlignment = -4107;           // alignment - bottom 
            workSheet_range.WrapText = true;
            workSheet_range.Font.Bold = true;
            workSheet_range.Font.Size = 8;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 192, 192));

            workSheet_range = worksheet.get_Range("E3", "K3");    //Sample data
            workSheet_range.HorizontalAlignment = -4108;          //alignment - center
            workSheet_range.VerticalAlignment = -4107;           // alignment - bottom 
            workSheet_range.Font.Bold = true;
            workSheet_range.WrapText = true;
            workSheet_range.Font.Size = 8;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(150, 150, 150));

            workSheet_range = worksheet.get_Range("L3", "V3");    // Test results
            workSheet_range.HorizontalAlignment = -4108;          //alignment - center
            workSheet_range.VerticalAlignment = -4107;           // alignment - bottom 
            workSheet_range.WrapText = true;
            workSheet_range.Font.Bold = true;
            workSheet_range.Font.Size = 8;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 192, 192));

            workSheet_range = worksheet.get_Range("W3", "AA3");   // Totals per volume
            workSheet_range.HorizontalAlignment = -4108;          //alignment - center
            workSheet_range.VerticalAlignment = -4107;           // alignment - bottom 
            workSheet_range.WrapText = true;
            workSheet_range.Font.Bold = true;
            workSheet_range.Font.Size = 8;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(150, 150, 150));

            workSheet_range = worksheet.get_Range("AB3", "AC3");  // Version+Device
            workSheet_range.HorizontalAlignment = -4108;          //alignment - center
            workSheet_range.VerticalAlignment = -4107;           // alignment - bottom 
            workSheet_range.WrapText = true;
            workSheet_range.Font.Bold = true;
            workSheet_range.Font.Size = 8;
            workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();
            workSheet_range.Interior.Color = System.Drawing.ColorTranslator.ToOle(System.Drawing.Color.FromArgb(192, 192, 192));


            for (int i = 1; i <= labels.Length; i++)
            {
                worksheet.Cells[3, i] = labels[i - 1];
            }

            try
            {
                workbook.SaveAs(path, Microsoft.Office.Interop.Excel.XlFileFormat.xlWorkbookNormal, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Microsoft.Office.Interop.Excel.XlSaveAsAccessMode.xlNoChange, Type.Missing, Type.Missing, Type.Missing, Type.Missing, Type.Missing);
                success = true;
            }

            catch (Exception e)
            {
                OnProcessError(e.Message, LM.Translate("EXCEL_FILE_ERROR_CAPTION", MessagesLabelsEN.EXCEL_FILE_ERROR_CAPTION));
                //MessageBox.Show(new Form() { TopMost = true }, e.Message, LM.Translate("EXCEL_FILE_ERROR_CAPTION", MessagesLabelsEN.EXCEL_FILE_ERROR_CAPTION));
                success = false;
            }

            workbook.Close(false, false, Type.Missing);
            app.Quit();
            return success;
        }


        public void FillExcel(string path, List<string> data)
        {
            double num;
            DateTime temp;
            int rowIndex;
            int i = 2;
            string MinCell = String.Empty;
            string MaxCell = String.Empty;
            //string DateTimeFormat;
            //string DateFormat;
            bool IsValidFile = true;

            //while (CanAccessFile(path) == false)
            //{
            //    if (MessageBox.Show(new Form() {TopMost = true}, path + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
            //    {
            //        IsValidFile = false;
            //        MessageBox.Show(new Form() { TopMost = true }, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE);
            //        return;
            //    }

            //}

            //if (IsValidFile == true)
            //{
            //if (dateFormat == "0")            // Parse dates as mm/dd or dd/mm
            //{
            //    DateTimeFormat = UniversalStrings.DateTimeFormatEU;
            //    DateFormat = UniversalStrings.DateFormatEU;
            //}
            //else
            //{
            //    DateTimeFormat = UniversalStrings.DateTimeFormatUS;
            //    DateFormat = UniversalStrings.DateFormatUS;
            //}

            Microsoft.Office.Interop.Excel.Application app = new Microsoft.Office.Interop.Excel.Application();
            Microsoft.Office.Interop.Excel.Workbook workbook;

            try
            {
                workbook = (Microsoft.Office.Interop.Excel.Workbook)app.Workbooks.Add(System.Reflection.Missing.Value);
            }

            catch
            {
                Thread.CurrentThread.CurrentCulture = ExcelCultureInfo;
                workbook = (Microsoft.Office.Interop.Excel.Workbook)app.Workbooks.Add(System.Reflection.Missing.Value);
            }

            Microsoft.Office.Interop.Excel.Range workSheet_range = null;
            try
            {
                workbook = app.Workbooks.Open(path, 0, false, 5, "", "", false, Microsoft.Office.Interop.Excel.XlPlatform.xlWindows, "",
                                              true, false, 0, true, false, false);

                Microsoft.Office.Interop.Excel.Worksheet worksheet = (Microsoft.Office.Interop.Excel.Worksheet)workbook.Sheets[1];


                rowIndex = worksheet.UsedRange.Count / 29 + 1;
                MinCell = "A" + rowIndex.ToString();
                MaxCell = "AC" + rowIndex.ToString();

                workSheet_range = worksheet.get_Range(MinCell, MaxCell);
                workSheet_range.HorizontalAlignment = -4108;          // horizontal alignment is center
                workSheet_range.Font.Size = 8;
                workSheet_range.WrapText = true;
                workSheet_range.Cells.WrapText = true;
                workSheet_range.Borders.Color = System.Drawing.Color.Black.ToArgb();

                for (i = 2; i < data.Count; i++)
                {
                    if ((i == 2) || (i == 4) || (i == 7) || (i == 8))
                    {

                        if (DateTime.TryParseExact(data[i], DateTimeFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp))
                        {
                            worksheet.Cells[rowIndex, i - 1] = temp;   //Parsing to DateTime
                        }
                        else if (DateTime.TryParseExact(data[i], DateFormat, CultureInfo.InvariantCulture, DateTimeStyles.None, out temp))
                            worksheet.Cells[rowIndex, i - 1] = temp;   //Parsing to Date
                        else
                            worksheet.Cells[rowIndex, i - 1] = UniversalStrings.EMPTY_VALUE;
                    }
                    else
                    {
                        if (dsChar != char.Parse("."))
                        {
                            if (data[i] != "N.A.")
                                data[i] = data[i].Replace('.', ',');
                        }

                        if ((i != 3) && (Double.TryParse(data[i], out num)))
                        {           //Parsing to double
                            worksheet.Cells[rowIndex, i - 1] = num;
                        }
                        else if (i == 3)
                            worksheet.Cells[rowIndex, i - 1] = "'" + data[i];        // Parse from number to string
                        else
                            worksheet.Cells[rowIndex, i - 1] = data[i];
                    }

                }
                worksheet.Cells[rowIndex, i - 1] = data[1];
                worksheet.Cells[rowIndex, i] = data[0];

                workbook.Close(true, Type.Missing, Type.Missing);
                app.Quit();

            }

            catch (Exception e)
            {
                //MessageBox.Show(new Form() { TopMost = true }, LM.Translate("EXPORT_FILE_ERROR", MessagesLabelsEN.EXPORT_FILE_ERROR), LM.Translate("EXCEL_FILE_ERROR_CAPTION", MessagesLabelsEN.EXCEL_FILE_ERROR_CAPTION));
                workbook.Close(false, Type.Missing, Type.Missing);
                app.Quit();
                OnProcessError(LM.Translate("EXPORT_FILE_ERROR", MessagesLabelsEN.EXPORT_FILE_ERROR), LM.Translate("EXCEL_FILE_ERROR_CAPTION", MessagesLabelsEN.EXCEL_FILE_ERROR_CAPTION));
            }
            //}
        }

        private bool CanAccessFile(string path)
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
    }




    public class OldCsv_Support
    {
        LanguageManagement LM = LanguageManagement.CreateInstance();

        string LineStr = "";
        StreamWriter streamWriter = null;
        public char dsChar = Convert.ToChar(CultureInfo.CurrentCulture.NumberFormat.NumberDecimalSeparator);

        public void old_csv(string path)
        {
            try
            {

                FileStream aFile = null;


                string LineStr;
                if (dsChar == char.Parse(","))
                    LineStr = ";";
                else
                    LineStr = ",";

                var retainedLines = File.ReadAllLines(path)
                       .Skip(1) // if you have an header in your file and don't want it
                       .Where(x => x.Split(char.Parse(LineStr))[1] != "clear");

                //File.Delete(path);
                int len = path.Length;
                string oldfile_path = path.Insert(len - 4, "_old");
                File.Move(path, oldfile_path);
                streamWriter = new StreamWriter(path, false);//filename, true);
                string[] labels = UniversalStrings.NEW_FILE_INDIAN_LABELS_WHO6;

                foreach (string label in labels)
                {
                    WriteToFile(label);
                    NextColumn();
                }
                NextRow();



                CloseFile();

                File.AppendAllLines(path, retainedLines);

                //Adding --- to fields

                List<String> lines = new List<String>();
                using (StreamReader reader = new StreamReader(path, true))
                {
                    String line;
                    int count = 0;


                    while ((line = reader.ReadLine()) != null)
                    {

                        if (line.Contains(LineStr) && count >= 1)
                        {

                            String[] split = line.Split(char.Parse(LineStr));
                            //string oldstring = split[2];
                            split[2] = "---" + LineStr + split[2] + "";
                            line = String.Join(LineStr, split);
                            split[27] = "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "N.A" + LineStr + "N.A" + LineStr + "N.A" + LineStr + "N.A" + LineStr + split[27] + " ";
                            line = String.Join(LineStr, split);




                        }

                        lines.Add(line);
                        count++;
                    }
                }

                using (StreamWriter writer = new StreamWriter(path, false))
                {
                    foreach (String line in lines)
                        writer.WriteLine(line);
                }



            }
            catch (Exception ex)
            { }





        }
        void WriteToFile(string ValueData)
        {

            if (dsChar == char.Parse("."))
                ValueData = ValueData.Replace(",", ";");
            else
                ValueData = ValueData.Replace(";", ",");

            //  sw.Write(ValueData.Replace("\r\n", " "));
            LineStr = LineStr + ValueData.Replace("\r\n", " ");
        }
        void NextColumn()
        {
            //sw.Write(dsChar.ToString());
            if (dsChar == char.Parse(","))
                LineStr = LineStr + ";";
            else
                LineStr = LineStr + ",";
        }
        void NextRow()
        {
            //   sw.Write(sw.NewLine);

            streamWriter.WriteLine(LineStr);
            LineStr = "";
        }

        void CloseFile()
        {
            streamWriter.Close();
        }

        public void IN_old_csv(string path)
        {
            try
            {

                FileStream aFile = null;


                string LineStr;
                if (dsChar == char.Parse(","))
                    LineStr = ";";
                else
                    LineStr = ",";

                var retainedLines = File.ReadAllLines(path)
                       .Skip(1) // if you have an header in your file and don't want it
                       .Where(x => x.Split(char.Parse(LineStr))[1] != "clear");

                //File.Delete(path);
                int len = path.Length;
                string oldfile_path = path.Insert(len - 4, "_IN_old");
                File.Move(path, oldfile_path);
                streamWriter = new StreamWriter(path, false);//filename, true);
                string[] labels = UniversalStrings.NEW_FILE_INDIAN_LABELS_WHO6;
                foreach (string label in labels)
                {
                    WriteToFile(label);
                    NextColumn();
                }
                NextRow();
                CloseFile();

                File.AppendAllLines(path, retainedLines);

                //Adding --- to fields

                List<String> lines = new List<String>();
                using (StreamReader reader = new StreamReader(path, true))
                {
                    String line;
                    int count = 0;


                    while ((line = reader.ReadLine()) != null)
                    {

                        if (line.Contains(LineStr) && count >= 1)
                        {

                            String[] split = line.Split(char.Parse(LineStr));
                            //string oldstring = split[2];
                            line = String.Join(LineStr, split);
                            split[39] = "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + "---" + LineStr + split[39] + " ";
                            line = String.Join(LineStr, split);


                        }

                        lines.Add(line);
                        count++;
                    }
                }

                using (StreamWriter writer = new StreamWriter(path, false))
                {
                    foreach (String line in lines)
                        writer.WriteLine(line);
                }



            }
            catch (Exception ex)
            { }
        }

    }
}
