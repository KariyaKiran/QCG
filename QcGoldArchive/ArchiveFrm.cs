
using CsvHelper;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

using System.Text.RegularExpressions;
using System.Data.OleDb;
using iTextSharp;
using iTextSharp.text.pdf;
using System.Threading;
using iTextSharp.text;
using Font = iTextSharp.text.Font;
using System.Globalization;
using System.Web.UI.WebControls;
using Label = System.Windows.Forms.Label;
using Button = System.Windows.Forms.Button;
using TextBox = System.Windows.Forms.TextBox;

namespace QcGoldArchive
{
    public partial class ArchiveFrm : Form
    {

        static int SelectColumnIndex = 0;
        MainForm mf;
        Export export;
        ToolStrip ctl11 = new ToolStrip();
        ToolTip t1 = new ToolTip();
        private ToolTip _toolTip = new ToolTip();
        private Control _currentToolTipControl = null;
        char separator = ';';
        //System.Windows.Forms.CheckBox check = null;
        public ArchiveFrm(MainForm MyParent_)
        {
            InitializeComponent();

            mf = MyParent_;
            LM.GetSubControlsData(this);
            //this.Average.

            //t1.Show(string.Format(LM.Translate("Archive_tooltip", MessagesLabelsEN.SAVE), Environment.NewLine), generate);
            generate.ToolTipText = string.Format(LM.Translate("Archive_tooltip", MessagesLabelsEN.SAVE), Environment.NewLine);

        }

       
        Logclass Logs = new Logclass();
        DataGridViewRow row,row2,row3,temprow;
        int count = 0;
        string test;
        bool Key_up = false;
        List<string> _row1 = new List<string>();
        List<string> _row2 = new List<string>();
        List<string> avg_rpt = new List<string>();
        List<string> _avg = new List<string>();
        bool IsShown = false;
      
        DataGridViewCellEventArgs f;
        private string[] appSettings = XmlUtility.ReadSettings(Application.StartupPath + @"\settings.xml");
        BindingSource bs = new BindingSource();
        LanguageManagement LM = LanguageManagement.CreateInstance();
        public DataTable ReadCsv(string name)
        {
            //DataTable dt = new DataTable("Data");

            //using (OleDbConnection cn = new OleDbConnection("Provider=Microsoft.ACE.OLEDB.12.0; Data Source=\"" + Path.GetDirectoryName(name) + "\"; Extended Properties = 'text;HDR=yes;FMT=Delimited(,)';"))
            //{
            //    using (OleDbCommand cmd = new OleDbCommand(string.Format("select * from [{0}]", new FileInfo(name).Name), cn))
            //    {
            //        cn.Open();
            //        using (OleDbDataAdapter adapter = new OleDbDataAdapter(cmd))
            //        {
            //            adapter.Fill(dt);
            //        }
            //    }
            //    return dt;
            //}

            // your code here 
            
                
            char LineStr;
            if (mf.dsChar == char.Parse(","))
                LineStr = ';';
            else
                LineStr = ',';

            FileStream aFile = null;
            String path = appSettings[2].Substring(0, appSettings[2].Length - 3) + "csv";
            List<String> lines = new List<String>();
            DataTable dt = new DataTable();
            try
            {
                aFile = new FileStream(path, FileMode.Append, FileAccess.Read);
                string CSVFilePathName = name;


                string[] Lines = File.ReadAllLines(CSVFilePathName);

                string[] Fields;
                Fields = Lines[0].Split(new char[] { LineStr });
                int Cols = Fields.GetLength(0);

                //1st row must be column names; force lower case to ensure matching later on.
                for (int i = 0; i < Cols - 1; i++)
                    dt.Columns.Add(Fields[i], typeof(string));
                DataRow Row;
                for (int i = 1; i < Lines.GetLength(0); i++)
                {
                    Fields = Lines[i].Split(new char[] { LineStr });
                    Row = dt.NewRow();
                    for (int f = 0; f < Cols - 1; f++)
                        Row[f] = Fields[f];
                    dt.Rows.Add(Row);
                }
                csvdata.DataSource = dt;
                Logs.LogFile_Entries("CSV file loaded" + " \tReadCsv in" + this.FindForm().Name, "Info");
            }
            catch
            {
                string filename;
                filename = mf.GetFileName(path);
                bool cancelled = false;
                //while (IsFileOpen(filename) && cancelled == false)
                while (mf.CanAccessFile(path) == false && cancelled == false)
                {
                    this.Invoke(new Action(() =>
                    {
                        Logs.LogFile_Entries(MessagesLabelsEN.ACCESS_ERROR + " \tReadCsv in" + this.FindForm().Name, "Error");
                        if (MessageBox.Show(this, path + LM.Translate("ACCESS_ERROR", MessagesLabelsEN.ACCESS_ERROR), UniversalStrings.QC_GOLD_ARCHIVE, MessageBoxButtons.RetryCancel) == DialogResult.Cancel)
                        {
                            this.Invoke(new Action(() => { MessageBox.Show(this, LM.Translate("IMPORT_FAILED", MessagesLabelsEN.IMPORT_FAILED), UniversalStrings.QC_GOLD_ARCHIVE); }));
                            cancelled = true;
                            Logs.LogFile_Entries(MessagesLabelsEN.IMPORT_FAILED + " \tReadCsv in" + this.FindForm().Name, "Error");
                        };
                    }));
                }

                if (!cancelled)
                {
                    try
                    {
                        string CSVFilePathName = name;


                        string[] Lines = File.ReadAllLines(CSVFilePathName);

                        string[] Fields;
                        Fields = Lines[0].Split(new char[] { LineStr });
                        int Cols = Fields.GetLength(0);

                        //1st row must be column names; force lower case to ensure matching later on.
                        for (int i = 0; i < Cols-1; i++)
                            dt.Columns.Add(Fields[i], typeof(string));
                        DataRow Row;
                        for (int i = 1; i < Lines.GetLength(0); i++)
                        {
                            Fields = Lines[i].Split(new char[] { LineStr });
                            Row = dt.NewRow();
                            for (int f = 0; f < Cols-1; f++)
                                Row[f] = Fields[f];
                            dt.Rows.Add(Row);
                        }
                        csvdata.DataSource = dt;
                        Logs.LogFile_Entries("CSV file loaded" + " \tReadCsv in" + this.FindForm().Name, "Info");

                    }
                    //catch(Exception ex)
                    //{
                    //    // MessageBox.Show(LM.Translate("Wrong_data_error", MessagesLabelsEN.WRONG_DATA_ERROR));
                    //}
                    catch (Exception ex)
                    {
                        // MessageBox.Show(LM.Translate("Wrong_data_error", MessagesLabelsEN.WRONG_DATA_ERROR));
                       
                    }
                

                }
               
            }



                


            
            return dt;
        }



        /* private void Csvselect_Click(object sender, EventArgs e)
         {



             Document doc = new Document(PageSize.LETTER, 10, 10, 42, 35);
             PdfWriter wri = PdfWriter.GetInstance(doc, new FileStream("Test.pdf", FileMode.Create));
             doc.Open();
             PdfPTable table = new PdfPTable(5);
             for (int j = 0; j < 5; j++)
             {
                 table.AddCell(new Phrase(csvdata.Columns[j].HeaderText));
             }

             table.HeaderRows = 2;

             for (int i = 0; i < 5; i++)
             {
                 for (int k = 0; k < 5; k++)
                 {
                     if (csvdata[k, i].Value != null)

                     {
                         table.AddCell(new Phrase(csvdata[k, i].Value.ToString()));
                     }

                 }
             }

             doc.Add(table);
             Paragraph para = new Paragraph("Hiii.\n");
             doc.Add(para);


             doc.Close();

         }*/
       

        private void DesignExample_Load(object sender, EventArgs e)
        {
            combofilter.Items.Add(LM.Translate("Arc_tesdate", MessagesLabelsEN.TEST_DATE));
            combofilter.Items.Add(LM.Translate("Arc_Accession", MessagesLabelsEN.ACCESSION));
            combofilter.Items.Add(LM.Translate("Arc_Pat_ID", MessagesLabelsEN.PATIENT_ID));
            combofilter.Items.Add(LM.Translate("Arc_Pat_Name", MessagesLabelsEN.PATIENT_NAME));
            
            //check = new System.Windows.Forms.CheckBox();
            //check.Size = new Size(15,15);
            //this.csvdata.Controls.Add(check);
            //chk.HeaderText = "";
            //chk.Name = "chk";
            //csvdata.Columns.Add(chk1);


            string path;

            path = appSettings[2];
            string csvpath = path.Substring(0, path.Length - 3) + "csv";
            // path = Application.StartupPath+ @"\QC_gold_archive.csv";
            
            try
            {
                csvdata.DataSource = ReadCsv(csvpath);
            }
            catch {  }
           
            try
            {
                string dt = csvdata.Columns[0].DefaultCellStyle.Format;
                if (dt == "MM-dd-yyyy HH:mm")
                {
                    csvdata.Columns[0].DefaultCellStyle.Format = "MM/dd/yyyy HH:mm";
                }
                else if (csvdata.Columns[0].DefaultCellStyle.Format == "dd-MM-yyyy HH:mm")
                {
                    csvdata.Columns[0].DefaultCellStyle.Format = "dd/MM/yyyy HH:mm";
                }
            }
            catch { }
               
            combofilter.SelectedIndex = 0;


            foreach (DataGridViewColumn item in csvdata.Columns)
            {
                item.DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
            }
            //Archive_Translate();

            //csvdata.ClearSelection();
            //this.csvdata.Sort(this.csvdata.Columns[LM.Translate("TEST_DATE", MessagesLabelsEN.TEST_DATE)], ListSortDirection.Descending);
           

        }

        /*readonly MainForm mf = new MainForm();
        private void Csvdata_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
            if (e.RowIndex >= 0)
            {
               

                DataGridViewRow row = this.csvdata.Rows[e.RowIndex];
                mf.Testdate.Text = row.Cells["TEST DATE"].Value.ToString();
            }

        }*/
        //public void print()
        //{
           
        //    if (mf.patid.Text == "")
        //    {
        //        mf.UpdateStatus();

        //    }
        //}
        private void Csvdata_CellContentClick(object sender, DataGridViewCellEventArgs e)
        {
          
        }

        private void Txtsearch_KeyDown(object sender, KeyEventArgs e)
        {
        }

        private void BtnFilter_Click(object sender, EventArgs e)
        {
            Logs.LogFile_Entries("Filter button clicked" + " \tBtnFilter_Click in" + this.FindForm().Name, "Info");
            BindingSource bs = new BindingSource();
            bs.DataSource = csvdata.DataSource;
            
            try
            {
                /// MessageBox.Show("DataSource type BEFORE = " + csvdata.DataSource.GetType().ToString());
                if (combofilter.SelectedIndex == 2)
                {

                    //string patient_id = LM.Translate("", MessagesLabelsEN.PATIENT_ID);
                    //string rowFilter = string.Format("[{0}] = '{1}'", patient_id, txtsearch.Text);


                    (csvdata.DataSource as DataTable).DefaultView.RowFilter = string.Format("["+LM.Translate("PATIENT_ID", MessagesLabelsEN.PATIENT_ID)+"] = '{0}'", txtsearch.Text);

                }

                if (combofilter.SelectedIndex == 0)
                {
                    string Test_date = LM.Translate("TEST_DATE", MessagesLabelsEN.TEST_DATE);
                    (csvdata.DataSource as DataTable).DefaultView.RowFilter = string.Format("Convert([{0}], 'System.String') LIKE '%{1}%'", Test_date, txtsearch.Text.Replace(".", "/"), txtsearch.Text.Replace(".", "-"));
                }

                if (combofilter.SelectedIndex == 1)
                {
                  

                    (csvdata.DataSource as DataTable).DefaultView.RowFilter = string.Format("[" + LM.Translate("ACCESSION", MessagesLabelsEN.ACCESSION) + "] = '{0}'", txtsearch.Text);
                    //(csvdata.DataSource as DataTable).DefaultView.RowFilter = string.Format("Field = '{0}'", txtsearch.Text);
                }


                if (combofilter.SelectedIndex == 3)
                {

                    (csvdata.DataSource as DataTable).DefaultView.RowFilter = string.Format("["+LM.Translate("PAT_NAME",MessagesLabelsEN.PATIENT_NAME)+ "] LIKE '{0}%'", txtsearch.Text);
                }

            }

            catch (Exception ex)
            {
                  
            }


        }


       




        private void Csvselect_Click(object sender, EventArgs e)
        {
            //SaveArchive();
            //MainForm mf = new MainForm();
            display_csv_data();
            mf.focus(null, null);
            mf.fructosetxt_Enter(null, null);
            mf.liquifictiontxt_Enter(null, null);
            mf.testerdesg_Enter(null, null);
            mf.testername_Enter(null, null);
            mf.refbydr_Enter(null, null);
            mf.refbydr2_Enter(null, null);
            mf.refbydr3_Enter(null, null);
            refreshfrm();
            mf.SettingsButton.Enabled = true;
            this.Close();


            //mf.ShowDialog();
        }

        private void Average_Report()
        {
            Average_report ar = new Average_report();
            row = this.csvdata.SelectedRows[0];
            row2 = this.csvdata.SelectedRows[1];
            avg_rpt.Clear();
            avg_rpt.Insert(0, row2.Cells[40].Value.ToString());
            avg_rpt.Insert(1, row2.Cells[39].Value.ToString());
            mf.Avg_rowdata1.Clear();
            mf.Avg_rowdata2.Clear();

            float str = 0.0f;
            for (int i = 0; i <= row.Cells.Count-1; i++)
            {
                _row1.Add(row.Cells[i].Value.ToString());
                _row2.Add(row2.Cells[i].Value.ToString());
            }
            for (int j = 12; j <= 27 ; j++)
            {
                mf.Avg_rowdata1.Add(_row2[j]);
                mf.Avg_rowdata2.Add(_row1[j]);
                if (!Regex.IsMatch(_row1[j], @"^[0-9]+(\.[0-9]+)?$") || !Regex.IsMatch(_row2[j], @"^[0-9]+(\.[0-9]+)?$"))
                {

                    if (_row1[j].Contains("<") || _row2[j].Contains("<"))
                    {
                        if (_row1[j] == _row2[j])
                        {
                            _avg.Add(_row1[j]);
                        }
                        else
                        {
                            _avg.Add("N/A");
                        }
                    }
                    else if (_row1[j].Contains(">") || _row2[j].Contains(">"))
                    {
                        if (_row1[j] == _row2[j])
                        {
                            _avg.Add(_row1[j]);
                        }
                        else
                        {
                            _avg.Add("N/A");
                        }
                    }
                    else if(_row1[j]==_row2[j])
                    {
                        if(_row1[j]=="---")
                        {
                            _avg.Add("---");
                        }
                        else
                        {
                            _avg.Add("N/A");
                        }
                    }
                    else
                    {
                        _avg.Add("N/A");
                    }
                }
                else
                {
                    str = ((float.Parse(_row1[j]) + float.Parse(_row2[j])) / 2);
                    _avg.Add(Math.Round(str,2).ToString());
                }
                
            }
            mf.Avg_data.AddRange(_avg);

            Logs.LogFile_Entries("Averaging report generated" + " \tAverage_Report in" + this.FindForm().Name, "Info");

        }

        public void display_csv_data()
        {
            try
            {                
                {
                    
                    mf.Testdate.Text = csvdata.CurrentRow.Cells[0].Value.ToString();
                    mf.patid.Text = csvdata.CurrentRow.Cells[1].Value.ToString();
                    mf.patdob.Text = csvdata.CurrentRow.Cells[3].Value.ToString();
                    mf.absent.Text = csvdata.CurrentRow.Cells[4].Value.ToString();
                    mf.accession.Text = csvdata.CurrentRow.Cells[5].Value.ToString();
                    mf.collecteddt.Text = csvdata.CurrentRow.Cells[6].Value.ToString();
                    mf.rcvdt.Text = csvdata.CurrentRow.Cells[7].Value.ToString();
                    mf.type.Text = csvdata.CurrentRow.Cells[8].Value.ToString();
                    mf.volume.Text = csvdata.CurrentRow.Cells[9].Value.ToString();
                    mf.wbc.Text = csvdata.CurrentRow.Cells[10].Value.ToString();
                    mf.ph.Text = csvdata.CurrentRow.Cells[11].Value.ToString();
                    mf.conc.Text = csvdata.CurrentRow.Cells[12].Value.ToString();
                    mf.concbox.Text = csvdata.CurrentRow.Cells[12].Value.ToString();
                    mf.totmot.Text = csvdata.CurrentRow.Cells[13].Value.ToString();
                    mf.totmotbox.Text = csvdata.CurrentRow.Cells[13].Value.ToString();
                    mf.progmotility.Text = csvdata.CurrentRow.Cells[14].Value.ToString();
                    mf.rapidProg.Text = csvdata.CurrentRow.Cells[14].Value.ToString();
                    mf.nonprgmot.Text = csvdata.CurrentRow.Cells[15].Value.ToString();
                    mf.npmbox.Text = csvdata.CurrentRow.Cells[15].Value.ToString();
                    mf.immotility.Text = csvdata.CurrentRow.Cells[16].Value.ToString();
                    mf.immotbox.Text = csvdata.CurrentRow.Cells[16].Value.ToString();
                    mf.morph.Text = csvdata.CurrentRow.Cells[17].Value.ToString();
                    mf.mnfbox.Text = csvdata.CurrentRow.Cells[17].Value.ToString();
                    mf.msc.Text = csvdata.CurrentRow.Cells[18].Value.ToString();
                    mf.mscbox.Text = csvdata.CurrentRow.Cells[18].Value.ToString();
                    mf.pmsc.Text = csvdata.CurrentRow.Cells[19].Value.ToString();
                    mf.rapid_pmscbox.Text = csvdata.CurrentRow.Cells[19].Value.ToString();
                    mf.fsc.Text = csvdata.CurrentRow.Cells[20].Value.ToString();
                    mf.fscbox.Text = csvdata.CurrentRow.Cells[20].Value.ToString();
                    mf.velocity.Text = csvdata.CurrentRow.Cells[21].Value.ToString();
                    mf.velocitybox.Text = csvdata.CurrentRow.Cells[21].Value.ToString();
                    mf.smi.Text = csvdata.CurrentRow.Cells[22].Value.ToString();
                    mf.smibox.Text = csvdata.CurrentRow.Cells[22].Value.ToString();
                    mf.sperm.Text = csvdata.CurrentRow.Cells[23].Value.ToString();
                    mf.spermbox.Text = csvdata.CurrentRow.Cells[23].Value.ToString();
                    mf.motsperm.Text = csvdata.CurrentRow.Cells[24].Value.ToString();
                    mf.motspermbox.Text = csvdata.CurrentRow.Cells[24].Value.ToString();
                    mf.progsperm.Text = csvdata.CurrentRow.Cells[25].Value.ToString();
                    mf.progspermbox.Text = csvdata.CurrentRow.Cells[25].Value.ToString();
                    mf.funcsperm.Text = csvdata.CurrentRow.Cells[26].Value.ToString();
                    mf.funcspermbox.Text = csvdata.CurrentRow.Cells[26].Value.ToString();
                    mf.label2.Text = csvdata.CurrentRow.Cells[27].Value.ToString();
                    mf.mnfsbox.Text = csvdata.CurrentRow.Cells[27].Value.ToString();

                    mf.awData1 = csvdata.CurrentRow.Cells[38].Value.ToString();
                    mf.awData = csvdata.CurrentRow.Cells[38].Value.ToString();
                    mf.AVGStr1 = csvdata.CurrentRow.Cells[35].Value.ToString();
                    mf.ODStr1 = csvdata.CurrentRow.Cells[37].Value.ToString();
                    mf.ODStr = csvdata.CurrentRow.Cells[37].Value.ToString();
                    mf.patname.Text = csvdata.CurrentRow.Cells[2].Value.ToString();
                    mf.testername.Text = csvdata.CurrentRow.Cells[28].Value.ToString();
                    mf.testerdesg.Text = csvdata.CurrentRow.Cells[29].Value.ToString();
                    mf.fructosetxt.Text = csvdata.CurrentRow.Cells[30].Value.ToString();
                    mf.liquifictiontxt.Text = csvdata.CurrentRow.Cells[31].Value.ToString();
                    mf.refbydr.Text = csvdata.CurrentRow.Cells[32].Value.ToString();
                    mf.refbydr2.Text = csvdata.CurrentRow.Cells[33].Value.ToString();
                    mf.refbydr3.Text = csvdata.CurrentRow.Cells[34].Value.ToString();
                    mf.textBox2.Text = csvdata.CurrentRow.Cells[39].Value.ToString();
                    mf.textBox1.Text = csvdata.CurrentRow.Cells[40].Value.ToString();
                    mf.textBox16.Text = csvdata.CurrentRow.Cells[41].Value.ToString();
                    mf.textBox7.Text = csvdata.CurrentRow.Cells[42].Value.ToString();
                    mf.textBox6.Text = csvdata.CurrentRow.Cells[43].Value.ToString();
                    mf.textBox3.Text = csvdata.CurrentRow.Cells[44].Value.ToString();
                    mf.textBox5.Text = csvdata.CurrentRow.Cells[45].Value.ToString();
                    mf.textBox4.Text = csvdata.CurrentRow.Cells[46].Value.ToString();
                    mf.textBox15.Text = csvdata.CurrentRow.Cells[47].Value.ToString();
                    mf.textBox12.Text = csvdata.CurrentRow.Cells[48].Value.ToString();
                    mf.textBox11.Text = csvdata.CurrentRow.Cells[49].Value.ToString();
                    mf.textBox8.Text = csvdata.CurrentRow.Cells[50].Value.ToString();
                    mf.textBox10.Text = csvdata.CurrentRow.Cells[51].Value.ToString();
                    mf.textBox9.Text = csvdata.CurrentRow.Cells[52].Value.ToString();
                    mf.textBox14.Text = csvdata.CurrentRow.Cells[53].Value.ToString();
                    mf.textBox13.Text = csvdata.CurrentRow.Cells[54].Value.ToString();
                    mf.richTextBox1.Text = csvdata.CurrentRow.Cells[55].Value.ToString();
                    mf.panel1.Enabled = true;
                    mf.checkBox1.Checked = true;
                    if (int.Parse(appSettings[11]) == 1)
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
                    mf.exported = true;
                    mf.patname.Enabled = true;
                    mf.fructosetxt.Enabled = true;
                    mf.liquifictiontxt.Enabled = true;
                    mf.testername.Enabled = true;
                    mf.testerdesg.Enabled = true;
                    mf.refbydr.Enabled = true;
                    mf.refbydr2.Enabled = true;
                    mf.refbydr3.Enabled = true;
                    mf.textBox1.Enabled = true;
                    mf.textBox2.Enabled = true;
                    mf.footer = true;
                    mf.DEL_csv_records = false;
                    mf.archive_data = true;
                    mf.CNTStr1 = csvdata.CurrentRow.Cells[36].Value.ToString();
                    mf.CNTStr = csvdata.CurrentRow.Cells[36].Value.ToString();
                    mf.reportData.Clear();
                    mf.reportData.Insert(0, csvdata.CurrentRow.Cells[40].Value.ToString());
                    mf.reportData.Insert(1, csvdata.CurrentRow.Cells[39].Value.ToString());
                    mf.reportData.Insert(2, csvdata.CurrentRow.Cells[0].Value.ToString());
                    mf.reportData.Insert(3, csvdata.CurrentRow.Cells[1].Value.ToString());
                    mf.reportData.Insert(4, csvdata.CurrentRow.Cells[2].Value.ToString());
                    mf.reportData.Insert(5, csvdata.CurrentRow.Cells[3].Value.ToString());
                    for (int i = 6; i <= csvdata.CurrentRow.Cells.Count - 1; i++)
                    {
                        mf.reportData.Insert(i, csvdata.CurrentRow.Cells[i - 2].Value.ToString());
                    }
                    mf.AVGStr = csvdata.CurrentRow.Cells[35].Value.ToString();

                   



                    if ((Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 0)|| mf.archive_data==true)
                        mf.SaveButton.Enabled = false;
                        mf.devicesn.Text =  csvdata.CurrentRow.Cells[40].Value.ToString();

                    mf.isFresh = false;
                    mf.isFrozen = false;
                    mf.isWashed = false;

                    if (mf.reportData.Contains("FRESH")|| mf.reportData.Contains("FRISCH") || mf.reportData.Contains("FRAIS") || mf.reportData.Contains("FRESCO")||mf.reportData.Contains("新鲜"))
                    {
                        mf.isFresh = true;
                    }
                    if (mf.reportData.Contains("FROZEN") || mf.reportData.Contains("GEFROREN") || mf.reportData.Contains("CONGELE") || mf.reportData.Contains("CONGELATO") || mf.reportData.Contains("冷冻"))
                    {
                        mf.isFrozen = true;
                    }
                    if (mf.reportData.Contains("WASHED") || mf.reportData.Contains("GEWASCHEN") || mf.reportData.Contains("LAVE") || mf.reportData.Contains("LAVATO") || mf.reportData.Contains("洗涤"))
                    {
                        mf.isWashed = true;
                    }
                    mf.translateType1();
                    //n.Hide();
                    //mf.ShowDialog(this);
                    //
                    foreach (Control c in mf.Controls)
                    {
                        if (c is TextBox)
                        {
                            c.Enabled = true;
                        }
                    }
                    Logs.LogFile_Entries("Displayed selected data on main screen" + " \tdisplay_csv_data in" + this.FindForm().Name, "Info");

                }


            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.ToString());
                Logs.LogFile_Entries(ex.ToString() + " \tdisplay_csv_data in" + this.FindForm().Name, "Error");
            }

        }
        public void placeholder()
        {
            for (int i = 0; i <= 3; i++)
            {
                if (combofilter.SelectedIndex == i)
                {
                    txtsearch.Text = LM.Translate("Enter",MessagesLabelsEN.Enter) + " " + combofilter.Text;
                    txtsearch.ForeColor = Color.Silver;

                }
            }
            Logs.LogFile_Entries("Filter by "+ combofilter.Text + "" + " \tplaceholder in" + this.FindForm().Name, "Info");
        }
        private void Combofilter_SelectedIndexChanged(object sender, EventArgs e)
        {
            placeholder();




        }

        private void Txtsearch_TextChanged_1(object sender, EventArgs e)
        {
          
        }

        private void Txtsearch_Leave(object sender, EventArgs e)
        {
            if (txtsearch.Text == "")
            {
                txtsearch.Text = LM.Translate("Enter", MessagesLabelsEN.Enter) + " " + combofilter.Text;
                txtsearch.ForeColor = Color.Gray;
            }
        }

        private void Txtsearch_Enter(object sender, EventArgs e)
        {
            if (txtsearch.Text == LM.Translate("Enter", MessagesLabelsEN.Enter) + " " + combofilter.Text)
            {
                txtsearch.Text = "";
                txtsearch.ForeColor = Color.Black;
            }
        }

        private void Selectdiffcsv_Click_1(object sender, EventArgs e)
        {
            Average_Report();
        }

        private void Txtsearch_KeyPress(object sender, KeyPressEventArgs e)
        {

        }

        private void BtnShowAll_Click(object sender, EventArgs e)
        {
            string path;

            path = appSettings[2];
            string csvpath = path.Substring(0, path.Length - 3) + "csv";
           
            csvdata.DataSource = ReadCsv(csvpath);
            
        }
        Thread th;
        private void Button1_Click(object sender, EventArgs e)
        {
            this.Close();
            Logs.LogFile_Entries("Archive form closed" + " \tButton1_Click in" + this.FindForm().Name, "Info");
            th = new Thread(opennewfrm);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();

        }
        private void opennewfrm(object obj)
        {
           
        }

        private void ArchiveFrm_FormClosed(object sender, FormClosedEventArgs e)
        {

        }

        private void csvdata_MultiSelectChanged(object sender, EventArgs e)
        {

        }


        public void refreshfrm()
        {
            mf.Avg_frclbl.Text = "";
            mf.Avg_liqlbl.Text = "";
            mf.Avg_desg.Text = "";
            mf.Avg_tesname.Text = ""; mf.Avg_patname.Text = "";
            mf.comments.Visible = false;
            mf.Avg_desg.Visible = false;
            mf.Avg_frclbl.Visible = false;
            mf.Avg_liqlbl.Visible = false;
            mf.Avg_tesname.Visible = false;
            mf.Avg_desg.Visible = false;
            mf.Avg_patname.Visible = false;
           
            mf.refbydrlbl.Text = LM.Translate("REF_BY_DR", MessagesLabelsEN.REF_BY_DR).TrimEnd(':');
            mf.SettingsButton.Enabled = true;
            mf.Resultinfo.Text = LM.Translate("RESULT_INFORMATION", MessagesLabelsEN.RESULT_INFO);
            string path = Application.StartupPath + @"\settings.xml";
            appSettings = XmlUtility.ReadSettings(path);


            if (Boolean.Parse(appSettings[7]) == false)
            {
                Logs.LogFile_Entries("Print result strip UI displayed" + " \trefereshfrm in" + this.FindForm().Name, "Info");
                mf.resultstrip.Text = LM.Translate("PRINT_RESULTS_STRIP", MessagesLabelsEN.EXPORT_PDF);
                mf.PrintButton.Image = System.Drawing.Image.FromFile(Application.StartupPath + "\\assets\\report1.png");
                //pdfcheck.Enabled = false;
                mf.ArchiveButton.Enabled = false;
                // mf.pictureBox1.Location = new Point(120, 56);

                //pictureBox1.Image = imageList1.Images[2];
                foreach (Control c in mf.Controls)
                {
                    if (c is System.Windows.Forms.TextBox || c is System.Windows.Forms.Button || c is Label)
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

            else if (Boolean.Parse(appSettings[7]) == true && int.Parse(appSettings[11]) == 1)
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
                mf.patname.Enabled = false;
                mf.fructosetxt.Enabled = false;
                mf.liquifictiontxt.Enabled = false;
                mf.testername.Enabled = false;
                mf.testerdesg.Enabled = false;
                mf.refbydr.Enabled = false;
                mf.refbydr2.Enabled = false;
                mf.refbydr3.Enabled = false;
            }

            if (int.Parse(appSettings[11]) == 1 && Boolean.Parse(appSettings[7]) == true)
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

            mf.translateType1();


        }




        private void csvdata_MouseClick(object sender, MouseEventArgs e)
        {

           
      
            // Handle checkbox state change here
            if (csvdata.SelectedRows.Count ==csvdata.Rows.Count)
            {
                row = this.csvdata.SelectedRows[0];
                csvdata.ClearSelection();
                csvdata.CurrentCell = csvdata.Rows[row.Index].Cells[0];

            }

            if (count >= 0)
            {
                if (ModifierKeys.HasFlag(Keys.Control))
                {
                    try 
                    
                    {
                      
                      
                            row = this.csvdata.SelectedRows[0];
                        
                        try
                        {
                            row2 = this.csvdata.SelectedRows[1];
                        }
                        catch { csvdata.CurrentCell = csvdata.Rows[row.Index].Cells[0]; }
                        bool check_sample = false;

                        if (row != null && row2 != null)
                        {
                            for (int i = 4; i <= 11; i++)
                            {
                                 if (row.Cells[i].Value.ToString() == row2.Cells[i].Value.ToString())
                                {
                                    check_sample = true;
                                }
                                else
                                {
                                    check_sample = false;
                                    break;
                                }
                            }
                            if (row.Cells[30].Value.ToString() == row2.Cells[30].Value.ToString() && row.Cells[31].Value.ToString() == row2.Cells[31].Value.ToString()&& check_sample == true )
                            {
                                check_sample = true;
                            }
                            else
                            {

                                check_sample = false;
                            }
                            if(check_sample == true && (row.Cells[2].Value.ToString().ToUpper() == row2.Cells[2].Value.ToString().ToUpper()))
                            {
                                check_sample = true;
                            }
                            else if (check_sample == true && (row.Cells[2].Value.ToString() == "---" || row2.Cells[2].Value.ToString() == "---"))
                            {
                                
                                
                                    check_sample = true;
                                
                                
                            }
                            else
                            {
                                check_sample = false;
                            }
                            string Test1_Time = row.Cells[0].Value.ToString().Substring(10, 5);
                            string Test2_Time = row2.Cells[0].Value.ToString().Substring(10, 5);
                            TimeSpan duration = DateTime.Parse(Test1_Time).Subtract(DateTime.Parse(Test2_Time));
                            //MessageBox.Show(duration.Duration().ToString());
                            if (row.Cells[1].Value.ToString() == row2.Cells[1].Value.ToString() && csvdata.SelectedRows.Count <= 2 && duration.Duration() <= TimeSpan.Parse("00:10") && check_sample == true && row.Cells[0].Value.ToString() != row2.Cells[0].Value.ToString())
                            {
                               
                                row.Selected = true;
                                generate.Enabled = true;

                                csvselect.Enabled = false;
                               
                            }
                            else
                            {
                             
                                    MessageBox.Show(LM.Translate("Arc_same_test", MessagesLabelsEN.Arc_same_test), LM.Translate("Arc_err", MessagesLabelsEN.Arc_same_test),MessageBoxButtons.OK,MessageBoxIcon.Information);
                               

                                if (csvdata.SelectedRows.Count <= 2)
                                {
                                    row.Selected = false;
                                    csvdata.CurrentCell = csvdata.Rows[row2.Index].Cells[0];
                                }

                                try
                                {

                                    row3 = this.csvdata.SelectedRows[2];
                                    temprow = this.csvdata.SelectedRows[1];

                                    row.Selected = false;
                                    csvdata.CurrentCell = csvdata.Rows[row3.Index].Cells[0];
                                    row2.Selected = true;
                                    row3.Selected = true;

                                }
                                catch
                                {
                                    row.Selected = false;
                                }
                            }
                        }
                    }
                    catch(Exception ex)
                    {
                        //MessageBox.Show(ex.ToString());
                    }
                }
                if(csvdata.SelectedRows.Count == 1)
                {
                    csvselect.Enabled = true;
                    generate.Enabled = false;

                }

            }

        }

        private void csvdata_RowEnter(object sender, DataGridViewCellEventArgs e)
        {
            
            try
            {

                string data = string.Empty;
                foreach (DataGridViewRow row in csvdata.Rows)
                {
                    if (row.Cells[SelectColumnIndex].Value != null &&
                           Convert.ToBoolean(row.Cells[SelectColumnIndex].Value) == true)
                    {
                        foreach (DataGridViewCell cell in row.Cells)
                        {
                            if (cell.OwningColumn.Index != SelectColumnIndex)
                            {
                                data += (cell.Value + " "); // do some thing
                            }
                        }
                        data += "\n";
                    }
                }
                //MessageBox.Show(data, "Data");
            }
            catch { }







        }

        private void csvdata_CellContentClick_1(object sender, DataGridViewCellEventArgs e)
        {
            
        }

        private void csvdata_CurrentCellDirtyStateChanged(object sender, EventArgs e)
        {
           
        }

        private void csvdata_RowDirtyStateNeeded(object sender, QuestionEventArgs e)
        {
            
        }

        private void csvdata_DragDrop(object sender, DragEventArgs e)
        {
            row = this.csvdata.SelectedRows[0];
            row.Selected = false;
        }

        private void csvdata_DoubleClick_1(object sender, MouseEventArgs e)
        {

           


        }

        private void csvdata_KeyUp(object sender, KeyEventArgs e)
        {
            
        }
         
        private void csvdata_CellClick(object sender, DataGridViewCellEventArgs e)
        {
           
        }

        private void csvdata_DragDrop(object sender, EventArgs e)
        {
            try
            {
                row = this.csvdata.SelectedRows[0];
                row.Selected = false;
            }
            catch { }
        }

        private void csvdata_DragEnter(object sender, DragEventArgs e)
        {

        }

        private void Csvdata_CellValueChanged(object sender, DataGridViewCellEventArgs e)
        {
            //if (e.ColumnIndex == chk.Index && e.RowIndex != -1)
            //{
            //    // Handle checkbox state change here
            //    chk.TrueValue = true;
            //}
        }

        private void Csvdata_MouseUp(object sender, MouseEventArgs e)
        {
            if(csvdata.SelectedRows.Count>2)
            {
                row = this.csvdata.SelectedRows[0];
                csvdata.ClearSelection();
                csvdata.CurrentCell = csvdata.Rows[row.Index].Cells[0];
            }
            

           


        }

        private void Csvdata_KeyDown(object sender, KeyEventArgs e)
        {
           
            switch (e.KeyData & Keys.KeyCode)
            {
                case Keys.Up:
                case Keys.ShiftKey:
                case Keys.Right:
                case Keys.Down:
                case Keys.Left:
                
                    e.Handled = true;
                    e.SuppressKeyPress = true;
                    break;
            }
            //if (count >= 0)
            //{
            //    if (ModifierKeys.HasFlag(Keys.Shift))
            //    {
            //        try
            //        {
            //            row = this.csvdata.SelectedRows[0];
            //            row2 = this.csvdata.SelectedRows[1];
            //            row.Selected = false;
            //            row2.Selected = false;
            //        }
            //        catch
            //        {
            //            //row.Selected = false;
            //        }
            //    }

            //}
        }

        private void csvdata_MouseDoubleClick(object sender, MouseEventArgs e)
        {
            if (csvdata.SelectedRows.Count <= 2)
            {
                try
                {
                    row2 = this.csvdata.SelectedRows[1];
                    row = this.csvdata.SelectedRows[0];
                    row.Selected = false;

                    csvdata.CurrentCell = csvdata.Rows[row2.Index].Cells[0];
                }
                catch
                {
                    //csvdata.CurrentCell = csvdata.Rows[row2.Index].Cells[0];
                }
            }
            if (csvdata.SelectedRows.Count > 2)
            {
                try
                {

                    row3 = this.csvdata.SelectedRows[2];
                    temprow = this.csvdata.SelectedRows[1];

                    row.Selected = false;
                    csvdata.CurrentCell = csvdata.Rows[row3.Index].Cells[0];
                    row2.Selected = true;
                    row3.Selected = true;

                }
                catch
                {
                    row.Selected = false;
                }
            }
            if (ModifierKeys.HasFlag(Keys.Control))
            {
                try

                {
                    row = null; row2 = null; row3 = null;

                    row = this.csvdata.SelectedRows[0];

                    try
                    {
                        row2 = this.csvdata.SelectedRows[1];
                    }
                    catch { csvdata.CurrentCell = csvdata.Rows[row.Index].Cells[0]; }
                    bool check_sample = false;
                    if (row != null && row2 != null)
                    {
                        for (int i = 4; i <= 11; i++)
                        {
                            if (row.Cells[i].Value.ToString() == row2.Cells[i].Value.ToString())
                            {
                                check_sample = true;
                            }
                            else
                            {
                                check_sample = false;
                                break;
                            }
                        }
                        if (row.Cells[30].Value.ToString() == row2.Cells[30].Value.ToString() && row.Cells[31].Value.ToString() == row2.Cells[31].Value.ToString())
                        {
                            check_sample = true;
                        }
                        else
                        {
                            check_sample = false;
                        }
                        string Test1_Time = row.Cells[0].Value.ToString().Substring(10, 5);
                        string Test2_Time = row2.Cells[0].Value.ToString().Substring(10, 5);
                        TimeSpan duration = DateTime.Parse(Test1_Time).Subtract(DateTime.Parse(Test2_Time));
                        //MessageBox.Show(duration.Duration().ToString());
                        if (row.Cells[1].Value.ToString() == row2.Cells[1].Value.ToString() && csvdata.SelectedRows.Count <= 2 && duration.Duration() <= TimeSpan.Parse("00:10") && check_sample == true && row.Cells[0].Value.ToString() != row2.Cells[0].Value.ToString())
                        {
                            mf.Avg_data.Clear();
                            row.Selected = true;
                            generate.Enabled = true;

                            csvselect.Enabled = false;
                            for (int i = 0; i <= 11; i++)
                            {
                                mf.Avg_data.Add(row2.Cells[i].Value.ToString());
                            }
                        }
                        else
                        {

                            if (csvdata.SelectedRows.Count <= 2)
                            {
                                row.Selected = false;
                                csvdata.CurrentCell = csvdata.Rows[row2.Index].Cells[0];
                            }

                            try
                            {

                                row3 = this.csvdata.SelectedRows[2];
                                temprow = this.csvdata.SelectedRows[1];

                                row.Selected = false;
                                csvdata.CurrentCell = csvdata.Rows[row3.Index].Cells[0];
                                row2.Selected = true;
                                row3.Selected = true;

                            }
                            catch
                            {
                                row.Selected = false;
                            }
                        }
                    }
                }
                catch (Exception ex)
                {
                    //MessageBox.Show(ex.ToString());
                }
            }



        }

        private void ArchiveFrm_MouseMove(object sender, MouseEventArgs e)
        {
            
        }

        private void generate_MouseMove(object sender, MouseEventArgs e)
        {
            Control control = GetChildAtPoint(e.Location);
            if (control != null)
            {
                if (!control.Enabled )
                {
                    string toolTipString = _toolTip.GetToolTip(control);
                    // trigger the tooltip with no delay and some basic positioning just to give you an idea
                    _toolTip.Show(toolTipString, control, control.Width / 2, control.Height / 2);
                    _currentToolTipControl = control;
                }
            }
            else
            {
                if (_currentToolTipControl != null) _toolTip.Hide(_currentToolTipControl);
                _currentToolTipControl = null;
            }

            //Control ctrl = this.GetChildAtPoint(e.Location);

            //if (ctrl != null)
            //{
            //    if (ctrl == this.generate && !IsShown)
            //    {
            //        string tipstring = this.toolTip1.GetToolTip(this.generate);
            //        this.toolTip1.Show(tipstring, this.generate, this.generate.Width / 2,
            //                                                    this.generate.Height / 2);
            //        IsShown = true;
            //    }
            //}
            //else
            //{
            //    this.toolTip1.Hide(this.generate);
            //    IsShown = false;
            //}

        }

        private void generate_MouseHover(object sender, EventArgs e)
        {
            //toolTip1.Show("aaaaa", generate, generate.Width / 2, generate.Height / 2, 200);
        }

        private void label1_MouseMove(object sender, MouseEventArgs e)
        {
            Control control = GetChildAtPoint(e.Location);
            if (control != null)
            {
                if (!control.Enabled)
                {
                    string toolTipString = _toolTip.GetToolTip(control);
                    // trigger the tooltip with no delay and some basic positioning just to give you an idea
                    _toolTip.Show(toolTipString, control, control.Width / 2, control.Height / 2);
                    _currentToolTipControl = control;
                }
            }
            else
            {
                if (_currentToolTipControl != null) _toolTip.Hide(_currentToolTipControl);
                _currentToolTipControl = null;
            }
        }

        private void textBox1_MouseMove(object sender, MouseEventArgs e)
        {

        }

        private void generate_Click(object sender, EventArgs e)
        {
            //var culture = CultureInfo.InvariantCulture.DateTimeFormat;
            //DateTimeStyles styles = DateTimeStyles.AssumeUniversal;
            //string Test1_Time = row.Cells[0].Value.ToString();
            //string Test2_Time = row2.Cells[0].Value.ToString();
            //try
            //{
            //    DateTime duration = DateTime.Parse(Test1_Time, culture, styles);
            //    DateTime duration1 = DateTime.Parse(Test2_Time, culture, styles);

            //    if (duration < duration1)
            //    {

            //        mf.Avg_data[0] = Test2_Time;
            //    }
            //    else
            //    {

            //        mf.Avg_data[0] = Test2_Time.ToString();
            //    }
            //}
            //catch
            //{
            //    mf.Avg_data[0] = Test2_Time.ToString();
            //}
            //display_AVG();
            //mf.Avg_data.Add("");
            //this.Close();
            //mf.SettingsButton.Enabled = false;
            
        }

        private void button1_Click_1(object sender, EventArgs e)
        {
          
            //Defaults
           
            //fRUCTOSE AND lIQUIFACTION


            this.Close();

        }

        private void ArchiveFrm_Leave(object sender, EventArgs e)
        {

        }

        private void Average_Click(object sender, EventArgs e)
        {
       
            mf.cleartext();
            mf.Avg_data.Clear();
            for (int i = 0; i <= 11; i++)
            {
                mf.Avg_data.Add(row2.Cells[i].Value.ToString());
            }
            var culture = CultureInfo.InvariantCulture.DateTimeFormat;
            DateTimeStyles styles = DateTimeStyles.AssumeUniversal;
            string Test1_Time = row.Cells[0].Value.ToString();
            string Test2_Time = row2.Cells[0].Value.ToString();
            try
            {
                DateTime duration = DateTime.Parse(Test1_Time, culture, styles);
                DateTime duration1 = DateTime.Parse(Test2_Time, culture, styles);

                if (duration < duration1)
                {

                    mf.Avg_data[0] = Test2_Time;
                }
                else
                {

                    mf.Avg_data[0] = Test2_Time.ToString();
                }
            }
            catch
            {
                mf.Avg_data[0] = Test2_Time.ToString();
            }
            display_AVG();
            mf.Avg_data.Add("");
            mf.Avg_data.Add(row.Cells[40].Value.ToString());
            this.Close();
            mf.SettingsButton.Enabled = false;
              row = null; row2 = null; row3 = null;
        }

        private void Confirm_Click(object sender, EventArgs e)
        {
            display_csv_data();
            mf.focus(null, null);
            mf.fructosetxt_Enter(null, null);
            mf.liquifictiontxt_Enter(null, null);
            mf.testerdesg_Enter(null, null);
            mf.testername_Enter(null, null);
            mf.refbydr_Enter(null, null);
            mf.refbydr2_Enter(null, null);
            mf.refbydr3_Enter(null, null);
            refreshfrm();
            mf.SettingsButton.Enabled = true;
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
            this.Close();

        }

        private void toolStripButton1_Click(object sender, EventArgs e)
        {
            this.Close();
            Logs.LogFile_Entries("Archive form closed" + " \tButton1_Click in" + this.FindForm().Name, "Info");
            th = new Thread(opennewfrm);
            th.SetApartmentState(ApartmentState.STA);
            th.Start();
        }

        private void generate_MouseHover_1(object sender, EventArgs e)
        {
            
        }

        private void generate_MouseMove_1(object sender, MouseEventArgs e)
        {
      
            //if (csvdata.Enabled == false && generate.Enabled == true)
            //{
            //    generate.ToolTipText = string.Format(LM.Translate("Average", MessagesLabelsEN.SAVE), Environment.NewLine);
            //}
            //if(csvdata.Enabled==true && generate.Enabled==false)
            //{
            //    generate.ToolTipText = string.Format(LM.Translate("Archive_tooltip", MessagesLabelsEN.Average), Environment.NewLine);
            //}
        }

        private void csvdata_DoubleClick_1(object sender, EventArgs e)
        {

        }

        private void csvdata_Click_1(object sender, EventArgs e)
        {

        }

        private void Csvdata_KeyPress(object sender, KeyPressEventArgs e)
        {
           

        }

        public void display_AVG()
        {

            //if on basic

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




            mf.refbydr.Visible = false;
            mf.refbydr2.Visible = false;
            mf.liquifictiontxt.Visible = false;
            mf.fructosetxt.Visible = false;
            mf.refbydr3.Visible = false;
            mf.refbtn.Visible = false;
            mf.comments.Visible = true;
            mf.fructoselbl.Visible = true;
            mf.liquilbl.Visible = true;
            mf.Testerinfo.Visible = true;
            mf.testerdesg.Visible = true;
            mf.label40.Visible = true;
            mf.patnamelbl.Visible = true;
            mf.refbydrlbl.Visible = true;
            mf.testernamelbl.Visible = true;
            mf.tesdesg.Visible = true;
            mf.refbydrlbl.Text = LM.Translate("COMMENTS", "COMMENTS");
            mf.Resultinfo.Text = LM.Translate("Avg_Result_info", "Average Result Information");
            mf.rptname.Text = LM.Translate("Avg_report", "Average Report");
            mf.pictureBox1.Image = mf.imageList1.Images[4];
            mf.testerdesg.Visible = false;
            mf.testername.Visible = false;
            mf.Avg_tesname.Visible = true;
            mf.Avg_desg.Visible = true;

            Average_Report();
            
            //Display Average values
            mf.conc.Text = _avg[0];
            mf.totmot.Text = _avg[1];
            mf.progmotility.Text = _avg[2];
            mf.nonprgmot.Text = _avg[3];
            mf.immotility.Text = _avg[4];
            mf.morph.Text = _avg[5];
            mf.msc.Text = _avg[6];
            mf.pmsc.Text = _avg[7];
            mf.fsc.Text = _avg[8];
            mf.velocity.Text = _avg[9];
            mf.smi.Text = _avg[10];
            mf.sperm.Text = _avg[11];
            mf.motsperm.Text = _avg[12];
            mf.progsperm.Text = _avg[13];
            mf.funcsperm.Text = _avg[14];
            mf.label2.Text = _avg[15];
            mf.Testdate.Text = mf.Avg_data[0];
            mf.patid.Text = _row1[1];
            mf.patdob.Text = _row1[3];
            mf.absent.Text = _row1[4];
            mf.accession.Text = _row1[5];
            mf.collecteddt.Text = _row1[6];
            mf.rcvdt.Text = _row1[7];
            mf.type.Text = _row1[8];
            mf.volume.Text = _row1[9];
            mf.wbc.Text = _row1[10];
            mf.devicesn.Text = _row1[40];
            mf.ph.Text = _row1[11];
      
            //Displaying sample and additional info 
            //patient name
            if ((_row1[2] != "---" && _row2[2] == "---"))
            {
                mf.patname.Visible = false;
                mf.Avg_patname.Visible = true;
                mf.Avg_patname.Text = _row1[2];

            }
            else if ((_row1[2] == "---" && _row2[2] != "---"))
            {
                mf.patname.Visible = false;
                mf.Avg_patname.Visible = true;
                mf.Avg_patname.Text = _row2[2];
                mf.Avg_data[2] = _row2[2];
            }
            else if (_row1[2].ToUpper()==_row2[2].ToUpper()&& _row1[2] != "---" && _row2[2] != "---")
            {
                mf.patname.Visible = false;
                mf.Avg_patname.Visible = true;
                mf.Avg_patname.Text = _row2[2].ToUpper();
                mf.Avg_data[2] = _row2[2].ToUpper();
            }
            else
            {
                mf.patname.Visible = true;
                mf.patname.Enabled = true;
                mf.Avg_patname.Visible = false;
            }
            //Tester name 
            if (_row1[28] != "---" && _row2[28] != "---")
            {

                if (_row1[28] == _row2[28])
                {
                    mf.Avg_tesname.Text = _row1[28];
                    mf.Avg_data.Add(_row1[28]);
                }
                else
                {
                    mf.Avg_tesname.Text = _row2[28] + " | " + _row1[28];
                    mf.Avg_data.Add(_row2[28] + " | " + _row1[28]);
                }

            }
            else if ((_row1[28] == "---" && _row2[28] != "---"))
            {
                mf.Avg_tesname.Text = _row2[28];
                mf.Avg_data.Add(_row2[28]);
            }
            else if ((_row1[28] != "---" && _row2[28] == "---"))
            {
                mf.Avg_tesname.Text = _row1[28];
                mf.Avg_data.Add(_row1[28]);
            }
            else
            {
                mf.Avg_tesname.Text = "";
                mf.Avg_data.Add(" ");
            }
            //designation
            if (_row1[29] != "---" && _row2[29] != "---")
            {

                if (_row1[29] == _row2[29])
                {
                    mf.Avg_desg.Text = _row1[29];
                    mf.Avg_data.Add(_row1[29]);
                }
                else
                {
                    mf.Avg_desg.Text = _row2[29] + " | " + _row1[29];
                    mf.Avg_data.Add(_row2[29] + " | " + _row1[29]);
                }

            }
            else if ((_row1[29] == "---" && _row2[29] != "---"))
            {
                mf.Avg_desg.Text = _row2[29];
                mf.Avg_data.Add(_row2[29]);
            }
            else if ((_row1[29] != "---" && _row2[29] == "---"))
            {
                mf.Avg_desg.Text = _row1[29];
                mf.Avg_data.Add(_row1[29]);
            }
            else
            {
                mf.Avg_desg.Text = "";
                mf.Avg_data.Add(" ");
            }
            //fructose
            if (_row1[30] != "---")
            {
                mf.Avg_frclbl.Text = _row1[30];
                mf.Avg_data.Add(_row1[30]);

            }
            else { mf.Avg_frclbl.Text = "N/A"; mf.Avg_data.Add("N/A"); }
            if (_row1[31] == "---")
            {
                mf.Avg_liqlbl.Text = "N/A";
                mf.Avg_data.Add("N/A");
            }
            else { mf.Avg_liqlbl.Text = _row1[31]; mf.Avg_data.Add(_row1[31]); }

        }

        private void csvdata_DoubleClick(object sender, EventArgs e)
        {

            //if (count >= 1)
            //{
            //    if (ModifierKeys.HasFlag(Keys.Control))
            //    {
            //        row = this.csvdata.SelectedRows[0];
            //        row2 = this.csvdata.SelectedRows[1];
            //        if (row.Cells[0].Value.ToString() == row2.Cells[0].Value.ToString() && csvdata.SelectedRows.Count <= 2)
            //        {
            //            row.Selected = true;
            //        }
            //        else
            //        {
            //            row.Selected = false;
            //        }
            //    }

            //}
        }
    }
}
