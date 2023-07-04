using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using QcGoldArchive;
using System.IO;
using System.Diagnostics;


namespace pdfReportTest
{
    public partial class Form1 : Form
    {

        PdfReportGenerator pdfReportGen;

        DataTable DeviceInfoTable, PatientInfoTable, SampleInfoTable, ParamInfoTable;

        public Form1()
        {
            InitializeComponent();
            DeviceInfoTable = new DataTable();
            PatientInfoTable = new DataTable();
            SampleInfoTable = new DataTable();
            ParamInfoTable = new DataTable();

            //CreatePatientInformation(ParamInfoTable);
            //CreateSampleInformationTable(SampleInfoTable);
            //CreateSemAnalTestResultsTable(ParamInfoTable);
        }

        private void addheader(DataTable DT)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("value", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = "PATIENT ID:";
            row0["value"] = "23452";
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = "BIRTH DATE";
            row1["value"] = "02/27/1988";
            DT.Rows.Add(row1);

        }


        private void CreatePatientInformation(DataTable DT)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("value", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = "PATIENT ID:";
            row0["value"] = "23452";
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = "BIRTH DATE";
            row1["value"] = "02/27/1988";
            DT.Rows.Add(row1);

        }

        private void CreateDeviceInformation(DataTable DT)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("value", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = "DEVICE SN#:";
            row0["value"] = "23452";
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = "SW VERSION:";
            row1["value"] = "01.02.33";
            DT.Rows.Add(row1);

        }

        private void CreateSystemInformation(DataTable DT)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("value", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = "SERIAL NUMBER:";
            row0["value"] = "2343";
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = "SOFTWARE VERSION:";
            row1["value"] = "4.23.44";
            DT.Rows.Add(row1);

        }

        private void CreateSampleInformationTable(DataTable DT)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("value", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = "TEST DATE:";
            row0["value"] = "03/12/18";
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = "ABSITIENCE (days):";
            row1["value"] = "3";
            DT.Rows.Add(row1);

            DataRow row2 = DT.NewRow();
            row2["parameter"] = "SAMPLE ID (Accession):";
            row2["value"] = "12345679";
            DT.Rows.Add(row2);

            DataRow row3 = DT.NewRow();
            row3["parameter"] = "COLLECTED (day/time):";
            row3["value"] = "08/09/18 14:08";
            DT.Rows.Add(row3);

            DataRow row4 = DT.NewRow();
            row4["parameter"] = "RECEIVED (day/time):";
            row4["value"] = "08/09/18 18:02";
            DT.Rows.Add(row4);

            DataRow row5 = DT.NewRow();
            row5["parameter"] = "TYPE:";
            row5["value"] = "FRESH";
            DT.Rows.Add(row5);

            DataRow row6 = DT.NewRow();
            row6["parameter"] = "VOLUME (ml):";
            row6["value"] = "4";
            DT.Rows.Add(row6);

            DataRow row7 = DT.NewRow();
            row7["parameter"] = "WBC CONC. (M/ml):";
            row7["value"] = "<= 1";
            DT.Rows.Add(row7);

            DataRow row8 = DT.NewRow();
            row8["parameter"] = "PH";
            row8["value"] = "7";
            DT.Rows.Add(row8);
        }

        private void CreateSemAnalTestResultsTable(DataTable DT)
        {
            DT.Columns.Add("parameter", typeof(string));
            DT.Columns.Add("result", typeof(string));

            DataRow row0 = DT.NewRow();
            row0["parameter"] = "CONC. (M/ml):";
            row0["result"] = "41.1";
            DT.Rows.Add(row0);

            DataRow row1 = DT.NewRow();
            row1["parameter"] = "MOTILITY (%):";
            row1["result"] = "25.1";
            DT.Rows.Add(row1);

            DataRow row2 = DT.NewRow();
            row2["parameter"] = "RAPID PROGRESSIVE a (%):";
            row2["result"] = "0";
            DT.Rows.Add(row2);

            DataRow row3 = DT.NewRow();
            row3["parameter"] = "SLOW PROGRESSIVE a (%):";
            row3["result"] = "1";
            DT.Rows.Add(row3);

            DataRow row4 = DT.NewRow();
            row4["parameter"] = "NON-PROGRESSIVE c (%):";
            row4["result"] = "2";
            DT.Rows.Add(row4);

            DataRow row5 = DT.NewRow();
            row5["parameter"] = "IMMOTILE d (%):";
            row5["result"] = "2";
            DT.Rows.Add(row5);

            DataRow row6 = DT.NewRow();
            row6["parameter"] = "N.MORPH (%):";
            row6["result"] = "1";
            DT.Rows.Add(row6);

            DataRow row7 = DT.NewRow();
            row7["parameter"] = "MSC (M/ml):";
            row7["result"] = "2";
            DT.Rows.Add(row7);

            DataRow row8 = DT.NewRow();
            row8["parameter"] = "PMSC <a> (M/ml):";
            row8["result"] = "0";
            DT.Rows.Add(row8);

            DataRow row9 = DT.NewRow();
            row9["parameter"] = "PMSC <b> (M/ml):";
            row9["result"] = "0";
            DT.Rows.Add(row9);

            DataRow row10 = DT.NewRow();
            row10["parameter"] = "FSC (M/ml):";
            row10["result"] = "0";
            DT.Rows.Add(row10);

            DataRow row11 = DT.NewRow();
            row11["parameter"] = "VELOCITY (mic/sec):";
            row11["result"] = "<1";
            DT.Rows.Add(row11);

            DataRow row12 = DT.NewRow();
            row12["parameter"] = "SMI:";
            row12["result"] = "2";
            DT.Rows.Add(row12);

            DataRow row13 = DT.NewRow();
            row13["parameter"] = "SPERM # (M/ejac):";
            row13["result"] = "350";
            DT.Rows.Add(row13);

            DataRow row14 = DT.NewRow();
            row14["parameter"] = "MOTILE SPERM (M/ejac):";
            row14["result"] = "85.7";
            DT.Rows.Add(row14);

            DataRow row15 = DT.NewRow();
            row15["parameter"] = "PROG. MOTILE SPERM (M/ejac):";
            row15["result"] = "4";
            DT.Rows.Add(row15);

            DataRow row16 = DT.NewRow();
            row16["parameter"] = "FUNCTIONAL SPERM (M/ejac):";
            row16["result"] = "2";
            DT.Rows.Add(row16);

        }

        private void btn_create_report_Click(object sender, EventArgs e)
        {
            pdfReportGen = new PdfReportGenerator();

            DataTable Headertable = new DataTable();
            addheader(Headertable);

            DataTable PatientInformationTable = new DataTable();
            CreatePatientInformation(PatientInformationTable);

            DataTable DeviceInformationTable = new DataTable();
            CreateDeviceInformation(DeviceInformationTable);

            DataTable SampleInformationTable = new DataTable();
            CreateSampleInformationTable(SampleInformationTable);


            DataTable SemAnalParametersTable = new DataTable();
            CreateSemAnalTestResultsTable(SemAnalParametersTable);
            DataRow SemAnalParametersDR = SemAnalParametersTable.NewRow();
            SemAnalParametersDR["parameter"] = "PARAMETER";
            SemAnalParametersDR["result"] = "RESULT";



            // Generate the report
            bool pdfCreated;
            pdfCreated = pdfReportGen.GenerateSEMENANALYSISReport(Environment.CurrentDirectory + "/Semen Analysis Report.pdf", "Semen Analysis Report.pdf", "SEMEN ANALYSIS TEST REPORT", "REPORT CONTINUED | PATIENT ID: 2345678789567556 | TEST DATE / TIME: 07/31/2013 08:30", "DEVICE INFORMATION", DeviceInformationTable, "PATIENT INFORMATION", PatientInformationTable, "SAMPLE INFORMATION", SampleInformationTable, SemAnalParametersDR, SemAnalParametersTable);
            if (pdfCreated)
            {
                lblSemenAnalElapsed.Text = "success";
                Process.Start(Environment.CurrentDirectory + "/Semen Analysis Report.pdf");
            }

        }
    }
}
