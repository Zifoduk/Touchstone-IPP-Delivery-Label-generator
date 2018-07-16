using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Configuration;
using System.Windows.Forms;
using Excel = Microsoft.Office.Interop.Excel;
using MVSTA = Microsoft.VisualStudio.Tools.Applications;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using iTextSharp.text;

namespace TS_Post_Database_Inserter
{
    public partial class Start : Form
    {
        Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        public string MasterExcel;
        public string MainExcel = "";
        public string OpenPDF;

        PdfReader reader;

        public Start()
        {
            InitializeComponent();
            //Testing
            //Config.AppSettings.Settings["MasterExcel"].Value = "";
            //Config.AppSettings.Settings["MainExcel"].Value = "";
            //Config.AppSettings.Settings["OpenPDF"].Value = "";

            //File Settings location
            MasterExcel = Config.AppSettings.Settings["MasterExcel"].Value;
            MainExcel = Config.AppSettings.Settings["MainExcel"].Value;
            OpenPDF = Config.AppSettings.Settings["OpenPDF"].Value;
            
            if (MasterExcel == "")
            {
                ElL.Text = "Select a Master Excel document!!";
                ElL.ForeColor = Color.DarkRed;
            }
            else
            {
                ElL.Text = MasterExcel;
                ElL.ForeColor = Color.Black;
            }

            if (MainExcel == "")
            {
                PHExcelL.Text = "Select a Main Excel document!!";
                PHExcelL.ForeColor = Color.DarkRed;
            }
            else
            {
                PHExcelL.Text = MasterExcel;
                PHExcelL.ForeColor = Color.Black;
            }

            if(OpenPDF == "")
            {
                LpdfL.Text = "Select a Label PDF file!!";
                LpdfL.ForeColor = Color.DarkRed;
                PDFL.Text = "";
            }
            else
            {
                LpdfL.Text = OpenPDF;
                LpdfL.ForeColor = Color.Black;
                reader = new PdfReader(OpenPDF);
                PDFL.Text = ("Number of Labels found: " + reader.NumberOfPages);
            }
        }

        private void OpenEF_Click(object sender, EventArgs e)
        {
            OpenFileDialog File = openFileDialog1;
            File.Filter = "Excel Files(*.xlsx)|*.xlsx";
            if (File.ShowDialog() == DialogResult.OK)
            {
                if (MainExcel != "")
                {
                    MECheck mec = new MECheck(this, File.FileName, 1);
                    mec.ShowDialog();
                    Config.AppSettings.Settings["MainExcel"].Value = MainExcel;
                    Config.Save(ConfigurationSaveMode.Full);
                }
                else
                {
                    MainExcel = File.FileName;
                    PHExcelL.Text = File.FileName;
                    PHExcelL.ForeColor = Color.Black;
                    MEC mec = new MEC(1);
                    Config.AppSettings.Settings["MainExcel"].Value = MainExcel;
                    mec.ShowDialog();
                }
            }
        }

        private void OpenMEF_Click(object sender, EventArgs e)
        {
            OpenFileDialog File = openFileDialog1;
            File.Filter = "Excel Files(*.xlsx)|*.xlsx";
            if (File.ShowDialog() == DialogResult.OK)
            { 
                if (MasterExcel != "")
                {
                    MECheck mec = new MECheck(this, File.FileName, 0);
                    mec.ShowDialog();
                    Config.AppSettings.Settings["MasterExcel"].Value = MasterExcel;
                    Config.Save(ConfigurationSaveMode.Full);
                }
                else
                {
                    MasterExcel = File.FileName;
                    ElL.Text = File.FileName;
                    ElL.ForeColor = Color.Black;
                    MEC mec = new MEC(0);
                    Config.AppSettings.Settings["MasterExcel"].Value = MasterExcel;
                    mec.ShowDialog();
                }
            }
        }

        private void OpenPDF_Click(object sender, EventArgs e)
        {
            OpenFileDialog File = openFileDialog1;
            File.Filter = "PDF Files(*.pdf)|*.pdf";
            if (File.ShowDialog() == DialogResult.OK)
            {
                OpenPDF = File.FileName;
                LpdfL.Text = File.FileName;
                LpdfL.ForeColor = Color.Black;
                Config.AppSettings.Settings["OpenPDF"].Value = OpenPDF;
                Config.Save(ConfigurationSaveMode.Full);
                reader = new PdfReader(File.FileName);
                PDFL.Text = ("Number of Labels found: " + reader.NumberOfPages);
            }
            // Launch PDF Number found
        }

        private void Launch_Click(object sender, EventArgs e)
        {
            //Excel.Application excel = new Excel.Application();
            //Excel.Workbook sheet = excel.Workbooks.Open(Excel_Path);
            //Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            CNTL c = new CNTL();
            if (MasterExcel == "" || OpenPDF == "" || MainExcel == "")
                c.ShowDialog();
            else
                LaunchMethod();

            if (MasterExcel == "")
            {
                ElL.Text = "Select a Master Excel document!!";
                ElL.ForeColor = Color.DarkRed;
            }
            if (MainExcel == "")
            {
                PHExcelL.Text = "Select a Main Excel document!!";
                PHExcelL.ForeColor = Color.DarkRed;
            }
            if (OpenPDF == "")
            {
                LpdfL.Text = "Select a Label PDF!!";
                LpdfL.ForeColor = Color.DarkRed;
            }
        }

        internal void LaunchMethod()
        {
            CustInfo Cust = new CustInfo(this);
            Cust.ShowDialog();            
        }

        private void Start_FormClosing(object sender, FormClosingEventArgs e)
        {
            Config.Save(ConfigurationSaveMode.Full);
            
        }
    }
}
