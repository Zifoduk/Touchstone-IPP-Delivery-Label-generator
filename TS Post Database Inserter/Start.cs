using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using TS_Post_Database_Inserter.Properties;

namespace TS_Post_Database_Inserter
{
    public partial class Start : Form
    {
        private readonly Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        public string Folder;
        public string Tempfolder;
        public string CurrentSrc;
        public string MainExcel = "";
        public string OpenPDF;

        private PdfReader reader;

        public Start()
        {
            InitializeComponent();


            //Testing
            //Config.AppSettings.Settings["MasterExcel"].Value = "";
            //Config.AppSettings.Settings["MainExcel"].Value = "";
            //Config.AppSettings.Settings["OpenPDF"].Value = "";
            //Config.AppSettings.Settings["Folder"].Value = "";

            //File Settings location
            Folder = Config.AppSettings.Settings["Folder"].Value;

            FindFiles(false);
            if (!String.IsNullOrEmpty(Folder) || !Directory.Exists(Tempfolder) || !Directory.Exists(CurrentSrc))
                Setup(Folder);
            CheckExcel();


            Config.Save(ConfigurationSaveMode.Full);

            UpdateUI();

            Config.AppSettings.Settings["OpenPDF"].Value = OpenPDF;
        }

        private void OpenMFol_Click(object sender, EventArgs e)
        {
            var fb = folderBrowserDialog1;
            if (fb.ShowDialog() == DialogResult.OK)
            {
                if (Folder != "")
                {
                    var mec = new MECheck(this, fb.SelectedPath, 0);
                    mec.ShowDialog();
                    Config.AppSettings.Settings["Folder"].Value = Folder;
                    Config.Save(ConfigurationSaveMode.Full);
                    FindFiles(false);
                    CheckExcel();
                    UpdateUI();
                }
                else
                {
                    Folder = fb.SelectedPath;
                    Config.AppSettings.Settings["Folder"].Value = Folder;
                    Config.Save(ConfigurationSaveMode.Full);
                    var mec = new MEC();
                    mec.ShowDialog();
                    FindFiles(false);
                    CheckExcel();
                    UpdateUI();
                }
            }
        }

        private void Launch_Click(object sender, EventArgs e)
        {
            //Excel.Application excel = new Excel.Application();
            //Excel.Workbook sheet = excel.Workbooks.Open(Excel_Path);
            //Excel.Worksheet x = excel.ActiveSheet as Excel.Worksheet;
            FindFiles(true);

            var c = new CNTL();
            if (Folder == "" || OpenPDF == "" || MainExcel == "")
                c.ShowDialog();
            else
                LaunchMethod();

            if (Folder == "")
            {
                MFol.Text = "Select a Master Excel document!!";
                MFol.ForeColor = Color.DarkRed;
            }

            //change
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
            var Cust = new CustInfo(this);
            Cust.ShowDialog();
        }

        private void Start_FormClosing(object sender, FormClosingEventArgs e)
        {
            Config.Save(ConfigurationSaveMode.Full);
        }


        ///////////////////////////////

        internal void CheckExcel()
        {
            Console.WriteLine(File.Exists(MainExcel));
            if (File.Exists(MainExcel))
            {
                var creation = File.GetCreationTime(MainExcel);
                var NewFile = Path.GetDirectoryName(Path.GetDirectoryName(MainExcel)) +
                              @"\Archives\XLSX\" + DateTime.Now.ToString("yyyy_MM_dd") + " Archived" +
                              Path.GetExtension(MainExcel);
                if ((creation.Date - DateTime.Now.Date).TotalDays > 62)
                {
                    File.Move(MainExcel, NewFile);
                    var st = new FileStream(MainExcel, FileMode.Create, FileAccess.ReadWrite);
                }
            }
            else
            {
                try
                {
                    File.WriteAllBytes(MainExcel, Resources.SourceExcel);
                }
                catch
                {
                    if (Folder == "")
                        throw new Exception("Failed to copy Source Excel");
                }
                finally
                {
                    Console.WriteLine("Excel creating passed");
                }
            }
        }

        private void FindFiles(bool create)
        {
            CurrentSrc = Folder + @"\Insert Label PDFs to edit";
            Tempfolder = Folder + @"\temp";
            MainExcel = CurrentSrc + @"\Main.xlsx";
            OpenPDF = CurrentSrc + @"\src.pdf";
            if (create)
                SortPDFS();
        }

        private void SortPDFS()
        {
            var srcPDF = CurrentSrc + @"\src.pdf";
            PdfReader.unethicalreading = true;
            try
            {
                var CheckFilesTemp = new List<string>();
                foreach (var files in Directory.GetFiles(CurrentSrc))
                    if (files.Contains(".pdf"))
                        CheckFilesTemp.Add(files);


                if (File.Exists(srcPDF) && CheckFilesTemp.Count > 1)
                {
                    if (!Directory.Exists(Tempfolder))
                        Directory.CreateDirectory(Tempfolder);
                    File.Move(srcPDF, Tempfolder + @"\src" + DateTime.Now.ToString("yyyy_MM_dd") + ".pdf");
                }

                if (Directory.Exists(CurrentSrc) && !CheckFilesTemp.Contains("src.pdf") && CheckFilesTemp.Count > 0)
                {
                    using (var stream = new MemoryStream())
                    {
                        using (var doc = new Document())
                        {
                            var pdf = new PdfCopy(doc, stream);
                            pdf.CloseStream = false;
                            doc.Open();

                            PdfReader reader = null;
                            PdfImportedPage page = null;

                            var FilesTemp = new List<string>();
                            foreach (var Files in Directory.GetFiles(CurrentSrc))
                                if (Files.Contains(".pdf"))
                                    FilesTemp.Add(Files);

                            try
                            {
                                if (FilesTemp.Count > 0)
                                    foreach (var file in FilesTemp)
                                    {
                                        reader = new PdfReader(file);
                                        for (var i = 0; i < reader.NumberOfPages; i++)
                                        {
                                            Console.WriteLine(i);
                                            page = pdf.GetImportedPage(reader, i + 1);
                                            pdf.AddPage(page);
                                        }

                                        pdf.FreeReader(reader);
                                        reader.Close();
                                    }
                                else
                                    throw new Exception("Not PDFs in source");
                            }
                            catch (Exception ee)
                            {
                                throw new Exception("Not PDFs in source", ee);
                            }
                        }

                        using (var streamX = new FileStream(srcPDF, FileMode.Create))
                        {
                            stream.WriteTo(streamX);
                        }
                    }

                    var h = 0;
                    foreach (var f in Directory.GetFiles(CurrentSrc))
                    {
                        h++;
                        if (!f.Contains("src.pdf"))
                            File.Move(f,
                                Folder + @"\Archives\PDF\" + DateTime.Now.ToString("yyyy_MM_dd") + " " + h + ".pdf");
                    }
                }
                else
                {
                    Console.WriteLine(CurrentSrc);
                    throw new Exception("CurrentSrc Folder doesn't exist or no pdfs in source file");
                }
            }
            catch (Exception ee)
            {
                Console.WriteLine(ee);
            }
        }

        private void UpdateUI()
        {
            var CheckFilesTemp = new List<string>();
            foreach (var files in Directory.GetFiles(CurrentSrc))
                if (files.Contains(".pdf"))
                    CheckFilesTemp.Add(files);

            if (Folder != "")
            {
                MFol.Text = Folder;
                MFol.ForeColor = Color.Black;
                PHExcelL.Text = MainExcel;
                PHExcelL.ForeColor = Color.Black;

                int index = 0;
                foreach(var file in CheckFilesTemp)
                    if (!file.Contains("src.pdf"))
                        index++;

                if (CheckFilesTemp.Contains("src.pdf"))
                {
                    LpdfL.Text = OpenPDF;
                    LpdfL.ForeColor = Color.Black;
                    reader = new PdfReader(OpenPDF);
                    PDFL.Text = "Number of Labels found: " + reader.NumberOfPages;
                }
                else
                {

                    PDFL.ForeColor = Color.DarkRed;
                    PDFL.Text =
                        "Error: unable to find source PDFs.\r\nMake sure Label PDFs are in folder 'insert label pdf here'";

                    LpdfL.Text = "Error";
                    LpdfL.ForeColor = Color.DarkRed;
                }
            }
            else
            {
                MFol.Text = "Select a Master folder!!";
                MFol.ForeColor = Color.DarkRed;
                LpdfL.Text = "Select a Label PDF file!!";
                LpdfL.ForeColor = Color.DarkRed;
                PDFL.Text = "";
                PHExcelL.Text = "No main Excel found";
                PHExcelL.ForeColor = Color.DarkRed;
            }
        }

        private void SetupMasFod_Click(object sender, EventArgs e)
        {
            var folderBrowser = folderBrowserDialog1;
            if (folderBrowser.ShowDialog() != DialogResult.OK) return;
            var folder = folderBrowser.SelectedPath;
            Setup(Folder);
        }

        private void Setup(String folder)
        {
            Directory.CreateDirectory(folder + @"\Insert Label PDFs to edit");
            Directory.CreateDirectory(folder + @"\temp");
            Directory.CreateDirectory(folder + @"\Archives");
            Directory.CreateDirectory(folder + @"\Archives\PDF");
            Directory.CreateDirectory(folder + @"\Archives\XLSX");
            Directory.CreateDirectory(folder + @"\Master");
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OpenMDIR_Click(object sender, EventArgs e)
        {
            if (Folder != null || Folder != "")
            {
                try{Process.Start(Folder);}
                catch (Exception ee){throw new Exception("Failed to open folder", ee);}
            }
        }
    }
}