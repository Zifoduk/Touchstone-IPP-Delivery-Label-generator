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
using NPOI.XSSF.UserModel;

namespace TS_Post_Database_Inserter
{

    [ExceptionWrapper]
    public partial class Start : Form
    {
        private readonly Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);
        public string Folder;
        public string Tempfolder;
        public string CurrentSrc;
        public string MainExcel = "";
        public string MasterExcel = "";
        public string OpenPDF;
        public string Archives;
        public string Master;

        private PdfReader reader;

        public XSSFWorkbook Workbook;
        public XSSFSheet WorkSheet;
        public FileStream Filestream;

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
            bool FolderExists;
            bool MainExcelExists;
            bool OpenPDFExists;

            if (Directory.Exists(Folder))
                FolderExists = true;
            else FolderExists = false;

            if (File.Exists(MainExcel))
                MainExcelExists = true;
            else MainExcelExists = false;

            if (File.Exists(OpenPDF))
                OpenPDFExists = true;
            else OpenPDFExists = false;


            FindFiles(true);

            var c = new CNTL();
            if (!FolderExists || !OpenPDFExists || !MainExcelExists)
                c.ShowDialog();
            else if (FolderExists && OpenPDFExists && MainExcelExists)
            {
                Console.WriteLine("Folder: " + Folder + ", OpenPDF: " + OpenPDF + ", MainExcel: " + MainExcel);
                LaunchMethod();
            }

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
            try
            {
                if (!File.Exists(MasterExcel))
                    File.WriteAllBytes(MasterExcel, Resources.SourceExcel);
                Filestream = new FileStream(MasterExcel, FileMode.Open, FileAccess.ReadWrite);
                Workbook = new XSSFWorkbook(Filestream);
                WorkSheet = Workbook.GetSheetAt(0) as XSSFSheet;
            }
            catch (Exception ex)
            {
                throw new ExcelDocumentOpenException(ex);
            }
            finally
            {
                Console.WriteLine("error passed");
            }

            Console.WriteLine(File.Exists(MainExcel));
            if (File.Exists(MainExcel))
            {
                var creation = File.GetCreationTime(MainExcel);
                var NewFile = Path.GetDirectoryName(Path.GetDirectoryName(MainExcel)) +
                              @"\Archives\XLSX\" + DateTime.Now.ToString("yyyy_MM_dd") + " Archived" +
                              Path.GetExtension(MainExcel);
                if ((creation.Date - DateTime.Now.Date).TotalDays > 62)
                {
                    if (File.Exists(MasterExcel))
                    {
                        int Endrow = WorkSheet.PhysicalNumberOfRows;
                    }
                    File.Move(MainExcel, NewFile);
                    File.WriteAllBytes(MainExcel, Resources.SourceExcel);
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
            Archives = Folder + @"\Archives";
            Master = Folder + @"\Master";
            MasterExcel = Master + @"\Master.xlsx";
            if (create)
                SortPDFS();
        }

        private void SortPDFS()
        {
            var srcPDF = CurrentSrc + @"\src.pdf";
            PdfReader.unethicalreading = true;
            var CheckFilesTemp = new List<string>();
            try
            {
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
                                {
                                }

                                //throw new Exception("Not PDFs in source");
                            }
                            catch (Exception ee)
                            {
                                throw new Exception("Not PDFs in source", ee);
                            }
                            finally
                            {
                                pdf.Close();
                                doc.Close();
                            }

                        }

                        using (var streamX = new FileStream(srcPDF, FileMode.Create))
                        {
                            stream.WriteTo(streamX);
                            streamX.Close();
                        }

                        stream.Close();
                    }

                  
                }

                Console.WriteLine(CheckFilesTemp.Count);
                Console.WriteLine("After");

            }
            catch (Exception ee)
            {
                Console.WriteLine(ee);
            }
            finally
            {
                for (var index = 0; index < CheckFilesTemp.Count; index++)
                {
                    var file = CheckFilesTemp[index];
                    if (!file.Contains("src.pdf"))
                    {
                        File.SetLastWriteTime(file, DateTime.Now);
                        File.Move(file,
                            Archives + @"\PDF\" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm") + " " + (index + 1) + ".pdf");
                        
                    }
                }
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

                if(CheckFilesTemp.Count > 0)
                    foreach(var file in CheckFilesTemp) { 
                        if (file.Contains("src.pdf"))
                        {
                            LpdfL.Text = OpenPDF;
                            LpdfL.ForeColor = Color.Black;
                            reader = new PdfReader(OpenPDF);
                            PDFL.Text = "Number of Labels found: " + reader.NumberOfPages;
                        }
                        else
                        {
                            LpdfL.Text = "Error";
                            LpdfL.ForeColor = Color.DarkRed;
                            PDFL.ForeColor = Color.DarkRed;
                            PDFL.Text =
                                "Error: unable to find source PDFs.\r\nMake sure Label PDFs are in folder 'Insert Label PDFs to edit'";

                        }
                }else
                {
                    LpdfL.Text = "Error";
                    LpdfL.ForeColor = Color.DarkRed;
                    PDFL.ForeColor = Color.DarkRed;
                    PDFL.Text = "Src PDF not created - Refer to Manual(Section 2.3)";
                    PDFNum.Text =
                        "Error: unable to find source PDFs.\r\nMake sure Label PDFs are in folder 'Insert Label PDFs to edit'";
                    PDFNum.ForeColor = Color.DarkRed;
                }


                if (index > 0)
                { 
                    PDFNum.Text = "Number of PDFs found in folder: " + index;
                    PDFNum.ForeColor = Color.Black;
                }
                else
                {
                    PDFNum.Text = "No PDFs found, Copy and Paste Label PDFs into\r\n folder";
                    PDFNum.ForeColor = Color.DarkRed;

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
            var folder = "";
            var folderBrowser = folderBrowserDialog1;
            if (folderBrowser.ShowDialog() == DialogResult.OK)
            {
                Console.WriteLine("First check: " + folderBrowser.SelectedPath);
                folder = folderBrowser.SelectedPath;
                Console.WriteLine("Second check: " + folder);
                Setup(folder);
            }
            Console.WriteLine("Third check: " + folder);

            DialogResult changeFolderDialoge = MessageBox.Show("Do you want to set the new master folder as the default?", "Maybe", MessageBoxButtons.YesNo);
            if (changeFolderDialoge == DialogResult.Yes)
            {
                Folder = folder;
                Console.WriteLine("Folder: " + Folder);
                FindFiles(false);
                UpdateUI();
            }

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