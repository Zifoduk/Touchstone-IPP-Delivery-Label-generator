using System;
using System.Collections.Generic;
using System.Configuration;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using iTextSharp.text;
using iTextSharp.text.pdf;
using TS_Post_Database_Inserter.Properties;
using NPOI.XSSF.UserModel;
using Syroot.Windows.IO;

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

        public SelectedTree? selectedTree = null;

        private PdfReader reader;

        public XSSFWorkbook masterWorkbook;
        public XSSFSheet masterSheet;
        public FileStream masterFileStream;
        
        public Completed _completed;

        public Start()
        {
            InitializeComponent();


            //Testing
            Config.AppSettings.Settings["MainExcel"].Value = "";
            Config.AppSettings.Settings["OpenPDF"].Value = "";
            Config.AppSettings.Settings["Folder"].Value = "";

            //File Settings location
            Folder = Config.AppSettings.Settings["Folder"].Value;

            FindFiles();
            if (!String.IsNullOrEmpty(Folder) || !Directory.Exists(Tempfolder) || !Directory.Exists(CurrentSrc))
                Setup(Folder);
            CheckExcel();

            UpdateUI();
            SaveLocation();

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
                    FindFiles();
                    CheckExcel();
                    UpdateUI();
                    SaveLocation();
                }
                else
                {
                    Folder = fb.SelectedPath;
                    Config.AppSettings.Settings["Folder"].Value = Folder;
                    var mec = new MEC();
                    mec.ShowDialog();
                    FindFiles();
                    CheckExcel();
                    UpdateUI();
                    SaveLocation();
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


            FindFiles();
            
            if (!FolderExists || !OpenPDFExists || !MainExcelExists)
                MessageBox.Show("Error 455: Cannot launch Edittor\r\nFix settings in RED", "Error 455",MessageBoxButtons.OK);
            else if (FolderExists && OpenPDFExists && MainExcelExists)
            {
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
            SaveLocation();
        }


        ///////////////////////////////

        internal void CheckExcel()
        {
            if (!File.Exists(MasterExcel))
            {
                try
                {
                    File.WriteAllBytes(MasterExcel, Resources.SourceExcel);
                }
                catch
                {
                    if (Folder == "")
                        throw new Exception("Failed to copy Source Excel to Master");
                }
                finally
                {
                    Console.WriteLine("Excel creating passed");
                }
            }


            if (File.Exists(MainExcel) && File.Exists(MasterExcel))
            {
                var lastWriteTime = File.GetLastWriteTime(MainExcel);
                var timeDifference = (TimeSpan)lastWriteTime.Subtract(DateTime.Now);
                var differenceOfTime = (int) Math.Round(Math.Abs(timeDifference.TotalMinutes));


                if (differenceOfTime > 30)
                {
                    /////Write to Master Excel for archive if main exists and is over time limit
                    Int64 masterChecksize = 0;
                    Int64 mainChecksize;
                    if (File.Exists(MasterExcel))
                        masterChecksize = new FileInfo(MasterExcel).Length;
                    if (File.Exists(MainExcel))
                        mainChecksize = new FileInfo(MainExcel).Length;

                    if (masterChecksize < 10)
                    {
                        if (File.Exists(MasterExcel))
                            File.Delete(MasterExcel);
                        File.WriteAllBytes(Master + @"\Master.xlsx", Resources.SourceExcel);
                    }


                    FileStream tempMasterFileStream = new FileStream(MasterExcel, FileMode.Open, FileAccess.ReadWrite);
                    XSSFWorkbook masterWorkbook = new XSSFWorkbook(tempMasterFileStream);
                    XSSFSheet masterSheet = masterWorkbook.GetSheetAt(0) as XSSFSheet;

                    var copy = false;
                    using (FileStream masterFileStream = new FileStream(MasterExcel, FileMode.Create, FileAccess.ReadWrite))
                    {

                        int masterMaxRow = masterSheet.PhysicalNumberOfRows;
                        int masterCurrentRow = masterMaxRow;


                        using (var mainFileStream = new FileStream(MainExcel, FileMode.Open, FileAccess.ReadWrite))
                        {
                            XSSFWorkbook mainWorkbook = new XSSFWorkbook(mainFileStream);
                            if (mainWorkbook.GetSheetAt(0) is XSSFSheet mainSheet)
                            {

                                var mainRows = mainSheet.PhysicalNumberOfRows - 1;
                                var mainColumns = mainSheet.GetRow(0).LastCellNum;
                                if (mainRows > 0)
                                {
                                    copy = true;
                                    for (int y = 1; y <= mainRows; y++)
                                    {
                                        var dontWrite = 0;
                                        for (int i = 1; i < masterMaxRow; i++)
                                        {
                                            var barcodeCellContentMaster =
                                                masterSheet.GetRow(i).GetCell(2).ToString();
                                            var barcodeCellContentMain =
                                                mainSheet.GetRow(y).GetCell(2).ToString();
                                            if (barcodeCellContentMaster == barcodeCellContentMain)
                                                dontWrite++;
                                        }
                                        if (dontWrite == 0)
                                        {
                                            masterSheet.CreateRow(masterCurrentRow);
                                            for (int x = 0; x < mainColumns; x++)
                                            {
                                                masterSheet.GetRow(masterCurrentRow).CreateCell(x);
                                                masterSheet.GetRow(masterCurrentRow).GetCell(x)
                                                    .SetCellValue(mainSheet.GetRow(y).GetCell(x).ToString());
                                            }

                                            masterCurrentRow += 1;
                                        }
                                    }
                                }
                            }
                            mainFileStream.Close();
                            mainFileStream.Dispose();

                        }
                        masterFileStream.Close();
                        masterFileStream.Dispose();
                    }

                    using (FileStream masterFileStream = new FileStream(MasterExcel, FileMode.Create, FileAccess.Write))
                    {
                        masterWorkbook.Write(masterFileStream);
                    }



                    ////Archieve mainexcel and recreate
                    if (copy)
                    {
                        var newExcelArchiveFile = Path.GetDirectoryName(Path.GetDirectoryName(MainExcel)) +
                                                  @"\Archives\XLSX\" + DateTime.Now.ToString("yyyy_MM_dd_HH-MM-ss") +
                                                  " Archived" +
                                                  Path.GetExtension(MainExcel);
                        File.Move(MainExcel, newExcelArchiveFile);
                        File.WriteAllBytes(MainExcel, Resources.SourceExcel);
                    }
                    copy = false;
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

        private void FindFiles()
        {
            CurrentSrc = Folder + @"\Insert Label PDFs to edit";
            Tempfolder = Folder + @"\temp";
            MainExcel = CurrentSrc + @"\Main.xlsx";
            OpenPDF = CurrentSrc + @"\src.pdf";
            Archives = Folder + @"\Archives";
            Master = Folder + @"\Master";
            MasterExcel = Master + @"\Master.xlsx";
        }

        private void MoveSrc()
        {
            var srcPdf = CurrentSrc + @"\src.pdf";
            PdfReader.unethicalreading = true;

            var newSrc = Tempfolder + @"\src" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm") + ".pdf";
            
            if (File.Exists(srcPdf))
            {
                if (!Directory.Exists(Tempfolder))
                    Directory.CreateDirectory(Tempfolder);
                File.Move(srcPdf, newSrc);
                if (File.Exists(newSrc))
                    Console.WriteLine("newSrc Exists");
            }
        }

        private void AggregatePdfs()
        {
            var srcPdf = CurrentSrc + @"\src.pdf";
            PdfReader.unethicalreading = true;

            var checkFilesTemp = new List<string>();
            foreach (var files in Directory.GetFiles(CurrentSrc))
                if (files.Contains(".pdf"))
                    checkFilesTemp.Add(files);

            try
            {
                if (Directory.Exists(CurrentSrc) && !checkFilesTemp.Contains("src.pdf") && checkFilesTemp.Count > 0)
                {
                    using (var stream = new MemoryStream())
                    {
                        using (var doc = new Document())
                        {
                            var pdf = new PdfCopy(doc, stream) {CloseStream = false};
                            doc.Open();

                            PdfReader reader = null;
                            var filesTemp = Directory.GetFiles(CurrentSrc).Where(File => File.Contains(".pdf")).ToList();

                            try
                            {
                                if (filesTemp.Count > 0)
                                    foreach (var file in filesTemp)
                                    {
                                        reader = new PdfReader(file);
                                        for (var i = 0; i < reader.NumberOfPages; i++)
                                        {
                                            PdfImportedPage page = null;
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

                        using (var streamX = new FileStream(srcPdf, FileMode.Create))
                        {
                            stream.WriteTo(streamX);
                            streamX.Close();
                        }

                        stream.Close();
                    }
                }

            }
            catch (Exception ee)
            {
                Console.WriteLine(ee);
            }
            finally
            {

            }
        }

        public void ArchivePDF()
        {

            var FilesTemp = new List<string>();
            foreach (var Files in Directory.GetFiles(CurrentSrc))
                if (Files.Contains(".pdf"))
                    FilesTemp.Add(Files);

            for (var index = 0; index < FilesTemp.Count; index++)
            {
                var file = FilesTemp[index];
                if (!file.Contains("src.pdf"))
                {
                    File.SetLastWriteTime(file, DateTime.Now);
                    File.Move(file,
                        Archives + @"\PDF\" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm") + " " + (index + 1) + ".pdf");

                }
            }
        }

        private void UpdateUI()
        {
            sourceTree.Nodes.Clear();
            downloadsTree.Nodes.Clear();
            DirectoryInfo directoryInfo = null;
            
            directoryInfo = new DirectoryInfo(KnownFolders.Downloads.Path);
            foreach (FileInfo fileInfo in directoryInfo.GetFiles())
            {
                TreeNode d_fileNode = new TreeNode();
                d_fileNode.Text = fileInfo.Name;
                if (fileInfo.Name.Contains(".pdf"))
                    downloadsTree.Nodes.Add(d_fileNode);
            }

            var CheckFilesTemp = new List<string>();
            foreach (var files in Directory.GetFiles(CurrentSrc))
                if (files.Contains(".pdf"))
                    CheckFilesTemp.Add(files);

            if (Folder != "")
            {

                MFol.Text = Folder;
                MFol.ForeColor = Color.Black;
                if (Directory.Exists(CurrentSrc))
                {
                    directoryInfo = new DirectoryInfo(CurrentSrc);
                    foreach (FileInfo fileInfo in directoryInfo.GetFiles())
                    {
                        TreeNode fileNode = new TreeNode();
                        fileNode.Text = fileInfo.Name;
                        fileNode.ImageIndex = 0;
                        fileNode.SelectedImageIndex = 0;
                        if (fileInfo.Name.Contains(".pdf"))
                            sourceTree.Nodes.Add(fileNode);
                    }
                }
                else
                {
                    
                    TreeNode fileNode = new TreeNode();
                    fileNode.Text = "error folder not found";
                    sourceTree.Nodes.Add(fileNode);
                    
                }
                    ///downloads tree
                //webBrowser1.Document.GetElementById("menu").Style = "display:none";
                PHExcelL.Text = MainExcel;
                PHExcelL.ForeColor = Color.Black;

                int index = 0;

                if(CheckFilesTemp.Count > 0) {
                    foreach (var file in CheckFilesTemp)
                    {
                        if (file.Contains("src.pdf"))
                        {
                            LpdfL.Text = OpenPDF;
                            LpdfL.ForeColor = Color.Black;
                            reader = new PdfReader(OpenPDF);
                            PDFL.Text = "Number of Labels found: " + reader.NumberOfPages;
                            reader.Dispose();
                            reader.Close();
                        }
                        else
                        {
                            index ++;
                            LpdfL.Text = "Error";
                            LpdfL.ForeColor = Color.DarkRed;
                            PDFL.ForeColor = Color.DarkRed;
                            PDFL.Text =
                                "Error: unable to find source PDFs.\r\nMake sure Label PDFs are in folder 'Insert Label PDFs to edit'";

                        }

                    }
                }
                else
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
                folder = folderBrowser.SelectedPath;
                Setup(folder);
                DialogResult changeFolderDialoge = MessageBox.Show("Do you want to set the new master folder as the default?", "Set as default", MessageBoxButtons.YesNo);
                if (changeFolderDialoge == DialogResult.Yes)
                {
                    Folder = folder;
                    FindFiles();
                    UpdateUI();
                }
            }
        }

        private void Setup(String folder)
        {
            if(!Directory.Exists(folder + @"\Insert Label PDFs to edit"))
                Directory.CreateDirectory(folder + @"\Insert Label PDFs to edit");
            if (!Directory.Exists(folder + @"\temp"))
                Directory.CreateDirectory(folder + @"\temp");
            if (!Directory.Exists(folder + @"\Archives"))
                Directory.CreateDirectory(folder + @"\Archives");
            if (!Directory.Exists(folder + @"\Archives\PDF"))
                Directory.CreateDirectory(folder + @"\Archives\PDF");
            if (!Directory.Exists(folder + @"\Archives\XLSX"))
                Directory.CreateDirectory(folder + @"\Archives\XLSX");
            if (!Directory.Exists(folder + @"\Master"))
                Directory.CreateDirectory(folder + @"\Master");
            if (!File.Exists(folder + @"\Label.lbx"))
                File.WriteAllBytes(folder + @"\Label.lbx", Resources.Label);
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OpenMDIR_Click(object sender, EventArgs e)
        {
            if (Folder != null || Folder != "")
            {
                try
                {
                    Process.Start(Folder);
                }
                catch (Exception ee)
                {
                    MessageBox.Show("Master Folder isnt selected");
                }
            }
        }

        private void RefreshBtn_Click(object sender, EventArgs e)
        {
            DialogResult CheckCorrectPDFsInFolder =
                MessageBox.Show("Have all of the required Labels been copied into the correct folder?",
                    "Check your labels", MessageBoxButtons.YesNo);
            if (CheckCorrectPDFsInFolder == DialogResult.Yes)
            {
                MoveSrc();
                FindFiles();
                CheckExcel();
                AggregatePdfs();
                UpdateUI();
                
            }

        }

        private void SaveLocation()
        {
            Config.AppSettings.Settings["Folder"].Value = Folder;
            try
            {

            }
            catch (Exception e)
            {
                MessageBox.Show("Error:" + e, "Error thrown", MessageBoxButtons.OK, MessageBoxIcon.Error);
                Console.WriteLine(e);
                throw;
            }

            Config.Save(ConfigurationSaveMode.Full, true);
        }

        private void moveLeftBtn_Click(object sender, EventArgs e)
        {
            if (selectedTree == SelectedTree.downloads)
            {
                var fileName = downloadsTree.SelectedNode.Text;
                var fileFullDirec = (KnownFolders.Downloads.Path + @"\" + fileName);
                Console.WriteLine(fileFullDirec);
                var newSrcFileName = (CurrentSrc + @"\" + fileName);
                if (!File.Exists(newSrcFileName))
                {
                    Console.WriteLine("Okay");
                    File.Move(fileFullDirec, newSrcFileName);
                    UpdateUI();
                }
                else
                {
                    Console.WriteLine("Fail");
                    MessageBox.Show("Cant move, File already exists in source folder", "Error", MessageBoxButtons.OK);
                    UpdateUI();
                }
            }
        }

        private void moveRightBtn_Click(object sender, EventArgs e)
        {
            if (selectedTree == SelectedTree.source)
            {
                var fileName = sourceTree.SelectedNode.Text;
                var fileFullDirec = (CurrentSrc + @"\" + fileName);
                Console.WriteLine(fileFullDirec);
                var newDownFileName = (KnownFolders.Downloads.Path + @"\" + fileName);
                if (!File.Exists(newDownFileName))
                {
                    Console.WriteLine("Okay");
                    File.Move(fileFullDirec, newDownFileName);
                    UpdateUI();
                }
                else
                {
                    Console.WriteLine("Fail");
                    MessageBox.Show("Cant move, File already existsin downloads folder", "Error", MessageBoxButtons.OK);
                    UpdateUI();
                }
            }
        }

        private void changedTreeSelected(object sender, TreeNodeMouseClickEventArgs e)
        {
            TreeView obj = sender as TreeView;
            Console.WriteLine(obj.Name);
            if (obj.Name == "sourceTree")
            {
                selectedTree = SelectedTree.source;
            }
            else if (obj.Name == "downloadsTree")
            {
                selectedTree = SelectedTree.downloads;
            }
        }
    }
    public enum SelectedTree
    {
        downloads,
        source
        
    }
}
