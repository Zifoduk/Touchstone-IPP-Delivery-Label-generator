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
using System.Threading;
using System.Threading.Tasks;
using System.Security.Permissions;
using iTextSharp.text.pdf.parser;
using System.Reflection;

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
        private int NumberofPDFs = 0;

        public SelectedTree? selectedTree = null;

        private PdfReader reader;

        public XSSFWorkbook masterWorkbook;
        public XSSFSheet masterSheet;
        public FileStream masterFileStream;
        
        public Completed _completed;

        private int? SelectedIndexD = null;
        private int? SelectedIndexS = null;




        //Main Operations

        public Start()
        {
            InitializeComponent();

            WaitLabel.Enabled = false;
            WaitLabel.Visible = false;
            //Testing
            /*Config.AppSettings.Settings["MainExcel"].Value = "";
            Config.AppSettings.Settings["OpenPDF"].Value = "";
            Config.AppSettings.Settings["Folder"].Value = "";*/

            //File Settings location
            Folder = Config.AppSettings.Settings["Folder"].Value;

            FindFiles();
            if ((!String.IsNullOrEmpty(Folder) && Directory.Exists(Folder)) && (!Directory.Exists(Tempfolder) || !Directory.Exists(CurrentSrc)))
                Setup(Folder);
            CheckExcel();

            UpdateUI();
            SaveLocation();

            Config.AppSettings.Settings["OpenPDF"].Value = OpenPDF;


            //Staring New Thread for Update
            //AutoUpdateUI();
            RunWatcher();

        }

        private bool IsValidPdf(string filepath)
        {
            bool Ret = true;

            PdfReader pdfReader = null;

            try
            {
                pdfReader = new PdfReader(filepath);
            }
            catch
            {
                Ret = false;
            }

            return Ret;
        }
        public void UpdateUI()
        {
            PdfReader.unethicalreading = true;
            sourceTree.Nodes.Clear();
            downloadsTree.Nodes.Clear();
            DirectoryInfo directoryInfo = null;

            //Downloads tree
            directoryInfo = new DirectoryInfo(KnownFolders.Downloads.Path);
            foreach (FileInfo fileInfo in directoryInfo.GetFiles())
            {
                if (IsValidPdf(fileInfo.FullName))
                {
                    var d_fileNode = new TreeNode
                    {
                        Text = fileInfo.Name
                    };

                    PdfReader reader = new PdfReader(fileInfo.FullName);

                    string[] PdfContent = new string[] { };


                    for (var i = 1; i <= reader.NumberOfPages; i++)
                    {
                        var temp = PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
                        PdfContent = temp.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                    }

                    if (fileInfo.Name.Contains(".pdf"))
                        foreach (string s in PdfContent)
                            if (s.Contains("ipostparcels"))
                                downloadsTree.Nodes.Add(d_fileNode);
                }
            }


            //Source Tree
            var CheckFilesTemp = new List<string>();
            if (Directory.Exists(CurrentSrc))
                foreach (var files in Directory.GetFiles(CurrentSrc))
                    if (files.Contains(".pdf"))
                        CheckFilesTemp.Add(files);

            if (Folder != "" && Directory.Exists(Folder))
            {

                MFol.Text = Folder;
                MFol.ForeColor = Color.Black;
                if (Directory.Exists(CurrentSrc))
                {
                    directoryInfo = new DirectoryInfo(CurrentSrc);
                    foreach (FileInfo fileInfo in directoryInfo.GetFiles())
                    {
                        if (IsValidPdf(fileInfo.FullName))
                        { 
                            var fileNode = new TreeNode
                            {
                                Text = fileInfo.Name,
                                ImageIndex = 0,
                                SelectedImageIndex = 0
                            };

                            PdfReader reader = new PdfReader(fileInfo.FullName);

                            string[] PdfContent = new string[] { };
                            for (var i = 1; i <= reader.NumberOfPages; i++)
                            {
                                var temp = PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
                                PdfContent = temp.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                            }

                            if (fileInfo.Name.Contains(".pdf"))
                                foreach (string s in PdfContent)
                                    if (s.Contains("ipostparcels"))
                                        sourceTree.Nodes.Add(fileNode);
                        }
                    }
                }
                else
                {

                    TreeNode fileNode = new TreeNode
                    {
                        Text = "error folder not found"
                    };
                    sourceTree.Nodes.Add(fileNode);

                }

                //webBrowser1.Document.GetElementById("menu").Style = "display:none";
                PHExcelL.Text = MainExcel;
                PHExcelL.ForeColor = Color.Black;

                int index = 0;

                if (CheckFilesTemp.Count > 0)
                {
                    foreach (var file in CheckFilesTemp)
                    {
                        if (file.Contains("src.pdf"))
                        {
                            LpdfL.Text = OpenPDF;
                            LpdfL.ForeColor = Color.Black;
                            reader = new PdfReader(OpenPDF);
                            PDFL.ForeColor = Color.Black;
                            PDFL.Text = "Number of Labels found: " + reader.NumberOfPages;
                            reader.Dispose();
                            reader.Close();
                        }
                        else
                        {
                            index++;
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
                    NumberofPDFs = index;
                    PDFNum.ForeColor = Color.Black;
                }
                else
                {
                    PDFNum.Text = "No PDFs found, Copy and Paste Label PDFs into\r\n folder";
                    PDFNum.ForeColor = Color.DarkRed;
                    NumberofPDFs = 0;

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

        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        private void RunWatcher()
        {
            FileSystemWatcher watcherDownloads = new FileSystemWatcher();
            FileSystemWatcher watcherSource = new FileSystemWatcher(); 
            watcherDownloads.Path = KnownFolders.Downloads.Path;
            watcherSource.Path = CurrentSrc;

            watcherDownloads.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite 
                | NotifyFilters.FileName | NotifyFilters.DirectoryName;

            watcherSource.NotifyFilter = NotifyFilters.LastAccess | NotifyFilters.LastWrite
                | NotifyFilters.FileName | NotifyFilters.DirectoryName;

            watcherDownloads.Filter = "*.pdf";
            watcherSource.Filter = "*.pdf";

            watcherDownloads.Changed += new FileSystemEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Downloads", "Changed"));
            watcherDownloads.Created += new FileSystemEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Downloads", "Created"));
            watcherDownloads.Deleted += new FileSystemEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Downloads", "Deleted"));
            watcherDownloads.Renamed += new RenamedEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Downloads", "Renamed"));

            watcherSource.Changed += new FileSystemEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Source", "Changed"));
            watcherSource.Created += new FileSystemEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Source", "Created"));
            watcherSource.Deleted += new FileSystemEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Source", "Deleted"));
            watcherSource.Renamed += new RenamedEventHandler((sender, args) => 
                                        WatcherChanged(sender, args, "Source", "Renamed"));

            watcherDownloads.EnableRaisingEvents = true;
            watcherSource.EnableRaisingEvents = true;

        }

        internal void LaunchMethod()
        {
            var Cust = new CustInfo(this);
            Cust.ShowDialog();
        }
        
        //private delegate void SetControlPropertyDelegate(Control control, String property, object propertyValue);
        /*private static void SetControlProperty(Control control, string property, object propertyValue)
        {
            if(control.InvokeRequired) 
                control.Invoke(new SetControlPropertyDelegate(setControlProperty), 
                    new object[] { control, property, propertyValue });            
            else 
                control.GetType().InvokeMember(property, BindingFlags.SetProperty, null, control, new object[] { control, property, propertyValue });
        }*/
        private bool Wait = false;
        public bool WaitStarted = false;
        private void Waiter()
        {
            WaitStarted = true;
            while (true)
            {
                if(Wait)
                {
                    WaitLabel.Invoke(new Action(() => WaitLabel.Enabled = true));
                    WaitLabel.Invoke(new Action(() => WaitLabel.Visible = true));
                    WaitLabel.Invoke(new Action(() => moveLeftBtn.Enabled = false));
                    WaitLabel.Invoke(new Action(() => moveRightBtn.Enabled = false));
                }
                else
                {
                    WaitLabel.Invoke(new Action(() => WaitLabel.Enabled = false));
                    WaitLabel.Invoke(new Action(() => WaitLabel.Visible = false));
                    WaitLabel.Invoke(new Action(() => moveLeftBtn.Enabled = true));
                    WaitLabel.Invoke(new Action(() => moveRightBtn.Enabled = true));
                    break;
                }
            }
            WaitStarted = false;
            Console.WriteLine("end loop");
        }



        //File Operations

        internal void CheckExcel()
        {
            Int64 mainChecksize = 0;
            if (File.Exists(MainExcel))
                mainChecksize = new FileInfo(MainExcel).Length;

            if (mainChecksize < 10)
            {
                if (File.Exists(MainExcel))
                    File.Delete(MainExcel);
                File.WriteAllBytes(MainExcel, Resources.SourceExcel);
            }

            if (File.Exists(MainExcel) && File.Exists(MasterExcel))
            {
                var lastWriteTime = File.GetLastWriteTime(MainExcel);
                var timeDifference = (TimeSpan)lastWriteTime.Subtract(DateTime.Now);
                var differenceOfTime = (int)Math.Round(Math.Abs(timeDifference.TotalMinutes));


                if (differenceOfTime > 30)
                {
                    /////Write to Master Excel for archive if main exists and is over time limit
                    Int64 masterChecksize = 0;
                    
                    if (File.Exists(MasterExcel))
                        masterChecksize = new FileInfo(MasterExcel).Length;

                    if (masterChecksize < 10)
                    {
                        if (File.Exists(MasterExcel))
                            File.Delete(MasterExcel);
                        File.WriteAllBytes(Master + @"\Master.xlsx", Resources.SourceExcel);
                    }



                    FileStream tempMasterFileStream = new FileStream(MasterExcel, FileMode.Open
                                                            , FileAccess.ReadWrite);
                    XSSFWorkbook masterWorkbook = new XSSFWorkbook(tempMasterFileStream);
                    XSSFSheet masterSheet = masterWorkbook.GetSheetAt(0) as XSSFSheet;

                    var copy = false;
                    using (FileStream masterFileStream = new FileStream(MasterExcel, FileMode.Create
                                                                        , FileAccess.ReadWrite))
                    {

                        int masterMaxRow = masterSheet.PhysicalNumberOfRows;
                        int masterCurrentRow = masterMaxRow;


                        using (var mainFileStream = new FileStream(MainExcel, FileMode.Open
                                                            , FileAccess.ReadWrite))
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
                        var newExcelArchiveFile = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(MainExcel)) +
                                                  @"\Archives\XLSX\" + DateTime.Now.ToString("yyyy_MM_dd_HH-MM-ss") +
                                                  " Archived" + System.IO.Path.GetExtension(MainExcel);
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

            var newSrc = Tempfolder + @"\src" + DateTime.Now.ToString("yyyy_MM_dd_HH_mm_ss") + ".pdf";
            
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
            checkFilesTemp.Clear();
            if(Directory.Exists(CurrentSrc))
                foreach (var files in Directory.GetFiles(CurrentSrc))
                    if (files.Contains(".pdf"))
                        checkFilesTemp.Add(files);

            try
            {
                /*if (Directory.Exists(CurrentSrc) && !checkFilesTemp.Contains("src.pdf")
                    && checkFilesTemp.Count > 0)*/
                if(Directory.Exists(CurrentSrc) && checkFilesTemp.Count > 0)
                {

                    if (File.Exists(srcPdf))
                        MoveSrc();
                    
                    using (var stream = new MemoryStream())
                    {
                        using (var doc = new Document())
                        {
                            var pdf = new PdfCopy(doc, stream) {CloseStream = false};
                            doc.Open();

                            PdfReader reader = null;
                            var filesTemp = Directory.GetFiles(CurrentSrc).Where(File => 
                                                        File.Contains(".pdf")).ToList();

                            foreach(var file in filesTemp)
                                if (file.Contains("src.pdf"))
                                    filesTemp.Remove(file);
                            

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


            Wait = false;
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

        private void Setup(String folder)
        {
            var _directories = new string[] {@"\Insert Label PDFs to edit", @"\temp", @"\Archives", @"\Archives\PDF", @"\Archives\XLSX", @"\Master"};
            var _labelFiles = (string)(@"/Label.lbx");

            foreach (var l in _directories)
            {
                if (!Directory.Exists(folder + l))
                {
                    Directory.CreateDirectory(folder + l);
                }
            }

            if (!File.Exists(folder + _labelFiles))
                File.WriteAllBytes(folder + _labelFiles, Resources.Label);
        }




        //End Operations

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

        public void OpenLBX()
        {
            try
            {
                System.Diagnostics.Process.Start(Folder + @"\Label.lbx");
                Console.WriteLine("Check if process started");
            }
            catch (Exception ee)
            {
                Console.WriteLine("Inner: " + ee.InnerException + ", Exception: " + ee);
                throw;
            }
        }




        //Events

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
                    Setup(Folder);
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
                    Setup(Folder);
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
                MessageBox.Show("Error 455: Cannot launch Edittor\r\nFix settings in RED", "Error 455", MessageBoxButtons.OK);
            else if (NumberofPDFs <= 0)
                MessageBox.Show("Error 487: No Label PDFs present in the Source Folder", "Error 487", MessageBoxButtons.OK);
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
                catch
                {
                    MessageBox.Show("Master Folder isn't selected");
                }
            }
        }

        private void MoveLeftBtn_Click(object sender, EventArgs e)
        {
            if (selectedTree == SelectedTree.downloads)
            {
                var fileName = downloadsTree.SelectedNode.Text;
                var fileFullDirec = (KnownFolders.Downloads.Path + @"\" + fileName);
                Console.WriteLine(fileFullDirec);
                var newSrcFileName = (CurrentSrc + @"\" + fileName);
                if (!File.Exists(newSrcFileName))
                {
                    SelectedIndexD = null;
                    Console.WriteLine("Okay");
                    File.Move(fileFullDirec, newSrcFileName);
                }
                else
                {
                    Console.WriteLine("Fail");
                    MessageBox.Show("Cant move, File already exists in source folder", 
                        "Error", MessageBoxButtons.OK);
                }
            }
        }

        private void MoveRightBtn_Click(object sender, EventArgs e)
        {
            if (selectedTree == SelectedTree.source)
            {
                var fileName = sourceTree.SelectedNode.Text;
                var fileFullDirec = (CurrentSrc + @"\" + fileName);
                Console.WriteLine(fileFullDirec);
                var newDownFileName = (KnownFolders.Downloads.Path + @"\" + fileName);
                if (!File.Exists(newDownFileName))
                {
                    SelectedIndexD = null;
                    Console.WriteLine("Okay");
                    File.Move(fileFullDirec, newDownFileName);
                }
                else
                {
                    Console.WriteLine("Fail");
                    MessageBox.Show("Cant move, File already existsin downloads folder",
                        "Error", MessageBoxButtons.OK);
                }
            }
        }

        private void ChangedTreeSelected(object sender, TreeNodeMouseClickEventArgs e)
        {
            var obj = (TreeView)sender;
            if (e.Node != null)
                if(obj.Name == "sourceTree")
                {
                    selectedTree = SelectedTree.source;
                    SelectedIndexS = e.Node.Index;
                    Console.WriteLine("Source tree > index: " + e.Node.Index);
                }
                else if (obj.Name == "downloadsTree")
                {
                    selectedTree = SelectedTree.downloads;
                    SelectedIndexD = e.Node.Index;
                    Console.WriteLine("Download tree > index: " + e.Node.Index);
                }
            else
                Console.WriteLine("No node selected");
            Console.WriteLine(" ");
        }
       
        private void Start_FormClosing(object sender, FormClosingEventArgs e)
        {
            SaveLocation();
        }

        private delegate void updatedelegate();
        private updatedelegate del;
        private string PreviousAcessedFile = null;
        private string PreviousReason = null;
        private void WatcherChanged(object sender, FileSystemEventArgs e, string obj, string arg)
        {
            Console.WriteLine(obj + ", Reason raised: " + arg + ", e: " + e.Name);
            if (e.Name == PreviousAcessedFile && PreviousReason == "Deleted")
                PreviousAcessedFile = null;
            else
                PreviousAcessedFile = e.Name;
            PreviousReason = arg;
            if (obj == "Source" && e.Name != "src.pdf" && new[] { "Deleted", "Created", "Renamed" }.Contains(arg))
            {
                if (!WaitStarted)
                {
                    ThreadStart Job = new ThreadStart(Waiter);
                    Thread thread = new Thread(Job);
                    Wait = true;
                    thread.Start();
                    Thread.Sleep(100);
                    del = this.UpdateUI;
                    this.Invoke(del, null);
                    del = this.AggregatePdfs;
                    this.Invoke(del, null);
                }

            }
            else
            {
                del = this.UpdateUI;
                this.Invoke(del, null);
            }
        }

    }


    public enum SelectedTree
    {
        downloads,
        source
        
    }
}
