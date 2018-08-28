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
using System.Runtime.InteropServices;

namespace TS_Post_Database_Inserter
{

    [ExceptionWrapper]
    public partial class Start : Form
    {
        private readonly Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

        public static string CheckDir(string dir)
        {

            if(!Directory.Exists(dir))
            {
                Directory.CreateDirectory(dir);
            }
            return dir;
        }
        public static Guid AppGuid
        { get { Assembly asm = Assembly.GetEntryAssembly();
                object[] attr = (asm.GetCustomAttributes(typeof(GuidAttribute), true));
                    return new Guid((attr[0] as GuidAttribute).Value);
            }
        }
        public static string UserRoamingDataFolder
        {
            get {
                Guid appGuid = AppGuid;
                string folderBase = Environment.GetFolderPath(Environment.SpecialFolder.ApplicationData);
                //string dir = string.Format(@"{0}\{1}\", folderBase, appGuid.ToString("B").ToUpper());
                string dir = string.Format(@"{0}\{1}\Data\", folderBase, Application.ProductName);
                return CheckDir(dir);
            }
        }
               
        public string RootFolder;
        public static string Tempfolder
        {
            get
            {
                string temp = UserRoamingDataFolder + @"temp\";
                if (!Directory.Exists(temp))
                {
                    Directory.CreateDirectory(temp);
                }
                return temp;
            }
        }
        public static string SourceFolder
        {
            get
            {
                string temp = UserRoamingDataFolder + @"Source PDFs\";
                if (!Directory.Exists(temp))
                {
                    Directory.CreateDirectory(temp);
                }
                return temp;
            }
        }
        public static string Archives
        {
            get
            {
                string temp = UserRoamingDataFolder + @"Archives\";
                if (!Directory.Exists(temp))
                {
                    Directory.CreateDirectory(temp);
                }
                return temp;
            }
        }
        public static string MasterFolder
        {
            get
            {
                string temp = UserRoamingDataFolder + @"Master\";
                if (!Directory.Exists(temp))
                {
                    Directory.CreateDirectory(temp);
                }
                return temp;
            }
        }

        public static void CheckFile(string file)
        {
            Int64 Checksize = 0;
            string filename = System.IO.Path.GetFileName(file);
            Console.WriteLine(filename);

            if (File.Exists(file))
            {
                Checksize = new FileInfo(file).Length;
                if (Checksize < 10)
                {
                    if (File.Exists(file))
                        File.Delete(file);
                    string resourceName = string.Join("", Assembly.GetExecutingAssembly()
                                        .GetManifestResourceNames().Where(Resource => Resource.Contains(filename)));
                    using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                    {
                        byte[] buffer = null;
                        using (BinaryReader binaryReader = new BinaryReader(stream))
                        {
                            buffer = binaryReader.ReadBytes((int)stream.Length);
                        }
                        File.WriteAllBytes(file, buffer);
                    }
                }
            }
            else
            {
                string resourceName = string.Join("", Assembly.GetExecutingAssembly()
                    .GetManifestResourceNames().Where(Resource => Resource.Contains(filename)));
                using (Stream stream = Assembly.GetExecutingAssembly().GetManifestResourceStream(resourceName))
                {
                    byte[] buffer = null;
                    using (BinaryReader binaryReader = new BinaryReader(stream))
                    {
                        buffer = binaryReader.ReadBytes((int)stream.Length);
                    }
                    File.WriteAllBytes(file, buffer);
                }
            }


        }
        public static string MainExcelFile
        {
            get
            {
                string temp = SourceFolder + @"Main.xlsx";
                if (!File.Exists(temp))
                {
                    CheckFile(temp);
                }
                return temp;
            }
        }
        public string _MainExcelFile = MainExcelFile;
        public static string MasterExcelFile
        {
            get
            {
                string temp = MasterFolder + @"Master.xlsx";
                if (!File.Exists(temp))
                {
                    CheckFile(temp);
                }
                return temp;
            }
        }
        public string _MasterExcelFile = MasterExcelFile;
        public static string SourcePdfFile
        {
            get
            {
                string temp = SourceFolder + @"src.pdf";
                /*if (!File.Exists(temp))
                {
                    CheckFile(temp);
                }*/
                return temp;
            }
        }
        public string _SourcePDFFile = SourcePdfFile;
        private static string LabelLBXFile
        {
            get
            {
                string temp = UserRoamingDataFolder + @"Label.lbx";
                CheckFile(temp);
                return temp;
            }
        }



        private int NumberofPDFs = 0;

        public SelectedTree? selectedTree = null;

        private PdfReader reader;

        public XSSFWorkbook masterWorkbook;
        public XSSFSheet masterSheet;
        public FileStream masterFileStream;
        
        public Completed _completed;

        private List<string> sourceFilesList = new List<string>();
        private List<string> downloadsFilesList = new List<string>();
        private int? SelectedIndexD = null;
        private int? SelectedIndexS = null;

        string[] srcPDFFiles = null;


        //Main Operations

        public Start()
        {
            InitializeComponent();
            
            WaitLabel.Enabled = false;
            WaitLabel.Visible = false;

            //File Settings location
            RootFolder = UserRoamingDataFolder;

            ArchiveExcel();
            UpdateUI();
            SaveLocation();

            Config.AppSettings.Settings["OpenPDF"].Value = SourcePdfFile;

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
            Console.WriteLine("Update UI raised >> " + DateTime.Now.ToShortTimeString());


            PdfReader.unethicalreading = true;

            //Downloads Tree
            var downloadfiles = Directory.GetFiles(KnownFolders.Downloads.Path, "*.pdf").ToArray();
            foreach (var file in downloadfiles)
            {
                if (IsValidPdf(file))
                {
                    try
                    {
                        using (PdfReader reader = new PdfReader(file))
                        { 
                            string[] PdfContent = new string[] { };

                            for (var i = 1; i <= reader.NumberOfPages; i++)
                            {
                                var temp = PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
                                PdfContent = temp.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                            }

                            foreach (string s in PdfContent)
                                if (s.Contains("ipostparcels"))
                                    downloadsFilesList.Add(System.IO.Path.GetFileNameWithoutExtension(file));
                                else
                                    continue;
                        }
                    }
                    catch
                    {
                        throw new Exception("Error reading PDF and adding to Download Treeview");
                    }
                }
            }


            //Rest of UI
            var CheckFilesTemp = new List<string>();
            if (Directory.Exists(SourceFolder))
            {
                foreach (var files in Directory.GetFiles(SourceFolder).Where(File => File.Contains(".pdf")))
                    CheckFilesTemp.Add(files);
            }

            if (RootFolder != "" && Directory.Exists(RootFolder))
            {
                //Source Tree
                MFol.Text = RootFolder;
                MFol.ForeColor = Color.Black;
                if (Directory.Exists(SourceFolder))
                {
                    srcPDFFiles = Directory.GetFiles(SourceFolder, "*.pdf").Where(File =>
                                            !File.Contains("src.pdf")).ToArray();
                    if (srcPDFFiles != null)
                    {
                        foreach (var file in srcPDFFiles)
                        {
                            using (MemoryStream memoryStream = new MemoryStream())
                            {
                                using (FileStream fs = File.OpenRead(file))
                                {
                                    memoryStream.SetLength(fs.Length);
                                    fs.Read(memoryStream.GetBuffer(), 0, (int)fs.Length);
                                }

                                if (IsValidPdf(file))
                                {
                                    try
                                    {
                                        using (PdfReader reader = new PdfReader(memoryStream))
                                        {
                                            string[] PdfContent = new string[] { };
                                            for (var i = 1; i <= reader.NumberOfPages; i++)
                                            {
                                                var temp = PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
                                                PdfContent = temp.Split(new[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                                            }

                                            foreach (string s in PdfContent)
                                                if (s.Contains("ipostparcels"))
                                                    sourceFilesList.Add(System.IO.Path.GetFileNameWithoutExtension(file));
                                        }
                                    }
                                    catch
                                    {
                                        throw new Exception("Error reading PDF and adding to Source Treeview");
                                    }

                                }
                            }
                        }
                        
                    }
                }
                else
                {
                    sourceFilesList.Add("nothing on source folder");
                }
;
                PHExcelL.Text = MainExcelFile;
                PHExcelL.ForeColor = Color.Black;

                int iPPdfsPresent = 0;
                int srcPresent = 0;

                if (CheckFilesTemp.Count > 0)                
                    foreach (var file in CheckFilesTemp)
                    {
                        if (file.Contains("src.pdf"))
                            srcPresent++;
                        else
                            iPPdfsPresent++;
                    }                
                else
                {
                    iPPdfsPresent = 0;
                    srcPresent = 0;
                }


                if(srcPresent > 0)
                {
                    SrcPdfL.Text = SourcePdfFile;
                    SrcPdfL.ForeColor = Color.Black;
                    reader = new PdfReader(SourcePdfFile);
                    SrcLabelsPresentL.ForeColor = Color.Black;
                    SrcLabelsPresentL.Text = "Number of Labels found: " + reader.NumberOfPages;
                    reader.Dispose();
                    reader.Close();
                }
                else
                {
                    SrcPdfL.Text = "Error Src.pdf Not found";
                    SrcPdfL.ForeColor = Color.DarkRed;
                    SrcLabelsPresentL.ForeColor = Color.DarkRed;
                    SrcLabelsPresentL.Text = "Src PDF not created - Refer to Manual(Section 2.3)";
                }
                                
                if (iPPdfsPresent > 0)
                {
                    PDFNum.Text = "Number of PDFs found in folder: " + iPPdfsPresent;
                    NumberofPDFs = iPPdfsPresent;
                    PDFNum.ForeColor = Color.Black;
                }
                else
                {
                    PDFNum.Text = 
                        "No PDFs found, Please move IPostParcels pdfs into main folder below";
                    PDFNum.ForeColor = Color.DarkRed;
                    NumberofPDFs = 0;
                    if (srcPresent > 0)
                        SrcPdfL.Text = "No pdfs found in folder, Using previous src.pdf, has "
                                                + reader.NumberOfPages + " number of labels";
                    else
                        SrcPdfL.Text = "No pdfs found in folder, previous src.pdf doesnt exist";
                    SrcPdfL.ForeColor = Color.DarkRed;
                }

            }
            else
            {
                MFol.Text = "Select a Master folder!!";
                MFol.ForeColor = Color.DarkRed;
                SrcPdfL.Text = "Select a Label PDF file!!";
                SrcPdfL.ForeColor = Color.DarkRed;
                SrcLabelsPresentL.Text = "";
                PHExcelL.Text = "No main Excel found";
                PHExcelL.ForeColor = Color.DarkRed;
            }

            sourceFilesList.ForEach(item => sourceList.Items.Add(item));
            downloadsFilesList.ForEach(item => downloadsList.Items.Add(item));
            Console.WriteLine("Update UI droppped >> " + DateTime.Now.ToShortTimeString());
        }

        [PermissionSet(SecurityAction.Demand, Name = "FullTrust")]
        private void RunWatcher()
        {
          /*  FileSystemWatcher watcherDownloads = new FileSystemWatcher();
            FileSystemWatcher watcherSource = new FileSystemWatcher(); 
            watcherDownloads.Path = KnownFolders.Downloads.Path;
            watcherSource.Path = SourceFolder;

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
            watcherSource.EnableRaisingEvents = true;*/

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




        //Variable Operations
       /* public static bool CheckFiles(string folder)
        {
            get
            {
                
            }
        }
        */





        //File Operations

        public void ArchiveExcel()
        {
            if (File.Exists(MainExcelFile) && File.Exists(MasterExcelFile))
            {
                var lastWriteTime = File.GetLastWriteTime(MainExcelFile);
                var timeDifference = (TimeSpan)lastWriteTime.Subtract(DateTime.Now);
                var differenceOfTime = (int)Math.Round(Math.Abs(timeDifference.TotalMinutes));

                if (differenceOfTime > 30)
                {
                    /////Write to Master Excel for archive if main exists and is over time limit

                    FileStream tempMasterFileStream = new FileStream(MasterExcelFile, FileMode.Open
                                                            , FileAccess.ReadWrite);
                    XSSFWorkbook masterWorkbook = new XSSFWorkbook(tempMasterFileStream);
                    XSSFSheet masterSheet = masterWorkbook.GetSheetAt(0) as XSSFSheet;

                    var copy = false;
                    using (FileStream masterFileStream = new FileStream(MasterExcelFile, FileMode.Create
                                                                        , FileAccess.ReadWrite))
                    {
                        int masterMaxRow = masterSheet.PhysicalNumberOfRows;
                        int masterCurrentRow = masterMaxRow;
                        
                        using (var mainFileStream = new FileStream(MainExcelFile, FileMode.Open
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
                        }
                    }

                    using (FileStream masterFileStream = new FileStream(MasterExcelFile, FileMode.Create, FileAccess.Write))
                    {
                        masterWorkbook.Write(masterFileStream);
                    }
                                       
                    ////Archieve Main Excel and recreate
                    if (copy)
                    {
                        var newExcelArchiveFile = Archives + @"XLSX\" + DateTime.Now.ToString("yyyy_MM_dd_HH-MM-ss") +
                                                  " Archived" + System.IO.Path.GetExtension(MainExcelFile);
                        File.Move(MainExcelFile, newExcelArchiveFile);
                        File.WriteAllBytes(MainExcelFile, Resources.Main);
                    }
                    copy = false;
                }
            }
            else
            {
                try
                {
                    File.WriteAllBytes(MainExcelFile, Resources.Main);
                }
                catch
                {
                    if (RootFolder == "")
                        throw new Exception("Failed to copy Source Excel");
                }
                finally
                {
                    Console.WriteLine("Excel creating passed");
                }
            }

        }

        private void MoveSrc()
        {
            var srcPdf = SourceFolder + @"\src.pdf";
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
            /*    Console.WriteLine("Aggregate PDF raised >> " + DateTime.Now.ToShortTimeString());
                var srcPDF = CurrentSrc + @"\src.pdf";
                PdfReader.unethicalreading = true;

                if(Directory.Exists(CurrentSrc))
                {
                    try
                    {

                        using (MemoryStream memoryStream = new MemoryStream())
                        using (Document doc = new Document())
                        using (PdfCopy pdf = new PdfCopy(doc, memoryStream))
                        {
                            doc.Open();
                            if(srcPDFFiles.Length > 0)
                            {
                                foreach(var file in srcPDFFiles)
                                {
                                    if (IsValidPdf(file))
                                    {
                                        using (MemoryStream _ReadOnlyStream = new MemoryStream())
                                        {
                                            using (FileStream fileStream = new FileStream(file, FileMode.Open, FileAccess.Read, FileShare.Read))
                                            {
                                                _ReadOnlyStream.SetLength(fileStream.Length);
                                                fileStream.Read(_ReadOnlyStream.GetBuffer(), 0, (int)_ReadOnlyStream.Length);
                                                fileStream.Flush();
                                                fileStream.Close();
                                            }

                                            using (PdfReader reader = new PdfReader(_ReadOnlyStream))
                                            {
                                                for (var i = 0; i < reader.NumberOfPages; i++)
                                                {
                                                    PdfImportedPage page = pdf.GetImportedPage(reader, i + 1);
                                                    pdf.AddPage(page);
                                                }
                                                pdf.FreeReader(reader);
                                                reader.Close();
                                            }

                                            _ReadOnlyStream.Flush();
                                            _ReadOnlyStream.Close();
                                        }
                                    }
                                }
                            }
                            pdf.Flush();
                            pdf.Close();
                            doc.Close();

                            using (var streamX = new FileStream(srcPDF, FileMode.Create, FileAccess.ReadWrite))
                            {
                                memoryStream.WriteTo(streamX);
                                streamX.Flush();
                                streamX.Close();
                            }
                            memoryStream.Flush();
                            memoryStream.Close();
                        }

                    }
                    catch
                    {
                        throw;
                    }
                }

                Wait = false;
                Console.WriteLine("Aggregare PDF dropped >> " + DateTime.Now.ToShortTimeString());*/

            Console.WriteLine("Aggregate PDF raised >> " + DateTime.Now.ToShortTimeString());


            var srcPdf = SourceFolder + @"\src.pdf";
            PdfReader.unethicalreading = true;

            var filesTemp = new List<string>();
            filesTemp.Clear();

            try
            {
                if (Directory.Exists(SourceFolder))
                {
                    foreach (var files in Directory.GetFiles(SourceFolder).Where(File =>
                           !File.Contains("src.pdf") && File.Contains(".pdf")))
                    {
                        filesTemp.Add(files);
                    }

                    if (filesTemp.Count > 0)
                    { 

                        if (File.Exists(srcPdf))
                            MoveSrc();

                        using (var stream = new MemoryStream())
                        {
                            using (var doc = new Document())
                            {
                                using (var pdf = new PdfCopy(doc, stream) { CloseStream = false })
                                {
                                    doc.Open();
                                    
                                    try
                                    {
                                        if (filesTemp.Count > 0)
                                            foreach (var file in filesTemp)
                                            {
                                                using (MemoryStream memoryStream = new MemoryStream())
                                                { 
                                                    using (FileStream fs = new FileStream(file, FileMode.Open, FileAccess.Read))
                                                    {
                                                        memoryStream.SetLength(fs.Length);
                                                        fs.Read(memoryStream.GetBuffer(), 0, (int)fs.Length);
                                                        fs.Flush();
                                                    }

                                                    using (PdfReader reader = new PdfReader(memoryStream))
                                                    {
                                                        for (var i = 0; i < reader.NumberOfPages; i++)
                                                        {
                                                            PdfImportedPage page = null;
                                                            page = pdf.GetImportedPage(reader, i + 1);
                                                            pdf.AddPage(page);
                                                        }
                                                        pdf.FreeReader(reader);
                                                    }
                                                    memoryStream.Flush();
                                                }
                                            }
                                    }
                                    catch (Exception ee)
                                    {
                                        throw new Exception("Not PDFs in source", ee);
                                    }
                                    pdf.Flush();
                                }
                            }

                            using (var streamX = new FileStream(srcPdf, FileMode.Create))
                            {
                                stream.WriteTo(streamX);
                                streamX.Flush();
                            }
                            stream.Flush();
                        }
                    }
                }

            }
            catch (Exception ee)
            {
                Console.WriteLine(ee);
            }
            Wait = false;
            
            Console.WriteLine("Aggregare PDF dropped >> " + DateTime.Now.ToShortTimeString());
        }

        public void ArchivePDF()
        {
            
            Console.WriteLine("Archive PDF raised >> " + DateTime.Now.ToShortTimeString());

            var FilesTemp = new List<string>();
            foreach (var Files in Directory.GetFiles(SourceFolder))
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


            Console.WriteLine("Archive PDF dropped >> " + DateTime.Now.ToShortTimeString());
        }


        //End Operations

        private void SaveLocation()
        {
            Config.AppSettings.Settings["Folder"].Value = RootFolder;
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
                System.Diagnostics.Process.Start(RootFolder + @"\Label.lbx");
                Console.WriteLine("Check if process started");
            }
            catch (Exception ee)
            {
                Console.WriteLine("Inner: " + ee.InnerException + ", Exception: " + ee);
                throw;
            }
        }

        void SaveMemoryStream(MemoryStream ms, string FileName)
        {
            FileStream outStream = File.OpenWrite(FileName);
            ms.WriteTo(outStream);
            outStream.Flush();
            outStream.Close();
        }




        //Events

        private void Launch_Click(object sender, EventArgs e)
        {
            bool FolderExists;
            bool MainExcelExists;
            bool OpenPDFExists;

            if (Directory.Exists(RootFolder))
                FolderExists = true;
            else FolderExists = false;

            if (File.Exists(MainExcelFile))
                MainExcelExists = true;
            else MainExcelExists = false;

            if (File.Exists(SourcePdfFile))
                OpenPDFExists = true;
            else OpenPDFExists = false;


            if (!FolderExists || !OpenPDFExists || !MainExcelExists)
                MessageBox.Show("Error 455: Cannot launch Edittor\r\nFix settings in RED", "Error 455", MessageBoxButtons.OK);
            else if (NumberofPDFs <= 0)
                MessageBox.Show("Error 487: No Label PDFs present in the Source Folder", "Error 487", MessageBoxButtons.OK);
            else if (FolderExists && OpenPDFExists && MainExcelExists)
            {
                LaunchMethod();
            }

            if (RootFolder == "")
            {
                MFol.Text = "Select a Master Excel document!!";
                MFol.ForeColor = Color.DarkRed;
            }

            //change
            if (MainExcelFile == "")
            {
                PHExcelL.Text = "Select a Main Excel document!!";
                PHExcelL.ForeColor = Color.DarkRed;
            }

            if (SourcePdfFile == "")
            {
                SrcPdfL.Text = "Select a Label PDF!!";
                SrcPdfL.ForeColor = Color.DarkRed;
            }
        }

        private void CloseBtn_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void OpenMDIR_Click(object sender, EventArgs e)
        {
            if (RootFolder != null || RootFolder != "")
            {
                try
                {
                    Process.Start(RootFolder);
                }
                catch
                {
                    MessageBox.Show("Master Folder isn't selected");
                }
            }
        }

        private void MoveLeftBtn_Click(object sender, EventArgs e)
        {
            foreach
            foreach(int i in downloadsList.CheckedIndices)
                downloadsList.Items[i]

            if (selectedTree == SelectedTree.downloads)
            {
                var fileName = downloadsList.SelectedNode.Text;
                var fileFullDirec = (KnownFolders.Downloads.Path + @"\" + fileName);
                Console.WriteLine(fileFullDirec);
                var newSrcFileName = (SourceFolder + @"\" + fileName);
                Console.WriteLine(newSrcFileName);
                if (!File.Exists(newSrcFileName))
                {
                    SelectedIndexD = null;
                    Console.WriteLine("Okay");
                    File.Move(fileFullDirec, newSrcFileName);
                    /*using (FileStream inStream = File.OpenRead(fileFullDirec))
                    using (MemoryStream storeStream = new MemoryStream())
                    {
                        storeStream.SetLength(inStream.Length);
                        inStream.Read(storeStream.GetBuffer(), 0, (int)inStream.Length);

                        storeStream.Flush();
                        inStream.Close();

                        SaveMemoryStream(storeStream, newSrcFileName);
                    }
                    File.Delete(fileFullDirec);*/
                }
                else
                {
                    Console.WriteLine("Fail");
                    MessageBox.Show("Cant move, File already exists in source folder", 
                        "Error", MessageBoxButtons.OK);
                }
            }
            Console.WriteLine("end barrier left");
        }

        private void MoveRightBtn_Click(object sender, EventArgs e)
        {
            /*if (selectedTree == SelectedTree.source)
            {
                var fileName = sourceTree.SelectedNode.Text;
                var fileFullDirec = (SourceFolder + @"\" + fileName);
                Console.WriteLine(fileFullDirec);
                var newDownFileName = (KnownFolders.Downloads.Path + @"\" + fileName);
                Console.WriteLine(newDownFileName);
                if (!File.Exists(newDownFileName))
                {
                    SelectedIndexD = null;
                    Console.WriteLine("Okay");
                    try
                    {
                        File.Move(fileFullDirec, newDownFileName);
                        /*using (FileStream inStream = File.OpenRead(fileFullDirec))
                        using (MemoryStream storeStream = new MemoryStream())
                        {
                            storeStream.SetLength(inStream.Length);
                            inStream.Read(storeStream.GetBuffer(), 0, (int)inStream.Length);

                            storeStream.Flush();
                            inStream.Close();

                            SaveMemoryStream(storeStream, newDownFileName);
                        }
                        File.Delete(fileFullDirec);
                    }
                    catch
                    {
                        throw;
                    }
                }
                else
                {
                    Console.WriteLine("Fail");
                    MessageBox.Show("Cant move, File already exists in downloads folder",
                        "Error", MessageBoxButtons.OK);
                }
            }
            Console.WriteLine("end barrier right");*/
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

        /*static Barrier barrier = new Barrier(2);
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
                    Console.WriteLine("waiting barrier");
                    barrier.SignalAndWait();
                    Console.WriteLine("barrier passed");
                    ThreadStart Job = new ThreadStart(Waiter);
                    Thread thread = new Thread(Job);
                    Wait = true;
                    thread.Start();
                    del = this.AggregatePdfs;
                    this.Invoke(del, null);
                    del = this.UpdateUI;
                    this.Invoke(del, null);

                }

            }
            else
            {
                del = this.UpdateUI;
                this.Invoke(del, null);
            }
        }*/

    }


    public enum SelectedTree
    {
        downloads,
        source
        
    }
}
