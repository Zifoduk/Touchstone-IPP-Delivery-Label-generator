using System;
using System.Collections.Generic;
using System.Configuration;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Windows.Forms;
using iTextSharp.text.pdf;
using iTextSharp.text.pdf.parser;
using NLog;
using NPOI.XSSF.UserModel;
using PostSharp.Aspects;

namespace TS_Post_Database_Inserter
{
    [ExceptionWrapper]
    public partial class CustInfo : Form
    {
        private readonly Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);


        private string OpenPDF = "";

        private readonly int MaxPg;
        private int CurrentPg;

        //g = currentpage - 1
        public int g;

        private readonly Start start;
        private readonly PdfReader reader;
        private readonly List<Pages> ResPages = new List<Pages>();
        private readonly List<CheckBox> listCheckBoxes = new List<CheckBox>();

        public FileStream Filestream;



        public CustInfo(Start f)
        {
            InitializeComponent();

            start = f;

            foreach (Control control in tabControl1.SelectedTab.Controls)
                if (control is CheckBox)
                    listCheckBoxes.Add(control as CheckBox);
            foreach (var c in listCheckBoxes)
            {
                c.CheckStateChanged += C_CheckStateChanged;
                if (c.CheckState == CheckState.Unchecked)
                    c.BackColor = Color.Red;
            }

            OpenPDF = Config.AppSettings.Settings["OpenPDF"].Value;


            reader = new PdfReader(start.OpenPDF);
            MaxPg = reader.NumberOfPages;
            CurrentPg = 0;

            ///////Init
            for (var i = 1; i <= reader.NumberOfPages; i++)
            {
                var tempPG = new Pages();
                var temp = PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
                tempPG.PDFtext                              = temp;
                tempPG.ResultArr                            = temp.Split(new[] {"\n"}, StringSplitOptions.RemoveEmptyEntries);
                ResPages.Add(tempPG);
            }

            var v                                           = -1;
            foreach (var o in ResPages[CurrentPg].ResultArr) v++;

            //////Name
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].Contains("___"))
                        x                                   = i + 2;

                pages.Name                                  = tArr[x];
            }

            //////Address
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var y                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].Contains("___"))
                        x                                   = i + 3;
                    if (tArr[i].IndexOf("Next Day") > -1)
                        y                                   = i - 2;
                }

                var tempAdArry = new List<string>();
                for (var i = x; i <= y; i++) tempAdArry.Add(tArr[i]);
                pages.Address = string.Join(",\r\n", tempAdArry.ToArray());
            }

            ////////Barcode
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].Contains("___"))
                        x                                   = i + 1;

                pages.Barcode                               = tArr[x];
            }

            ////////Delivery Date
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].IndexOf("Next Day") > -1)
                        x                                   = i - 1;

                pages.DeliveryDate                          = tArr[x];
            }

            ////////Consignment Number
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].IndexOf("Next Day") > -1)
                        x                                   = i + 1;

                pages.ConsignmentNumber                     = tArr[x];
            }

            ////////PostCode
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].IndexOf("Next Day") > -1)
                        x                                   = i + 3;

                pages.PostCode                              = tArr[x];
            }

            ////////Tel
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].IndexOf("Next Day") > -1)
                        x                                   = i + 5;

                pages.Telephone                             = tArr[x];
            }

            ////////Location
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                x                                           = tArr.Length - 2;
                pages.Location                              = tArr[x];
            }

            ////////Location Number
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                x                                           = tArr.Length - 1;
                pages.LocationNumber                        = tArr[x];
            }

            ////////Parcel Number
            foreach (var pages in ResPages)
            {
                var x                                       = 0;
                var tArr                                    = pages.ResultArr;
                for (var i                                  = 0; i < tArr.Length; i++)
                    if (tArr[i].IndexOf("Next Day") > -1)
                        x                                   = i + 4;
                pages.ParcelNumber                          = tArr[x];
            }

            InfoUpdate(ChangePage.Start);
        }

        public void InfoUpdate(ChangePage changePage)
        {
            //start
            if (changePage == ChangePage.Start)
            {
                CurrentPg++;
                g = CurrentPg - 1;

                foreach (var c in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                    ResPages[g].CheckStates.Add(c.CheckState);

                ResPages[g].IsViewed         = true;

                NameTB.Text                  = ResPages[g].Name;
                tabControl1.SelectedTab.Text = NameTB.Text + ", PDF page:" + CurrentPg;
                PgNumL.Text                  = "Page: " + CurrentPg;
                this.Text                    = NameTB.Text + ", PDF page:" + CurrentPg;
                AddressTB.Text               = ResPages[g].Address;
                BarTB.Text                   = ResPages[g].Barcode;
                DelTB.Text                   = ResPages[g].DeliveryDate;
                ConTB.Text                   = ResPages[g].ConsignmentNumber;
                PostTB.Text                  = ResPages[g].PostCode;
                TelTB.Text                   = ResPages[g].Telephone;
                LocatTB.Text                 = ResPages[g].Location;
                LocatNoTB.Text               = ResPages[g].LocationNumber;
                ParcelTB.Text                = ResPages[g].ParcelNumber;

                if (CurrentPg == MaxPg)
                    Continue.Text = "Finish";

                return;
            }

            //Change page to next page
            if (changePage == ChangePage.Next)
            {
                //foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                var idex = 0;
                foreach (var c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                {
                    ResPages[g].CheckStates[idex] = c.CheckState;
                    idex++;
                }

                ResPages[g].IsViewed              = true;

                if (CurrentPg < MaxPg)
                {
                    CurrentPg++;
                    g++;

                    if (ResPages[g].IsViewed)
                    {
                        var index                 = 0;
                        foreach (var checkBox in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                        {
                            checkBox.CheckState   = ResPages[g].CheckStates[index];
                            index++;
                        }
                    }
                    else if (!ResPages[g].IsViewed)
                    {
                        foreach (var checkBox in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                        {
                            checkBox.CheckState   = CheckState.Unchecked;
                            ResPages[g].CheckStates.Add(checkBox.CheckState);
                        }
                    }

                    //textfields
                    NameTB.Text                   = ResPages[g].Name;
                    tabControl1.SelectedTab.Text  = NameTB.Text + ", PDF page:" + CurrentPg;
                    PgNumL.Text                   = "Page: " + CurrentPg;
                    this.Text                     = NameTB.Text + ", PDF page:" + CurrentPg;
                    AddressTB.Text                = ResPages[g].Address;
                    BarTB.Text                    = ResPages[g].Barcode;
                    DelTB.Text                    = ResPages[g].DeliveryDate;
                    ConTB.Text                    = ResPages[g].ConsignmentNumber;
                    PostTB.Text                   = ResPages[g].PostCode;
                    TelTB.Text                    = ResPages[g].Telephone;
                    LocatTB.Text                  = ResPages[g].Location;
                    LocatNoTB.Text                = ResPages[g].LocationNumber;
                    ParcelTB.Text                 = ResPages[g].ParcelNumber;

                    //end
                    if (CurrentPg > 1) PrevBtn.Enabled = true;

                    if (CurrentPg == MaxPg)
                        Continue.Text = "Finish";
                }
                else if (CurrentPg == MaxPg)
                {
                    var checkStates          = new List<CheckState>();
                    var uncheckedNumber      = 0;
                    var pageCount            = ResPages.Count();
                    for (var index = 0; index < pageCount; index++)
                        foreach (var checkState in ResPages[index].CheckStates)
                            if (checkState == CheckState.Unchecked)
                                uncheckedNumber++;

                    if (uncheckedNumber > 0)
                    {
                        var chkinfo          = new CHKINFO();
                        chkinfo.ShowDialog();
                    }
                    else if (uncheckedNumber == 0)
                    {
                        var completed        = new Completed(this);
                        completed.ShowDialog();
                        Close();
                    }
                }
            }

            //Change page to previous page
            if (changePage == ChangePage.Previous)
            {
                var index                          = 0;
                foreach (var checkBox in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                {
                    ResPages[g].CheckStates[index] = checkBox.CheckState;
                    index++;
                }

                ResPages[g].IsViewed               = true;

                CurrentPg--;
                g                                  = CurrentPg - 1;

                if (ResPages[g].IsViewed)
                {
                    var i                          = 0;
                    foreach (var c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                    {
                        c.CheckState               = ResPages[g].CheckStates[i];
                        i++;
                    }
                }
                else if (!ResPages[g].IsViewed)
                    foreach (var c in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                        ResPages[g].CheckStates.Add(c.CheckState);

                if (CurrentPg + 1 > 1)
                {
                    if (CurrentPg + 1              == MaxPg)
                        Continue.Text              = "Next";

                    NameTB.Text                    = ResPages[g].Name;
                    tabControl1.SelectedTab.Text   = NameTB.Text + ", PDF page:" + CurrentPg;
                    PgNumL.Text                    = "Page: " + CurrentPg;
                    this.Text                      = NameTB.Text + ", PDF page:" + CurrentPg;
                    AddressTB.Text                 = ResPages[g].Address;
                    BarTB.Text                     = ResPages[g].Barcode;
                    DelTB.Text                     = ResPages[g].DeliveryDate;
                    ConTB.Text                     = ResPages[g].ConsignmentNumber;
                    PostTB.Text                    = ResPages[g].PostCode;
                    TelTB.Text                     = ResPages[g].Telephone;
                    LocatTB.Text                   = ResPages[g].Location;
                    LocatNoTB.Text                 = ResPages[g].LocationNumber;
                    ParcelTB.Text                  = ResPages[g].ParcelNumber;

                    if (CurrentPg                  == 1) PrevBtn.Enabled = false;
                }
            }
        }


        //events

        private void Continue_Click(object sender, EventArgs e)
        {
            InfoUpdate(ChangePage.Next);
        }

        private void PrevBtn_Click(object sender, EventArgs e)
        {
            if (CurrentPg > 1) InfoUpdate(ChangePage.Previous);
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void C_CheckStateChanged(object sender, EventArgs e)
        {
            CheckBox checkBox                      = null;
            if (sender is CheckBox) checkBox       = (CheckBox) sender;

            if (checkBox.BackColor == Color.Red)
                checkBox.BackColor                 = Color.Transparent;
            else if (checkBox.BackColor == Color.Transparent)
                checkBox.BackColor                 = Color.Red;
        }


        //excel parts

        public XSSFWorkbook Workbook;
        public XSSFSheet WorkSheet;

        public void PushExcel(Completed completed)
        {
            //write to excel
            completed.Progressbar.PerformStep();
            try
            {
                Filestream = new FileStream(start.MainExcel, FileMode.Open, FileAccess.ReadWrite);
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

            var CountRow = WorkSheet.PhysicalNumberOfRows;
            Console.WriteLine("CR: " + CountRow);
            var i = 0;
            foreach (var pages in ResPages)
            {
                float Calculate = (160/ (float)ResPages.Count);
                int math = (int)Math.Round(Calculate, 0, MidpointRounding.AwayFromZero);
                completed.Progressbar.Increment(math);

                var NRow = i + CountRow;
                WorkSheet.CreateRow(NRow);
                var propertyInfos = new List<PropertyInfo>();
                Console.WriteLine(NRow);
                if (i < ResPages.Count)
                    propertyInfos.AddRange(pages.GetType().GetProperties().Where(prop => prop.PropertyType == typeof(string) && prop.Name != "PDFtext"));
                
                var cellnum = 0;
                foreach (var propertyInfo in propertyInfos)
                {
                    if (WorkSheet.GetRow(WorkSheet.FirstRowNum).GetCell(cellnum).StringCellValue == propertyInfo.Name)
                    {
                        WorkSheet.GetRow(NRow).CreateCell(cellnum);
                        WorkSheet.GetRow(NRow).GetCell(cellnum).SetCellValue(propertyInfo.GetValue(pages, null).ToString());
                    }
                    cellnum++;
                }

                i++;
            }


            completed.Progressbar.PerformStep();
            using (var fileStream = new FileStream(start.MainExcel, FileMode.Create, FileAccess.Write))
            {
                try
                {
                    Workbook.Write(fileStream);
                    fileStream.Close();
                    completed.Progressbar.Increment(60);
                }
                catch (Exception ee)
                {
                    completed.Progressbar.Increment(-60);
                    completed.Progressbar.ForeColor = Color.Red;
                    throw ee;
                }
            }
        }

        protected void AppendExcel()
        {
            //if(start.)
        }



        //empty

        private void CustInfo_FormClosing(object sender, FormClosingEventArgs e)
        {
            //empty
        }
    }

    [Serializable]
    public class ExceptionWrapper : OnExceptionAspect
    {
        public override void OnException(MethodExecutionArgs args)
        {
            var ex = args.Exception;
            base.OnException(args);
            Console.WriteLine(ex);
            var logger = LogManager.GetCurrentClassLogger();
            logger.ErrorException(ex.ToString(), ex);
            Console.WriteLine("check code");
        }
    }

    public class ExcelDocumentOpenException : Exception
    {
        public ExcelDocumentOpenException()
            : base("Main Excel is being used by another application")
        {
        }

        public ExcelDocumentOpenException(Exception inner)
            : base("Main Excel is being used by another application", inner)
        {
        }
    }


    public class Pages
    {
        public string PDFtext { get; set; }
        public string[] ResultArr { get; set; }
        public string Name { get; set; }
        public string Address { get; set; }
        public string Barcode { get; set; }
        public string DeliveryDate { get; set; }
        public string ConsignmentNumber { get; set; }
        public string PostCode { get; set; }
        public string Telephone { get; set; }
        public string Location { get; set; }
        public string LocationNumber { get; set; }
        public string ParcelNumber { get; set; }
        public List<CheckState> CheckStates = new List<CheckState>();
        public bool IsViewed { get; set; }
    }


    public enum ChangePage
    {
        Start,
        Next,
        Previous
    }
}
 