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
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using iTextSharp.text;
using Excel = Microsoft.Office.Interop.Excel;
using MVSTA = Microsoft.VisualStudio.Tools.Applications;

namespace TS_Post_Database_Inserter
{
    public partial class CustInfo : Form
    {
        Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);

        string OpenPDF = "";

        int MaxPg;
        int CurrentPg;

        PdfReader reader;
        List<Pages> ResPages = new List<Pages>();
        List<CheckBox> CB = new List<CheckBox>();
        public CustInfo()
        {
            InitializeComponent();

            foreach(Control C in tabControl1.SelectedTab.Controls)
            { 
                if(C is CheckBox)
                    CB.Add(C as CheckBox);
            }
            foreach(CheckBox c in CB)
            {
                c.CheckStateChanged += C_CheckStateChanged;
                if (c.CheckState == CheckState.Unchecked)
                    c.BackColor = Color.Red;
            }

            Console.WriteLine("count:  = " + CB.Count);

            OpenPDF = Config.AppSettings.Settings["OpenPDF"].Value;

            reader = new PdfReader(OpenPDF);
            MaxPg = reader.NumberOfPages;
            CurrentPg = 0;

            ///////Init
            for (int i = 1; i <= reader.NumberOfPages; i++)
            {
                Pages tempPG = new Pages();
                string temp = PdfTextExtractor.GetTextFromPage(reader, i, new SimpleTextExtractionStrategy());
                tempPG.PDFtext = temp;
                tempPG.ResultArr = temp.Split(new string[] { "\n" }, StringSplitOptions.RemoveEmptyEntries);
                ResPages.Add(tempPG);
            }

            int v = -1;
            foreach (string o in ResPages[CurrentPg].ResultArr)
            {
                v++;
                Console.WriteLine("Line" + v +": " + o);
            }

            //////Name
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].Contains("___"))
                        x = i + 2;
                }

                p.Name = tArr[x];
                Console.WriteLine("name = " + p.Name);
            }

            //////Address
            foreach (Pages p in ResPages)
            {
                int x = 0;
                int y = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].Contains("___"))
                        x = i+3;
                    if (tArr[i].IndexOf("Next Day") > -1)
                    {
                        y = i-2;
                    }
                }
                
                List<string> tempAdArry = new List<string>();
                for(int i = x; i <= y; i++)
                {
                    tempAdArry.Add(tArr[i]);
                }
                p.Address = string.Join(",\n", tempAdArry.ToArray());
                Console.WriteLine("address = " + p.Address);
            }

            ////////Barcode
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].Contains("___"))
                        x = i + 1;
                }

                p.Barcode = tArr[x];
                Console.WriteLine("Barcode = " + p.Barcode);
            }

            ////////Delivery Date
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].IndexOf("Next Day") > -1)
                    {
                        x = i - 1;
                    }
                }

                p.DelDate = tArr[x];
                Console.WriteLine("Del Date = " + p.DelDate);
            }

            ////////Consignment Number
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].IndexOf("Next Day") > -1)
                    {
                        x = i + 1;
                    }
                }

                p.ConNumb = tArr[x];
                Console.WriteLine("Consignment Number = " + p.ConNumb);
            }

            ////////PostCode
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].IndexOf("Next Day") > -1)
                    {
                        x = i + 3;
                    }
                }

                p.PostCode = tArr[x];
                Console.WriteLine("PostCode = " + p.PostCode);
            }

            ////////Tel
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].IndexOf("Next Day") > -1)
                    {
                        x = i + 5;
                    }
                }

                p.Tel = tArr[x];
                Console.WriteLine("Telephone = " + p.Tel);
            }

            ////////Location
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                x = tArr.Length - 2;
                p.Locat = tArr[x];
                Console.WriteLine("Location = " + p.Locat);
            }

            ////////Location Number
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                x = tArr.Length - 1;
                p.LocatNo = tArr[x];
                Console.WriteLine("Location Number = " + p.LocatNo);
            }

            ////////Parcle Number
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                for (int i = 0; i < tArr.Length; i++)
                {
                    if (tArr[i].IndexOf("Next Day") > -1)
                    {
                        x = i + 4;
                    }
                }
                p.ParceNum = tArr[x];
                Console.WriteLine("Parcel number = " + p.ParceNum);
            }

            InfoUpdate(true);

            Console.WriteLine("End");
        }

        private void C_CheckStateChanged(object sender, EventArgs e)
        {
            CheckBox t = null;
            if (sender is CheckBox)
            {
                t = (CheckBox)sender;
            }

            if (t.BackColor == Color.Red)
                t.BackColor = Color.Transparent;
            else if (t.BackColor == Color.Transparent)
                t.BackColor = Color.Red;
        }

        public void InfoUpdate(bool more)
        {
            int g;
            if (!more)
            {
                if (CurrentPg > 1)
                {
                    if (CurrentPg == MaxPg)
                        Continue.Text = "Next";
                    CurrentPg--;
                    Console.WriteLine(CurrentPg);
                    g = CurrentPg - 1;
                    NameTB.Text = ResPages[g].Name;
                    tabControl1.SelectedTab.Text = (NameTB.Text + ", PDF page:" + CurrentPg);
                    PgNumL.Text = ("Page: " + CurrentPg);
                    AddressTB.Text = ResPages[g].Address;
                    BarTB.Text = ResPages[g].Barcode;
                    DelTB.Text = ResPages[g].DelDate;
                    ConTB.Text = ResPages[g].ConNumb;
                    PostTB.Text = ResPages[g].PostCode;
                    TelTB.Text = ResPages[g].Tel;
                    LocatTB.Text = ResPages[g].Locat;
                    LocatNoTB.Text = ResPages[g].LocatNo;
                    ParcelTB.Text = ResPages[g].ParceNum;
                    if (CurrentPg == 1)
                    {
                        PrevBtn.Enabled = false;
                    }
                }
            }
            else
            {
                if (CurrentPg < MaxPg)
                {
                    CurrentPg++;
                    Console.WriteLine(CurrentPg);
                    g = CurrentPg - 1;
                    NameTB.Text = ResPages[g].Name;
                    tabControl1.SelectedTab.Text = (NameTB.Text + ", PDF page:" + CurrentPg);
                    PgNumL.Text = ("Page: " + CurrentPg);
                    AddressTB.Text = ResPages[g].Address;
                    BarTB.Text = ResPages[g].Barcode;
                    DelTB.Text = ResPages[g].DelDate;
                    ConTB.Text = ResPages[g].ConNumb;
                    PostTB.Text = ResPages[g].PostCode;
                    TelTB.Text = ResPages[g].Tel;
                    LocatTB.Text = ResPages[g].Locat;
                    LocatNoTB.Text = ResPages[g].LocatNo;
                    ParcelTB.Text = ResPages[g].ParceNum;
                    if(CurrentPg > 1)
                    {
                        PrevBtn.Enabled = true;
                    }
                    if (CurrentPg == MaxPg)
                        Continue.Text = "Finish";
                }
            }
        }

        public class Pages
        {
            public string PDFtext { get; set; }
            public string[] ResultArr { get; set; }
            public string Name { get; set; }
            public string Address { get; set; }
            public string Barcode { get; set; }
            public string DelDate { get; set; }
            public string ConNumb { get; set; }
            public string PostCode { get; set; }
            public string Tel { get; set; }
            public string Locat { get; set; }
            public string LocatNo { get; set; }
            public string ParceNum { get; set; }
        }

        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Continue_Click(object sender, EventArgs e)
        {
            int i = 0;
            foreach (CheckBox l in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
            {
                if (l.Checked == false)
                {
                    l.BackColor = Color.Red;
                    i++;
                }
                    
            }
            if (i > 0)
            {
                CHKINFO CI = new CHKINFO();
                CI.ShowDialog();
            }
            else
            {
                if (CurrentPg < MaxPg)
                {
                    InfoUpdate(true);
                    foreach (CheckBox l in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                    {
                        l.CheckState = CheckState.Unchecked;
                         l.BackColor = Color.Red;
                    }
                }
                else if (CurrentPg == MaxPg)
                {
                    //end
                }
            }
        }

        private void PrevBtn_Click(object sender, EventArgs e)
        {
            if(CurrentPg > 1)
            {
                InfoUpdate(false);
            }
        }
    }
}
