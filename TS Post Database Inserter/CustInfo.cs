using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using System.Configuration;
using System.Windows.Forms;
using System.Reflection;
using iTextSharp.text.pdf.parser;
using iTextSharp.text.pdf;
using iTextSharp.text;
using NPOI.XSSF.UserModel;
using NPOI.XSSF.Model;
using NPOI.SS.UserModel;
using NPOI.SS.Util;

namespace TS_Post_Database_Inserter
{
    public partial class CustInfo : Form
    {
        Configuration Config = ConfigurationManager.OpenExeConfiguration(Application.ExecutablePath);


        string OpenPDF = "";

        int MaxPg;
        int CurrentPg;

        //g = currentpage - 1
        public int g = 0;

        Start st;
        PdfReader reader;
        List<Pages> ResPages = new List<Pages>();
        List<CheckBox> CB = new List<CheckBox>();
        
        public FileStream fs;

        public CustInfo(Start f)
        {
            InitializeComponent();
            st = f;

            foreach (Control C in tabControl1.SelectedTab.Controls)
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

                p.DeliveryDate = tArr[x];
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

                p.ConsignmentNumber = tArr[x];
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

                p.Telephone = tArr[x];
            }

            ////////Location
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                x = tArr.Length - 2;
                p.Location = tArr[x];
            }

            ////////Location Number
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                x = tArr.Length - 1;
                p.LocationNumber = tArr[x];
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
                p.ParcelNumber = tArr[x];
            }

            InfoUpdate(ChangePage.Start);
        }        

        public void InfoUpdate(ChangePage n)
        {

            //start
            if(n == ChangePage.Start)
            {
                CurrentPg++;
                g = CurrentPg - 1;

                foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                    ResPages[g].CheckStates.Add(c.CheckState);
                
                ResPages[g].IsViewed = true;               

                NameTB.Text = ResPages[g].Name;
                tabControl1.SelectedTab.Text = (NameTB.Text + ", PDF page:" + CurrentPg);
                PgNumL.Text = ("Page: " + CurrentPg);
                AddressTB.Text = ResPages[g].Address;
                BarTB.Text = ResPages[g].Barcode;
                DelTB.Text = ResPages[g].DeliveryDate;
                ConTB.Text = ResPages[g].ConsignmentNumber;
                PostTB.Text = ResPages[g].PostCode;
                TelTB.Text = ResPages[g].Telephone;
                LocatTB.Text = ResPages[g].Location;
                LocatNoTB.Text = ResPages[g].LocationNumber;
                ParcelTB.Text = ResPages[g].ParcelNumber;

                if (CurrentPg == MaxPg)
                    Continue.Text = "Finish";

                return;
            }
            
            //Change page to next page
            if (n == ChangePage.Next)
            {

                //foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                int i = 0;
                foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                {
                    ResPages[g].CheckStates[i] = c.CheckState;
                    i++;
                }

                ResPages[g].IsViewed = true;

                if (CurrentPg < MaxPg)
                {
                    CurrentPg++;
                    g++;

                    if (ResPages[g].IsViewed)
                    {
                        int u = 0;
                        foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                        {
                            c.CheckState = ResPages[g].CheckStates[u];
                            u++;
                        }
                    }
                    else if(!ResPages[g].IsViewed)
                    {
                        foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                        {
                            c.CheckState = CheckState.Unchecked;
                            ResPages[g].CheckStates.Add(c.CheckState);
                        }
                    }

                    //textfields
                    NameTB.Text = ResPages[g].Name;
                    tabControl1.SelectedTab.Text = (NameTB.Text + ", PDF page:" + CurrentPg);
                    PgNumL.Text = ("Page: " + CurrentPg);
                    AddressTB.Text = ResPages[g].Address;
                    BarTB.Text = ResPages[g].Barcode;
                    DelTB.Text = ResPages[g].DeliveryDate;
                    ConTB.Text = ResPages[g].ConsignmentNumber;
                    PostTB.Text = ResPages[g].PostCode;
                    TelTB.Text = ResPages[g].Telephone;
                    LocatTB.Text = ResPages[g].Location;
                    LocatNoTB.Text = ResPages[g].LocationNumber;
                    ParcelTB.Text = ResPages[g].ParcelNumber;

                    //end
                    if(CurrentPg > 1)
                    {
                        PrevBtn.Enabled = true;
                    }

                    if (CurrentPg == MaxPg)
                        Continue.Text = "Finish";
                }
                else if (CurrentPg == MaxPg)
                {
                    List<CheckState> s = new List<CheckState>();
                    int w = 0;
                    int q = ResPages.Count();
                    for (int t = 0; t < q; t++) 
                        foreach (CheckState c in ResPages[t].CheckStates)                        
                            if (c == CheckState.Unchecked)
                                w++;
                    
                    if(w > 0)
                    {
                        CHKINFO CI = new CHKINFO();
                        CI.ShowDialog();
                    }
                    else if(w==0)
                    {
                        PushExcel();
                    }
                }
            }

            //Change page to previous page
            if (n == ChangePage.Previous)
            {
                int i = 0;
                foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                {
                    ResPages[g].CheckStates[i] = c.CheckState;
                    i++;
                }

                ResPages[g].IsViewed = true;

                CurrentPg--;
                g = CurrentPg - 1;

                /*for (int i = 0; i < tabControl1.SelectedTab.Controls.OfType<CheckBox>().Count(); i++)
                {
                    Console.WriteLine(i);
                    foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                        c.CheckState = ResPages[g].CheckStates[i];
                }*/

                if (ResPages[g].IsViewed)
                {
                    int u = 0;
                    foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                    {
                        c.CheckState = ResPages[g].CheckStates[u];
                        u++;
                    }
                }
                else if (!ResPages[g].IsViewed)
                {
                    foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                        ResPages[g].CheckStates.Add(c.CheckState);
                }

                if (CurrentPg + 1 > 1)
                {
                    if (CurrentPg + 1 == MaxPg)
                        Continue.Text = "Next";

                    NameTB.Text = ResPages[g].Name;
                    tabControl1.SelectedTab.Text = (NameTB.Text + ", PDF page:" + CurrentPg);
                    PgNumL.Text = ("Page: " + CurrentPg);
                    AddressTB.Text = ResPages[g].Address;
                    BarTB.Text = ResPages[g].Barcode;
                    DelTB.Text = ResPages[g].DeliveryDate;
                    ConTB.Text = ResPages[g].ConsignmentNumber;
                    PostTB.Text = ResPages[g].PostCode;
                    TelTB.Text = ResPages[g].Telephone;
                    LocatTB.Text = ResPages[g].Location;
                    LocatNoTB.Text = ResPages[g].LocationNumber;
                    ParcelTB.Text = ResPages[g].ParcelNumber;

                    if (CurrentPg == 1)
                    {
                        PrevBtn.Enabled = false;
                    }

                }
            }

        }

        private void Continue_Click(object sender, EventArgs e)
        {
            InfoUpdate(ChangePage.Next);
        }

        private void PrevBtn_Click(object sender, EventArgs e)
        {
            if(CurrentPg > 1)
            {
                InfoUpdate(ChangePage.Previous);
            }
        }
        
        private void Cancel_Click(object sender, EventArgs e)
        {
            this.Close();
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

            int y = 0;
            foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
            {/*
                if (c == t)
                    ResPages[g]
                y++;*/
            }

        }


        public XSSFWorkbook WB;
        public XSSFSheet WS;

        void PushExcel()
        {
            try
            {
                fs = new FileStream(st.MainExcel, FileMode.Open, FileAccess.ReadWrite);
            }
            catch (Exception)
            {
                // show error dialoge
                throw;
            }

            WB = new XSSFWorkbook(fs);
            WS = WB.GetSheetAt(0) as XSSFSheet;

            //Console.WriteLine(WS.PhysicalNumberOfRows);
            //Console.WriteLine(WS.column);


            /*
int cR = WS.PhysicalNumberOfRows;
int cC = WS.GetRow(WS.FirstRowNum).LastCellNum;
int aR = 0;

//int aC = 1;
for (int i = 0 ; i < ResPages.Count; i++)
{
    aR = cR + i + 1;
    Console.WriteLine("i = " + i);
    Console.WriteLine("ar = " + aR);
    Console.WriteLine("cr = " + cR);

    foreach (Pages p in ResPages)
    {
        List<string> vs = new List<string>();
        List<string> vp = new List<string>();
        List<PropertyInfo> pi = new List<PropertyInfo>();
        foreach (var prop in p.GetType().GetProperties())
        {
            if (prop.PropertyType == typeof(string) && prop.Name != "PDFtext")
            {
                vs.Add(prop.GetValue(p, null).ToString());
                vp.Add(prop.Name.ToString());
                pi.Add(prop);
            }
        }
        foreach (PropertyInfo S in pi)                        
        {
            for(int y = 0; y < cC; y++)
            {   
                if (S.Name == WS.GetRow(WS.FirstRowNum).GetCell(y).StringCellValue)
                {
                    Console.WriteLine("TOP: " + WS.GetRow(WS.FirstRowNum).GetCell(y).StringCellValue + ", aR: " + aR + ", Y: " + y);
                    Console.WriteLine(S.GetValue(p, null).ToString());
                    WS.CreateRow(aR);
                    WS.GetRow(aR).CreateCell(y);
                    WS.GetRow(aR).GetCell(y).SetCellValue(S.GetValue(p, null).ToString());
                    Console.WriteLine("Check: " + WS.GetRow(aR).GetCell(y).StringCellValue);
                    Console.WriteLine("");
                }
            }
        }
    }
}   */



            int CountRow = WS.PhysicalNumberOfRows;
            int MaxColumns = WS.GetRow(0).LastCellNum;

            int i = 1;
            int NRow;
            foreach(Pages p in ResPages)
            {
                NRow = i + CountRow;
                List<PropertyInfo> Pi = new List<PropertyInfo>();
                if (i <+ ResPages.Count)
                {
                    foreach(var prop in p.GetType().GetProperties())
                    {
                        if(prop.GetType() == typeof(string) && prop.Name != "PDFtext")
                        {
                            Pi.Add(prop);
                        }
                    }
                }
                int c = 0;
                foreach(PropertyInfo S in Pi)
                {
                    if(WS.GetRow(WS.FirstRowNum).GetCell(c).StringCellValue == S.Name)
                    {
                        WS.CreateRow(NRow);
                        WS.GetRow(NRow).CreateCell(c);
                        WS.GetRow(NRow).GetCell(c).SetCellValue(S.GetValue(Pi, null).ToString());
                    }
                    c++;
                }
                i++;
            }

            using (var rs = new FileStream(st.MainExcel, FileMode.Create, FileAccess.Write))
            {               
                WB.Write(rs);
                rs.Close();
            }
        }

        private void CustInfo_FormClosing(object sender, FormClosingEventArgs e)
        {
            //empty
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
        Previous,
    }
}
