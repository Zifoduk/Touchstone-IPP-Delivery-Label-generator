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

        //g = currentpage - 1
        public int g = 0;



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

                p.DelDate = tArr[x];
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

                p.Tel = tArr[x];
            }

            ////////Location
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                x = tArr.Length - 2;
                p.Locat = tArr[x];
            }

            ////////Location Number
            foreach (Pages p in ResPages)
            {
                int x = 0;
                string[] tArr = p.ResultArr;
                x = tArr.Length - 1;
                p.LocatNo = tArr[x];
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
            }

            InfoUpdate(ChangePage.Start);
        }
        

        public void InfoUpdate(ChangePage n)
        {
            Console.WriteLine("NEW PAGE  ");
            Console.WriteLine("");

            //start
            if(n == ChangePage.Start)
            {
                CurrentPg++;
                g = CurrentPg - 1;

                Console.WriteLine(g);
                foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>().ToArray())
                    ResPages[g].CheckStates.Add(c.CheckState);
                
                Console.WriteLine(ResPages[g].checkboxes.Count);
                ResPages[g].viewed = true;               

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

                ResPages[g].viewed = true;

                if (CurrentPg < MaxPg)
                {
                    CurrentPg++;
                    g++;

                    if (ResPages[g].viewed)
                    {
                        int u = 0;
                        foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                        {
                            c.CheckState = ResPages[g].CheckStates[u];
                            u++;
                        }
                    }
                    else if(!ResPages[g].viewed)
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
                    DelTB.Text = ResPages[g].DelDate;
                    ConTB.Text = ResPages[g].ConNumb;
                    PostTB.Text = ResPages[g].PostCode;
                    TelTB.Text = ResPages[g].Tel;
                    LocatTB.Text = ResPages[g].Locat;
                    LocatNoTB.Text = ResPages[g].LocatNo;
                    ParcelTB.Text = ResPages[g].ParceNum;

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
                        //end
                        Console.WriteLine("end here");
                        Console.WriteLine("");
                    }
                }
            }

            //Change page to previous page
            if (n == ChangePage.Previous)
            {
                Console.WriteLine(g);
                int i = 0;
                foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                {
                    ResPages[g].CheckStates[i] = c.CheckState;
                    i++;
                }

                ResPages[g].viewed = true;

                CurrentPg--;
                g = CurrentPg - 1;

                /*for (int i = 0; i < tabControl1.SelectedTab.Controls.OfType<CheckBox>().Count(); i++)
                {
                    Console.WriteLine(i);
                    foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                        c.CheckState = ResPages[g].CheckStates[i];
                }*/

                if (ResPages[g].viewed)
                {
                    int u = 0;
                    foreach (CheckBox c in tabControl1.SelectedTab.Controls.OfType<CheckBox>())
                    {
                        c.CheckState = ResPages[g].CheckStates[u];
                        u++;
                    }
                }
                else if (!ResPages[g].viewed)
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
        public List<CheckBox> checkboxes = new List<CheckBox>();
        public List<CheckState> CheckStates = new List<CheckState>();
        public bool viewed { get; set; }
    }


    public enum ChangePage
    {
        Start,
        Next,
        Previous,
    }
}
