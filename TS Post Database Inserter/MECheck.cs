using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace TS_Post_Database_Inserter
{
    public partial class MECheck : Form
    {
        Start st;
        string FName;
        public MECheck(Start f, string h, int i)
        {
            InitializeComponent();
            string master = "Do you want to change the Master Folder?";
            label1.Text = master;
            st = f;
            Console.WriteLine(h);
            FName = h;
        }

        private void NoBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void YesBTN_Click(object sender, EventArgs e)
        {
            st.Folder = FName;

            MEC Mec = new MEC();
            Mec.ShowDialog();
            this.Close();
            
        }
    }
}
