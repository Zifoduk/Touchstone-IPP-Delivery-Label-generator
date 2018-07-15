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
        int x;
        public MECheck(Start f, string h, int i)
        {
            x = i;
            InitializeComponent();
            string main = "Do you want to change the Main Exel document?";
            string master = "Do you want to change the Main Exel document?";
            if (i == 0)
                label1.Text = master;
            if (i == 1)
                label1.Text = main;
            st = f;
            FName = h;
        }

        private void NoBtn_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void YesBTN_Click(object sender, EventArgs e)
        {
            if (x == 0)
            {
                st.ElL.Text = FName;
                st.MasterExcel = FName;
                st.ElL.ForeColor = Color.Black;
            }
            if (x == 1)
            {
                st.PHExcelL.Text = FName;
                st.MainExcel = FName;
                st.PHExcelL.ForeColor = Color.Black;
            }
            MEC Mec = new MEC(x);
            Mec.ShowDialog();
            this.Close();
            
        }
    }
}
