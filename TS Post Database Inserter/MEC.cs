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
    public partial class MEC : Form
    {
        public MEC(int i)
        {
            InitializeComponent();
            string main = "Main Excel document has changed";
            string master = "Master Excel document has changed";
            if (i == 0)
                label1.Text = master;
            if (i == 1)
                label1.Text = main;
        }

        private void ConBTN_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }
}
