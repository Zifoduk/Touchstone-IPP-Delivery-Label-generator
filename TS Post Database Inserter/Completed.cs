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
    public partial class Completed : Form
    {
        private string loading = "Please Wait";

        private string completed =
            "Operations have finished. \nWill now return to start window. \nRecommended task: open P-Touch application and \nopen the premade file called 'LabelTemplate' \nMake sure database is setup correctly.For help contact Admin";


        CustInfo CustomerInfo;
        public Completed(CustInfo cust)
        {
            Task task = Task.Factory.StartNew(() => InitializeComponent());
            task.Wait();
            OnScreenText.Text = loading;
            CustomerInfo = cust;
            cust.PushExcel(this);
            OnScreenText.Text = completed;
            OKBtn.Enabled = true;
        }

        private void OKBtn_Click(object sender, EventArgs e)
        {
            this.Close();
            CustomerInfo.Start.OpenLBX();
        }
    }
}
