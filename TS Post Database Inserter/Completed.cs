using System;
using System.Threading;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Reflection;

namespace TS_Post_Database_Inserter
{
    public partial class Completed : Form
    {
        private string loading = "Please Wait";

        private string completed =
            "Operations have finished.\nP-Touch application will now open";


        CustInfo CustomerInfo;
        public Completed(CustInfo cust)
        {
            Task task = Task.Factory.StartNew(() => InitializeComponent());
            task.Wait();
            CustomerInfo = cust;
        }



        private delegate void SetControlPropertiesDelegate(Control control, string Property, Object PropertyValue);
        public static void SetControlProperty(Control control, string Property, Object PropertyValue)
        {
            if(control.InvokeRequired)
            {
                control.Invoke(new SetControlPropertiesDelegate(SetControlProperty), new object[] { control, Property, PropertyValue });
            }
            else
            {
                control.GetType().InvokeMember(Property, BindingFlags.SetProperty, null, control, new object[] { PropertyValue });
            }
        }

        public void Bar()
        {
            while (true)
            {
                if (Progressbar.Value == Progressbar.Maximum)
                {
                    SetControlProperty(OnScreenText, "Text", completed);
                    SetControlProperty(OKBtn, "Enabled", true);
                }
                else
                {
                    SetControlProperty(OnScreenText, "Text", loading);
                    SetControlProperty(OKBtn, "Enabled", false);

                }
                Thread.Sleep(50);
            }
        }

        private void OKBtn_Click(object sender, EventArgs e)
        {
            this.Close();
            CustomerInfo.start.OpenLBX();
        }

        private void Completed_Shown(object sender, EventArgs e)
        {
            CustomerInfo.PushExcel(this);
            ThreadStart job = new ThreadStart(Bar);
            Thread thread = new Thread(job);
            thread.Start();
        }
    }
}
