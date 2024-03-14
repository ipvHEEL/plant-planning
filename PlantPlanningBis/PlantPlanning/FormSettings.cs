using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlantPlanning
{
    public partial class FormSettings : Form
    {
        FormMain fm;
        Int64  ID  = 0;

        public FormSettings(FormMain FM)
        {
            InitializeComponent();
            fm = FM;

            try
            {
                var data = fm.tdb.SystemTable.First();

                tbUrLic.Text = data.CompName;
                tbObject.Text = data.ObjectCode;
                tbCFOCode.Text = data.CFOCode;
                tbCFOSogl.Text = data.CFOSolut;
                tbAddress.Text = data.Address;
                ID = data.id;
            }
            catch (Exception xx)
            {
                tbUrLic.Text = "";
                tbObject.Text = "";
                tbCFOCode.Text = "";
                tbCFOSogl.Text = "";
                tbAddress.Text = "";
                ID = 0;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormSettings_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

          //  fm.MenuEnabled();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            fm.tdb.Pr_InsertUserCompData(ID, tbUrLic.Text, tbObject.Text, tbCFOCode.Text, tbCFOSogl.Text, tbAddress.Text);

            this.Close();
        }
    }
}
