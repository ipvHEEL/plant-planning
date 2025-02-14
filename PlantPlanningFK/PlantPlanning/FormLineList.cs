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
    public partial class FormLineList : Form
    {
        FormMain fm;
        ConsurmptionForm cf;

        public FormLineList(FormMain FM)
        {
            InitializeComponent();
            listBox1.Items.Clear();
            fm = FM;
            LoadData();
        }

        private void LoadData()
        {
            listBox1.Items.Clear();
            listBox1.Items.Add("ВСЕ ЛИНИИ:");

            for (int i=0; i<fm.LineList.Count; i++)
            {
                listBox1.Items.Add(fm.LineList[i]);
            }
            GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["ConsurmptionForm"] == null)
            {
                Cursor = Cursors.WaitCursor;
                cf = new ConsurmptionForm(fm); //, this
                cf.Visible = true;
                //   fs.MdiParent = this;
                cf.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                cf.BringToFront();
                cf.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormLineList_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }
    }
}
