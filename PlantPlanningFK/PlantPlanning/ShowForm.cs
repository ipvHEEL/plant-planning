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
    public partial class ShowForm : Form
    {
        public ShowForm()
        {
            InitializeComponent();
            progressBar1.Visible = true;
          //  progressBar1.Value = 99;
        }

        private void timer1_Tick(object sender, EventArgs e)
        {
           /* progressBar1.Value = progressBar1.Value + 7;
            
            if (progressBar1.Value >90)
            {
                progressBar1.Value = 1;
            }*/
        }
    }
}
