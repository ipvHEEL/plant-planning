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
    public partial class FormRequest : Form
    {
        FormMain fm;
        ConsurmptionForm cf;
        UserSelectedCell USC;

        public FormRequest(FormMain FM, ConsurmptionForm CF, UserSelectedCell usc)
        {
            InitializeComponent();
            fm = FM;
            cf = CF;
            USC = usc;

            label2.Text = USC.UserSelectedCode;
            label4.Text = USC.UserSelectedProdName;
            dateTimePicker1.Value = USC.UserSelectedDateTime.Date;
            dateTimePicker2.Value = USC.UserSelectedDateTime.Date;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double Quant;
            try
            {
                Quant = fm.ConvertStringToDouble(textBox1.Text.Trim());
                if (Quant > 0)
                {
                    fm.tdb.Pr_InsertUserAppData1(dateTimePicker1.Value.Date, dateTimePicker2.Value.Date, label2.Text, Quant);
                    fm.tdb.SaveChanges();
                }

                cf.button2.Enabled = true;
                cf.button3.Enabled = true;
                cf.Enabled = true;
                cf.UpdateApplication(USC.RowIndex, USC.ColIndex, Quant);            

                this.Close();
            }
            catch (Exception xx)
            {
                MessageBox.Show("Некорректные данные", "Сообщение системы");
                textBox1.Focus();
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            //
        }

        private void FormRequest_FormClosed(object sender, FormClosedEventArgs e)
        {
            cf.Enabled = true;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            cf.UserSelectedCode = "";
            cf.UserSelectedProdName = "";

            this.Dispose();
        }

        private void textBox1_Leave(object sender, EventArgs e)
        {
            //
        }
    }
}
