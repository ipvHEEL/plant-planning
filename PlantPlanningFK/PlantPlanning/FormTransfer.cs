using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel1 = Microsoft.Office.Interop.Excel;

namespace PlantPlanning
{
    public partial class FormTransfer : Form
    {
        FormMain fm;
        ConsurmptionForm cf;
        UserSelectedCell usc;
        List<Material> ml;
        ComboData CD;
        ComboData DD;

        List<ComboData> CDList;

        public FormTransfer(FormMain FM, ConsurmptionForm CF, UserSelectedCell USC, List<Material> ML)
        {
            InitializeComponent();

            fm = FM;
            cf = CF;
            usc = USC;
            ml = ML;

            if (ml != null)
            {
                    LoadData();                
            }
        }

        private void LoadData()
        {
            label2.Text = usc.UserSelectedCode;
            label4.Text = usc.UserSelectedProdName;
            dateTimePicker1.Value = usc.UserSelectedDateTime.Date;
            dateTimePicker2.Value = usc.UserSelectedDateTime.Date;
            cbSource.Items.Clear();
            cbSource.Text = "";

            CDList = new List<ComboData>();

            if (ml.Count > 0)
            {
                for (int i = 0; i < ml[0].RawStorageListAlter.Count; i++)
                {
                    ComboData cd = new ComboData();

                    cd.StorageID = ml[0].RawStorageListAlter[i].StorageID;
                    cd.StorageName = ml[0].RawStorageListAlter[i].Storage + " " + ml[0].RawStorageListAlter[i].Quantity.ToString("N3");
                    cd.StorageQuant = ml[0].RawStorageListAlter[i].Quantity;

                    CDList.Add(cd);
                }
            }

            cbSource.DataSource = CDList;
            cbSource.DisplayMember = "StorageName";
            cbSource.ValueMember = "StorageID";
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormTransfer_FormClosed(object sender, FormClosedEventArgs e)
        {
            cf.Enabled = true;
        }

        private void cbSource_SelectedIndexChanged(object sender, EventArgs e)
        {
             CD = (ComboData)cbSource.SelectedItem;
            int ID = 0;

            try
            {
                ID = CD.StorageID;
            }
            catch (Exception xx)
            { }

            cbDest.DataSource = null;
            cbDest.Items.Clear();
            cbDest.Text = "";

            var Data = fm.tdb.fn_select_LinkedCombo(ID).ToList();

            List<ComboData> Data1 = new List<ComboData>();

            for (int i = 0; i<Data.Count; i++)
            {
                ComboData dd = new ComboData();

                dd.StorageID = (int)Data[i].AreaID;
                dd.StorageName = Data[i].AreaName;
                dd.StorageQuant = 0;

                Data1.Add(dd);
            }

            cbDest.DataSource = Data1;
            cbDest.DisplayMember = "StorageName";
            cbDest.ValueMember = "StorageID";
        }

        private void button1_Click(object sender, EventArgs e)
        {
           if ( cbSource.SelectedItem != null)
            {
                if (cbDest.SelectedItem != null)
                {
                    CD = (ComboData)cbSource.SelectedItem;
                    ComboData DD = (ComboData)cbDest.SelectedItem;

                    double Quant = fm.ConvertStringToDouble(textBox1.Text);

                    if (Quant <= CD.StorageQuant)
                    {
                        var Res = fm.tdb.pr_InrestTransferData(usc.UserSelectedCode, dateTimePicker1.Value.Date, Quant, dateTimePicker2.Value.Date, CD.StorageID, DD.StorageID);
                        fm.tdb.SaveChanges();

                        cf.Enabled = true;
                        cf.UpdateStorageTask(usc.RowIndex, usc.ColIndex, Quant);

                        this.Close();
                    }
                    else
                    {
                        MessageBox.Show("Превышено количество материала на складе!", "Сообщение системы");
                    }
                }
                else
                {
                    MessageBox.Show("Не выбран склад-приемник!", "Сообщение системы");
                }
            }
           else
            {
                MessageBox.Show("Не выбран склад-источник!", "Сообщение системы");
            }
        }
    }
}
