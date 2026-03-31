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
    public partial class FormTaskClear : Form
    {
        ConsurmptionForm cf;
        FormMain fm;

        public FormTaskClear(ConsurmptionForm CF, FormMain FM)
        {
            InitializeComponent();

            cf = CF;
            fm = FM;

            LoadData();
        }

        private void LoadData()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();

            var App= fm.tdb.tApplication.ToList();

            List<AppData> ADList = new List<AppData>();

            foreach (var a in App)
            {
                AppData ad = new AppData()
                {
                    id = a.id,
                    DateWork = a.DateWork,
                    AlterDate = a.AlterDate,
                    MaterialCode = a.MaterialCode,
                    MaterialName = cf.CFMList.Where(x => x.MaterialCode == a.MaterialCode).Select(x => x.MaterialName).FirstOrDefault(),
                    Quantity = a.Quantity
                };

                ADList.Add(ad);
            }

            dataGridView1.DataSource = ADList;
            dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

            dataGridView2.DataSource = null;
            dataGridView2.Columns.Clear();

            var Stor = fm.tdb.tStorageTask.ToList();
            List<TransferData> StorList = new List<TransferData>();

            foreach (var s in Stor)
            {
                TransferData td = new TransferData();

                td.id = s.id;
                td.DateWork = s.DateWork;
                td.AlterDate = (DateTime)s.AlterDate;
                td.MaterialCode = s.MaterialCode;
                td.MaterialName = cf.CFMList.Where(x => x.MaterialCode == s.MaterialCode).Select(x => x.MaterialName).FirstOrDefault();
                td.Quantity = s.Quantity;
                td.SourceID = (int)s.Source;
                td.MesStorageName = fm.tdb.MesAreas.Where(x => x.AreaID == (int)s.Source).Select(x => x.AreaName).FirstOrDefault();
                td.DestID = (int)s.Dest;
                td.PlantStorageName = fm.tdb.PlantAreas.Where(x => x.AreaID == (int)s.Dest).Select(x => x.AreaName).FirstOrDefault();                

                StorList.Add(td);
            }

            dataGridView2.DataSource =StorList;
            dataGridView2.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void FormTaskClear_FormClosed(object sender, FormClosedEventArgs e)
        { 
            cf.Enabled = true; 
            
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var data = fm.tdb.tApplication;

            Cursor = Cursors.WaitCursor;
            data.RemoveRange(data);
            fm.tdb.SaveChanges();

            cf.ShowData();
            Cursor = Cursors.Default;
        //    this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.SelectedRows.Count > 0)
            {
                Cursor = Cursors.WaitCursor;

                foreach (DataGridViewRow r in dataGridView1.SelectedRows)
                {
                    Int64 ID =Convert.ToInt64( r.Cells["id"].Value);
                    var res = fm.tdb.Pr_DeleteAppRecord(ID);
                }

                fm.tdb.SaveChanges();

                GC.Collect();
                LoadData();
                cf.ShowData();
               
                Cursor = Cursors.Default;
            }
            else
            {
                MessageBox.Show("Выделите одну или несколько строк!", "Сообщение системы");
            }
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView2.SelectedRows.Count > 0)
            {
                Cursor = Cursors.WaitCursor;

                foreach (DataGridViewRow r in dataGridView2.SelectedRows)
                {
                    Int64 ID = Convert.ToInt64(r.Cells["id"].Value);
                    var res = fm.tdb.Pr_DeleteStorageTaskRecord(ID);
                }

                fm.tdb.SaveChanges();

                GC.Collect();
                LoadData();
                cf.ShowData();
                
                Cursor = Cursors.Default;
            }
            else
            {
                MessageBox.Show("Выделите одну или несколько строк!", "Сообщение системы");
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button6_Click(object sender, EventArgs e)
        {
            var data = fm.tdb.tStorageTask;

            Cursor = Cursors.WaitCursor;
            data.RemoveRange(data);
            fm.tdb.SaveChanges();

            cf.ShowData();
            Cursor = Cursors.Default;
          //  this.Close();
        }
    }
}
