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
    public partial class FormAddCard : Form
    {
        FormMain fm;
        Int64 ID = 0;

        public FormAddCard(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            LoadData();
        }

        private void LoadData()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();

            var data = (from x in fm.tdb.tWorkPlan
                        select x).ToList();

            dataGridView1.DataSource = data;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;
            dataGridView1.Columns["id"].Visible = false;

            dataGridView1.Refresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            try
            {
                ID = Convert.ToInt64(dataGridView1.CurrentRow.Cells["id"].Value.ToString());
            }
            catch (Exception xx)
            {
                ID = 0;
            }
            DateTime DateWork;

            if (ID > 0)
            {
                fm.POList.Clear();
                Cursor = Cursors.WaitCursor;

                lData LD;

                var twpD = fm.tdb.tWorkPlanDates.Where(x => x.owner == ID).ToList();

                foreach (var t in twpD)
                {                   
                    Int64 id = t.id;
                    var twpE = fm.tdb.tWorkPlanElement.Where(x => x.owner == id).ToList();

                    foreach (var te in twpE)
                    {
                        LD = new lData
                        {
                            DateWork = t.DateWork,
                            Line = fm.ReplaceLineName(te.LineName),
                            ProdCode = te.ProdCode.Replace("_П/УП", ""),
                            ProdCodeStr = te.ProdCode,
                            RePack = te.IsRePack,
                            Quantity = te.ProdQuantity,
                            ProdName = te.ProdName
                        };

                        DateWork = t.DateWork;
                        DateWork = new DateTime(DateWork.Year, DateWork.Month, 1);
                        fm.ExcelDatesList.Add(DateWork);

                        fm.ExcelDataList.Add(LD);
                    }               
                }

                fm.ExcelDatesList = fm.ExcelDatesList.Distinct().ToList();

                fm.GetPecipesNotes();
                fm.GreateOperList();
                fm.AlterProductNames();
                Cursor = Cursors.Default;
                this.Close();
            }
        }

        private void FormAddCard_FormClosed(object sender, FormClosedEventArgs e)
        {
            fm.Enabled = true;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
