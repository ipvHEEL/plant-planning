using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PlantPlanning.Context;

namespace PlantPlanning
{
    public partial class FormSaveCard : Form
    {
        FormMain fm;
        BindingSource bs;

        public FormSaveCard(FormMain FM)
        {
            InitializeComponent();
            textBox3.Text = DateTime.Today.Date.ToString("dd.MM.yyyy");
            fm = FM;

            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();

            var data = (from x in fm.tdb.tWorkPlan
                       select x).ToList();

            dataGridView1.DataSource = data;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;

            dataGridView1.Refresh();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (fm.POList.Count > 0)
            {
                string UserName = System.Environment.UserName;
                Int64 ID = 0;
                Int64 ID1 = 0;

                tWorkPlan twp = fm.tdb.tWorkPlan.Add(new tWorkPlan
                {
                    PlanName = textBox4.Text.Trim(),
                    DateWork = DateTime.Today.Date,
                    UserName = UserName
                });

                fm.tdb.SaveChanges();
                ID = twp.id;

                foreach (var pol in fm.POList)
                {
                    var twpD = fm.tdb.tWorkPlanDates.Add(new tWorkPlanDates
                    {
                        owner=ID,
                        DateWork=pol.DateWork
                    });

                    fm.tdb.SaveChanges();
                    ID1 = twpD.id;

                    foreach (var ddl in pol.DL)
                    {
                        if (ddl.LastRec == false)
                        {
                            var twpE = fm.tdb.tWorkPlanElement.Add(new tWorkPlanElement
                            {
                                owner = ID1,
                                LineName = ddl.Line,
                                ProdCode = ddl.ProdCodeStr,
                                IsRePack = ddl.RePack,
                                ProdName = ddl.ProdName,
                                ProdQuantity = ddl.Quantity
                            });

                            fm.tdb.SaveChanges();
                        }
                    }
                }
                            
                GC.Collect();
                MessageBox.Show("Данные успешно сохранены");
                this.Close();
            }
        }

        private void FormSaveCard_FormClosed(object sender, FormClosedEventArgs e)
        {
            fm.Enabled = true;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }
    }
}
