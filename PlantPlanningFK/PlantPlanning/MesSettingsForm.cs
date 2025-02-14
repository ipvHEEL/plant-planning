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
    public partial class MesSettingsForm : Form
    {
        FormMain fm;
        List<MesLocations> MLList;
        List<MesAreas> MAList;
        Int64 ID = 0;

        public MesSettingsForm(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            MLList = new List<MesLocations>();
            MAList = new List<MesAreas>();
            LoadData();
        }

        private void LoadData()
        {
            MAList = fm.tdb.MesAreas.ToList();
            dataGridViewParent.DataSource = MAList;

            dataGridViewParent.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
        }

        private void dataGridViewParent_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewChild.DataSource = null;
            dataGridViewChild.Columns.Clear();
            MLList = new List<MesLocations>();

            buttonParentOK.Enabled = false;
            buttonParentCancel.Enabled = false;
            buttonChildCancel.Enabled = false;
            buttonChildOK.Enabled = false;

            ID = 0;

            if (dataGridViewParent.RowCount>0)
            {
                if (Convert.ToBoolean(dataGridViewParent.CurrentRow.Cells[2].Value.ToString())==true)
                {
                    buttonParentCancel.Enabled = true;
                }
                else
                {
                    buttonParentOK.Enabled = true;
                }

                ID = Convert.ToInt64(dataGridViewParent.CurrentRow.Cells[0].Value.ToString());

                MLList = fm.tdb.MesLocations.Where(x => x.AreaID == ID).ToList();
                dataGridViewChild.DataSource = MLList;
                dataGridViewChild.Columns[1].Visible = false;
                dataGridViewChild.Columns[4].Visible = false;
                dataGridViewChild.Columns[5].Visible = false;

                dataGridViewChild.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;          
            }
        }

        private void MesSettingsForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            fm.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridViewParent.RowCount > 0)
            {
                fm.tdb.Pr_UpdateMesArea(Convert.ToInt64(dataGridViewParent.CurrentRow.Cells["ID"].Value.ToString()), true);

                dataGridViewChild.DataSource = null;
                dataGridViewChild.Columns.Clear();
                dataGridViewParent.DataSource = null;
                dataGridViewParent.Columns.Clear();

                LoadData();
                fm.GetStorageList();
            }
        }

        private void dataGridViewParent_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (dataGridViewParent.RowCount>0)
            {
                foreach (DataGridViewRow r in dataGridViewParent.Rows)
                {
                    if (Convert.ToBoolean(r.Cells[2].Value.ToString())==true)
                    {
                        r.Cells[2].Style.BackColor = Color.DarkGreen;
                    }
                }
            }
        }

        private void dataGridViewChild_SelectionChanged(object sender, EventArgs e)
        {
              buttonChildCancel.Enabled = false;
              buttonChildOK.Enabled = false;
            if (dataGridViewChild.RowCount > 0)
            {
                if (Convert.ToBoolean(dataGridViewChild.CurrentRow.Cells[3].Value.ToString()) == true)
                {
                    buttonChildCancel.Enabled = true;
                }
                else
                {
                    buttonChildOK.Enabled = true;
                }
            }
        }

        private void dataGridViewChild_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (dataGridViewChild.RowCount > 0)
            {
                foreach (DataGridViewRow r in dataGridViewChild.Rows)
                {
                    if (Convert.ToBoolean(r.Cells[3].Value.ToString()) == true)
                    {
                        r.Cells[3].Style.BackColor = Color.DarkGreen;
                    }
                }
            }
        }

        private void buttonParentCancel_Click(object sender, EventArgs e)
        {
            if (dataGridViewParent.RowCount > 0)
            {
                fm.tdb.Pr_UpdateMesArea(Convert.ToInt64(dataGridViewParent.CurrentRow.Cells["ID"].Value.ToString()), false);

                dataGridViewChild.DataSource = null;
                dataGridViewChild.Columns.Clear();
                dataGridViewParent.DataSource = null;
                dataGridViewParent.Columns.Clear();

                LoadData();
                fm.GetStorageList();
            }
        }

        private void buttonChildOK_Click(object sender, EventArgs e)
        {
            if (dataGridViewChild.RowCount > 0)
            {
                fm.tdb.Pr_UpdateMesLocation(Convert.ToInt64(dataGridViewChild.CurrentRow.Cells["ID"].Value.ToString()), true);

                dataGridViewChild.DataSource = null;
                dataGridViewChild.Columns.Clear();

                fm.GetStorageList();

                if (dataGridViewParent.RowCount > 0)
                {
                    if (Convert.ToBoolean(dataGridViewParent.CurrentRow.Cells[2].Value.ToString()) == true)
                    {
                        buttonParentCancel.Enabled = true;
                    }
                    else
                    {
                        buttonParentOK.Enabled = true;
                    }

                    ID = Convert.ToInt64(dataGridViewParent.CurrentRow.Cells[0].Value.ToString());

                    MLList = fm.tdb.MesLocations.Where(x => x.AreaID == ID).ToList();
                    dataGridViewChild.DataSource = MLList;
                    dataGridViewChild.Columns[1].Visible = false;
                    dataGridViewChild.Columns[4].Visible = false;
                    dataGridViewChild.Columns[5].Visible = false;

                    dataGridViewChild.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                }
            }
        }

        private void buttonChildCancel_Click(object sender, EventArgs e)
        {
            if (dataGridViewChild.RowCount > 0)
            {
                fm.tdb.Pr_UpdateMesLocation(Convert.ToInt64(dataGridViewChild.CurrentRow.Cells["ID"].Value.ToString()), false);

                dataGridViewChild.DataSource = null;
                dataGridViewChild.Columns.Clear();

                fm.GetStorageList();

                if (dataGridViewParent.RowCount > 0)
                {
                    if (Convert.ToBoolean(dataGridViewParent.CurrentRow.Cells[2].Value.ToString()) == true)
                    {
                        buttonParentCancel.Enabled = true;
                    }
                    else
                    {
                        buttonParentOK.Enabled = true;
                    }

                    ID = Convert.ToInt64(dataGridViewParent.CurrentRow.Cells[0].Value.ToString());

                    MLList = fm.tdb.MesLocations.Where(x => x.AreaID == ID).ToList();
                    dataGridViewChild.DataSource = MLList;
                    dataGridViewChild.Columns[1].Visible = false;
                    dataGridViewChild.Columns[4].Visible = false;
                    dataGridViewChild.Columns[5].Visible = false;

                    dataGridViewChild.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;
                }
            }
        }
    }
}
