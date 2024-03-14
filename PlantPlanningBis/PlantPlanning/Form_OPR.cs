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
using System.Data.SqlClient;

namespace PlantPlanning
{
    public partial class Form_OPR : Form
    {
        FormMain fm;
        BindingSource bs;
        List<PlanOPZList> pList;
        Int64 ID = 0;

        public Form_OPR(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            bs = new BindingSource();
            LoadData();
        }

        private void LoadData()
        {
            pList = fm.tdb.PlanOPZList.ToList();  // fm.helg.PlanOPZList.ToList();
            bs.DataSource = pList;
            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();
            dataGridView1.DataSource = bs;

            dataGridView1.Columns[0].Width = 50;
            dataGridView1.Columns[1].Width = 150;
            dataGridView1.Columns[2].Width = 700;
        }

        private void Form_OPR_Resize(object sender, EventArgs e)
        {
            if (this.Width<600)
            {
                this.Width = 600;
            }
            if (this.Height<800)
            {
                this.Height = 800;
            }
            button2.Left = (this.Width - button2.Width) / 2;
        }

        private void Form_OPR_FormClosed(object sender, FormClosedEventArgs e)
        {
            fm.Enabled = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount>0)
            {
                ID = Convert.ToInt64(dataGridView1.CurrentRow.Cells[0].Value.ToString());
                textBoxCodeValue.Text = dataGridView1.CurrentRow.Cells[1].Value.ToString();
                textBoxNameValue.Text = dataGridView1.CurrentRow.Cells[2].Value.ToString();
                checkBox1.Checked = Convert.ToBoolean(dataGridView1.CurrentRow.Cells[3].Value.ToString());
            }
            else
            {
                textBoxCodeValue.Text = "";
                textBoxNameValue.Text = "";
                checkBox1.Checked = false;
            }
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            bs.Filter = "";
            if (textBoxCode.Text.Trim()!="")
            {
                bs.Filter = " MatCode like '%" + textBoxCode.Text.Trim() + "%' ";
            }
            if (textBoxNaim.Text.Trim()!="")
            {
                if (bs.Filter.Length>0)
                {
                    bs.Filter = bs.Filter + " and MatName like '%" + textBoxNaim.Text.Trim() + "%' ";
                }
                else
                {
                    bs.Filter= " MatName like '%" + textBoxNaim.Text.Trim() + "%' ";
                }
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ID = 0;
            textBoxCodeValue.Text = "";
            textBoxNameValue.Text = "";
            checkBox1.Checked = false;
        }

        private void button2_Click(object sender, EventArgs e)
        {
          DialogResult dr = MessageBox.Show("Подтверждаете удаление записи?", "Сообщение системы", MessageBoxButtons.OKCancel);
            if (dr==DialogResult.OK)
            {
               // PlanOPZList pol = fm.helg.PlanOPZList.Where(x => x.ID == ID).First();
              //  fm.helg.DeleteObject(pol);
              //if (pol!=null)
                {
                    fm.tdb.pr_DeleteTestData(ID);
                    /*fm.helg.PlanOPZList.Remove(pol);
                    fm.helg.SaveChanges();*/
                }
                LoadData();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            fm.tdb.pr_InsertTestData1(ID, textBoxCodeValue.Text.Trim(), textBoxNameValue.Text.Trim(), checkBox1.Checked);

           /* if (ID!=0)
            {
                
                PlanOPZList pol = fm.helg.PlanOPZList.Where(x => x.ID == ID).First();
                if (pol!=null)
                {
               //     pol.MatCode = textBoxCodeValue.Text.Trim();
                 //   pol.MatName = textBoxNameValue.Text.Trim();
                 //   fm.helg.SaveChanges();
                }
            }
            else
            {
                PlanOPZList pol = new PlanOPZList
                {
                   // MatCode = textBoxCodeValue.Text.Trim(),
                  //  MatName = textBoxNameValue.Text.Trim()
                };
                //fm.helg.PlanOPZList.Add(pol);
                //fm.helg.SaveChanges();
            }*/
            LoadData();
        }
    }
}
