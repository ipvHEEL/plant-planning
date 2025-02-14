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
//using PlantPlanning.Context;

namespace PlantPlanning
{
    public partial class FormCommdet : Form
    {
        FormMain fm;
        Int64 ID = 0;

        List<Context.fn_select_CommentData_Result> Dlist;
        BindingSource bs;

        public FormCommdet(FormMain FM)
        {
            InitializeComponent();

            fm = FM;
            LoadData();
        }

        private void LoadData()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            /*  Dlist = fm.tdb.fn_select_CommentData().ToList();
              bs = new BindingSource() { DataSource = Dlist };


              dataGridView1.DataSource = bs;  */ //Dlist;


            string SqlString = "SELECT * from fn_select_CommentData()";  // "select * from tPartyDeclaration";
            bs = fm.ConnectionSource(fm.connTest, SqlString, "TSubstring");
            dataGridView1.DataSource = bs;

            dataGridView1.Columns[2].Width = 250;
            dataGridView1.Columns[6].Width = 500;
            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;

        }

        private void FormCommdet_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }

        private void dataGridView1_SelectionChanged(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                try
                {
                    ID = Convert.ToInt64(dataGridView1.CurrentRow.Cells["ID"].Value.ToString());
                    textBoxCode.Text = dataGridView1.CurrentRow.Cells["MaterialCode"].Value.ToString();
                    textBoxNaim.Text = dataGridView1.CurrentRow.Cells["MaterialName"].Value.ToString();
                    checkBox1.Checked = Convert.ToBoolean(dataGridView1.CurrentRow.Cells["IsActual"].Value.ToString());
                    textBoxComment.Text = dataGridView1.CurrentRow.Cells["Comment"].Value.ToString();
                }
                catch (Exception xx)
                {
                    ID = 0;
                    textBoxCode.Text = "";
                    textBoxNaim.Text = "";
                    textBoxComment.Text = "";
                    checkBox1.Checked = false;
                }
            }
        }

        private void button4_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            ID = 0;
            textBoxCode.Text = "";
            textBoxNaim.Text = "";
            textBoxComment.Text = "";
            checkBox1.Checked = false;

            textBoxCode.Focus();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                string FileName = openFileDialog1.FileName;
                Cursor = Cursors.WaitCursor;

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = false;
                Excel1.Workbook wBook = xlApp.Workbooks.Open(FileName);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets["Лист1"];
                Excel1.Range range1;

                Int32 CurrRow = 2;
                bool fl = true;
                string Code = "";
                string Naim = "";
                string Comment = "";
                string Username = System.Environment.UserName;

                while (fl==true)
                {
                    range1 = (Excel1.Range)wSheet.Cells[CurrRow, 1];

                    try
                    {
                        Code = range1.Value.ToString().Trim();
                        fl = true;
                    }
                    catch (Exception xx )
                    {
                        Code = "";
                        fl = false;
                    }

                    range1 = (Excel1.Range)wSheet.Cells[CurrRow, 2];

                    try
                    {
                        Naim = range1.Value.ToString().Trim();                       
                    }
                    catch (Exception xx)
                    {
                        Naim = "";
                    }

                    range1 = (Excel1.Range)wSheet.Cells[CurrRow, 3];

                    try
                    {
                        Comment = range1.Value.ToString().Trim();
                    }
                    catch (Exception xx)
                    {
                        Comment = "";
                    }

                    range1 = (Excel1.Range)wSheet.Cells[CurrRow, 4];
                    try
                    {
                        Username= range1.Value.ToString().Trim();
                    }
                    catch (Exception xx)
                    { }

                    if (fl==true)
                    {
                      var  res = fm.tdb.pr_InsertCommentData1(0, Code, Naim, true, Comment, Username);
                    }

                    CurrRow = CurrRow + 1;
                }

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp.Quit();

                LoadData();
                Cursor = Cursors.Default;
                               
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            var res = fm.tdb.pr_InsertCommentData1(ID, textBoxCode.Text.Trim(), textBoxNaim.Text.Trim(), checkBox1.Checked, textBoxComment.Text.Trim(), System.Environment.UserName);
            LoadData();
        }

        private void textBox1_TextChanged(object sender, EventArgs e)
        {
            string Filt = "";

            if ((textFilterCode.Text.Trim() != "") || (textBoxFilterNaim.Text.Trim() != ""))
            {           
                if (textFilterCode.Text.Trim() != "")
                {
                    Filt = " MaterialCode like '%" + textFilterCode.Text.Trim() + "%' ";
                }

                if (textBoxFilterNaim.Text.Trim() != "")
                {
                    if (Filt != "")
                    {
                        Filt = Filt + " and ";
                    }
                    Filt = Filt + " MaterialName like '%" + textBoxFilterNaim.Text.Trim() + "%' ";
                }      
            }

            bs.Filter = Filt;
        }
    

        private void button9_Click(object sender, EventArgs e)
        {
            textFilterCode.Text = "";
            bs.Filter = "";
            textBoxFilterNaim.Text = "";
            
        }
    }
}
