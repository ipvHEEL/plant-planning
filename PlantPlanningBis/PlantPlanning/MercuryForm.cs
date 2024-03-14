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
    public partial class MercuryForm : Form
    {
        FormMain fm;
        BindingSource bs = new BindingSource(); 
        public MercuryForm(FormMain FM)
        {
            InitializeComponent();

            fm = FM;
            LoadData();
        }

        private void LoadData()
        {
            Cursor = Cursors.WaitCursor;
            //  var MData = fm.tdb.pr_GetMercuryComponent().ToList();

            dataGridView1.DataSource = null;
            dataGridView1.Columns.Clear();

            string SqlString = "select * from fn_select_MercuryComponents()";
            bs = fm.ConnectionSource(fm.connTest, SqlString, "Mstring");
            dataGridView1.DataSource = bs;

            try
            {
                dataGridView1.Columns[1].Width = 300;
            }
            catch (Exception xx)
            { }

            try
            {
                dataGridView1.Columns[4].Visible = false;
            }
            catch (Exception xx)
            { }

            try
            {
                dataGridView1.Columns[5].Visible = false;
            }
            catch (Exception xx)
            { }

            try
            {
                dataGridView1.Columns[9].Width = 300;
            }
            catch (Exception xx)
            { }

            try
            {
                dataGridView1.Columns[12].Width = 300;
            }
            catch (Exception xx)
            { }

            dataGridView1.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
            dataGridView1.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            Cursor = Cursors.Default;
        }

        private void MercuryForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }

        private void dataGridView1_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {

        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                DateTime dt;

                try
                {
                    dt = Convert.ToDateTime(r.Cells["EndDate"].Value.ToString());
                }
                catch
                {
                    dt = Convert.ToDateTime("1999-01-01");
                }


                if (dt.Year < 2000)
                {
                    r.DefaultCellStyle.BackColor = Color.Plum;
                }
                else
                {
                    int days = (dt - DateTime.Today.Date).Days;
                    if (days > 30)
                    {
                        r.DefaultCellStyle.BackColor = Color.PaleGreen;
                    }
                    else if ((days > 10) && (days <= 30))
                    {
                        r.DefaultCellStyle.BackColor = Color.LightYellow;
                    }
                    else   //<=10
                    {
                        r.DefaultCellStyle.BackColor = Color.LightCoral;
                    }
                }
            }
            Cursor = Cursors.Default;
        }

        private void textBoxProdCode_TextChanged(object sender, EventArgs e)
        {

          /* */
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            bs.Filter = "";

            string FilterStr = "";

            if (textBoxProdCode.Text.Trim() != "")
            {
                FilterStr = " RPItemNo like '%" + textBoxProdCode.Text.Trim() + "%' ";
            }

            if (textBoxProdName.Text.Trim() != "")
            {
                if (FilterStr.Length > 0)
                {
                    FilterStr = FilterStr + " and ";
                }

                FilterStr = FilterStr + " RPName like '%" + textBoxProdName.Text.Trim() + "%' ";
            }

            if (textBoxLotNo.Text.Trim() != "")
            {
                if (FilterStr.Length > 0)
                {
                    FilterStr = FilterStr + " and ";
                }

                FilterStr = FilterStr + " RPLotNo like '%" + textBoxLotNo.Text.Trim() + "%' ";
            }

            if (textBoxMatCode.Text.Trim() != "")
            {
                if (FilterStr.Length > 0)
                {
                    FilterStr = FilterStr + " and ";
                }

                FilterStr = FilterStr + " MaterialItemNo like '%" + textBoxMatCode.Text.Trim() + "%' ";
            }

            if (textBoxMatName.Text.Trim() != "")
            {
                if (FilterStr.Length > 0)
                {
                    FilterStr = FilterStr + " and ";
                }

                FilterStr = FilterStr + " MaterialName like '%" + textBoxMatName.Text.Trim() + "%' ";
            }

            if (FilterStr.Length > 0)
            {
                bs.Filter = FilterStr;
            }
            Cursor = Cursors.Default;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            bs.Filter = "";

            textBoxProdCode.Text = "";
            textBoxProdName.Text = "";
            textBoxLotNo.Text = "";
            textBoxMatCode.Text = "";
            textBoxMatName.Text = "";
            Cursor = Cursors.Default;
        }

        private void button5_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                Cursor = Cursors.WaitCursor;
                Int32 Currow = 2;

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];
                Excel1.Range range1;

                for (int i = 0; i < dataGridView1.Columns.Count; i++)
                {
                    wSheet.Cells[1, i + 1] = dataGridView1.Columns[i].HeaderText;
                }

                try
                {
                    range1 = wSheet.Cells[1, 1] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                { }
                try
                {
                    range1 = wSheet.Cells[1, 2] as Excel1.Range;
                    range1.ColumnWidth = 60;
                }
                catch (Exception xx)
                { }
                try
                {
                    range1 = wSheet.Cells[1, 3] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 4] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 5] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 6] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 7] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 8] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 9] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 10] as Excel1.Range;
                    range1.ColumnWidth = 60;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 11] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 12] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 13] as Excel1.Range;
                    range1.ColumnWidth = 60;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 14] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
                try
                {
                    range1 = wSheet.Cells[1, 15] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }
              /*  try
                {
                    range1 = wSheet.Cells[1, 16] as Excel1.Range;
                    range1.ColumnWidth = 15;
                }
                catch (Exception xx)
                {
                }*/
                try
                {
                    range1 = wSheet.Columns["E", misValue] as Excel1.Range;
                    range1.EntireColumn.Hidden = true;
                    range1 = wSheet.Columns["F", misValue] as Excel1.Range;
                    range1.EntireColumn.Hidden = true;
                }
                catch (Exception xx)
                {
                }

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dataGridView1.Columns.Count]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[dataGridView1.Rows.Count + 1, 1]];
                range1.NumberFormat = "@";
                range1 = wSheet.Range[wSheet.Cells[2, 3], wSheet.Cells[dataGridView1.Rows.Count + 1, 3]];
                range1.NumberFormat = "@";
                range1 = wSheet.Range[wSheet.Cells[2, 9], wSheet.Cells[dataGridView1.Rows.Count + 1, 9]];
                range1.NumberFormat = "@";
                range1 = wSheet.Range[wSheet.Cells[2, 12], wSheet.Cells[dataGridView1.Rows.Count + 1, 12]];
                range1.NumberFormat = "@";

                foreach (DataGridViewRow r in dataGridView1.Rows)
                {
                    for (int i=0; i<r.Cells.Count - 2; i++) //
                    {
                        try
                        {
                            wSheet.Cells[Currow, i + 1] = r.Cells[i].Value;
                        }
                        catch (Exception xx)
                        {
                            wSheet.Cells[Currow, i + 1] = "ERROR!";
                        }
                    }

                    try
                    {
                        wSheet.Cells[Currow, r.Cells.Count - 1] = r.Cells[r.Cells.Count - 2].Value;
                    }
                    catch (Exception xx)
                    {
                        wSheet.Cells[Currow, r.Cells.Count - 1] = Convert.ToDateTime("1900-01-01");
                    }

                    try
                    {
                        wSheet.Cells[Currow, r.Cells.Count ] = r.Cells[r.Cells.Count - 1].Value;
                    }
                    catch (Exception xx)
                    {
                        wSheet.Cells[Currow, r.Cells.Count ] = Convert.ToDateTime("1900-01-01");
                    }

                    range1 = wSheet.Range[wSheet.Cells[Currow, 1], wSheet.Cells[Currow, dataGridView1.ColumnCount]];
                    Color ccolor = r.DefaultCellStyle.BackColor;
                    range1.Interior.Color = System.Drawing.ColorTranslator.ToOle(ccolor);

                    Currow = Currow + 1;
                }

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[dataGridView1.RowCount + 1, dataGridView1.ColumnCount]];
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;
                fm.SetBorders(range1, 1, true);

                wSheet = null;
                wBook = null;
                range1 = null;
                //   xlApp.Quit();
                xlApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Cursor = Cursors.Default;
            }
        }
    }
}
