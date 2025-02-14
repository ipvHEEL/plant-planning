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
using Excel1 = Microsoft.Office.Interop.Excel;

namespace PlantPlanning
{
    public partial class FormNavData : Form
    {
        FormMain fm;
        BindingSource bs;

        List<fn_select_StockBalancesWithComments_Result> ResList;
        List<StockBalancesShort> ShortResList;
        List<StockBalancesDetails> ResListDetails;

        DataTable dt;

        public FormNavData(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            textBoxCode.Text = "";
            textBoxNaim.Text = "";
            bs = new BindingSource();
            checkBox1.Checked = true;
            comboBox1.Text = comboBox1.Items[0].ToString();
            LoadData();
          //  PrepareData();
            ShowData();
        }

        private void PrepareData()
        {
     

            /*


            */
        }

        private void ShowData()
        {
            bs.DataSource = ShortResList;

            dataGridViewParent.DataSource = null;
            dataGridViewParent.Columns.Clear();
            dataGridViewParent.DataSource = bs;
            dataGridViewParent.Columns["Name1"].Width = 500;

            dataGridViewParent.Columns["Quantity"].DisplayIndex = 0;
            dataGridViewParent.Columns["Quantity"].DefaultCellStyle.Format="N0";
            dataGridViewParent.Columns["Quantity"].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridViewParent.Columns["Quantity"].HeaderText = "Кол-во";
            dataGridViewParent.Columns["Code"].DisplayIndex = 1;
            dataGridViewParent.Columns["Code"].HeaderText = "Код";
            dataGridViewParent.Columns["Group"].DisplayIndex = 2;
            dataGridViewParent.Columns["Group"].HeaderText= "Класс";
            dataGridViewParent.Columns["Name1"].DisplayIndex = 3;
            dataGridViewParent.Columns["Name1"].HeaderText = "Hазвание";
            dataGridViewParent.Columns["MinStatus"].Visible = false;
            dataGridViewParent.Columns["MinStatusValue"].Visible = false;

            dataGridViewParent.ReadOnly = true;
            dataGridViewParent.AllowUserToAddRows = false;
        }

        private void LoadData()
        {
            Cursor = Cursors.WaitCursor;
            ResList = new List<fn_select_StockBalancesWithComments_Result>();
            if (checkBox1.Checked == false)
            {
                ResList = fm.tdb.fn_select_StockBalancesWithComments().ToList();
            }
            else
            {
                ResList = fm.tdb.fn_select_StockBalancesWithComments().Where(x=>x.Storage.ToUpper().Contains("NEW")).ToList();
            }

            ShortResList = new List<StockBalancesShort>();

            ShortResList = ResList.GroupBy(x => new { x.MaterialCode, x.MatGroup, x.MaterialName }).
                    Select(g => new StockBalancesShort
                    {
                        Quantity =Convert.ToInt32(Math.Round( g.Sum(x => (double)x.Quantity),0)),
                        Code = g.Key.MaterialCode,
                        Group = g.Key.MatGroup,
                        Name1 = g.Key.MaterialName
                    }).ToList();

            foreach (var s in ShortResList)
            {
                s.SBDList = new List<StockBalancesDetails>();

                var F = ResList.Where(x => (x.MaterialCode == s.Code) && (x.MatGroup == s.Group)).ToList();

                foreach (var f in F)
                {
                    StockBalancesDetails sbd = new StockBalancesDetails
                    {
                        Lot = f.LotName,
                        Quantity = Convert.ToInt32(Math.Round((double)f.Quantity, 0)),
                        Storage = f.Storage + "(" + f.StorageType + ")",
                        ProdDate = ((DateTime)f.ProdDate).ToString("dd.MM.yyyy") + @"/" + ((DateTime)f.BBFDate).ToString("dd.MM.yyyy"),
                        Status = f.st,
                        TestQuality = f.TestQuality + "(" + f.TestOnStorage + ")",
                        Comment = f.Comment,
                        difDates = (Int32)f.difdates,
                        difCurDates = (Int32)f.difcurdates
                    };

                    if (((f.dpr + f.bbf) == 2) && (f.difdates > 0))
                    {
                        if ((f.difcurdates>0) &&(f.difcurdates<=10))
                        {
                            sbd.Difference = 0;
                        }
                        else if (f.difcurdates <= 0)
                        {
                            sbd.Difference = 0;
                        }
                        else if (((f.difcurdates*1.0/(f.difdates*1.0))>0)&& ((f.difcurdates * 1.0 / (f.difdates * 1.0)) <0.3333334))
                        {
                            sbd.Difference = 0.33;
                        }
                        else
                        {
                            sbd.Difference = 1;
                        }
                    }
                    else
                    {
                        sbd.Difference = -1;
                    }

                    s.SBDList.Add(sbd);
                }

                s.MinStatusValue = s.SBDList.Min(x => x.Difference);

                if (s.MinStatusValue<0)
                {
                    s.MinStatus = "Black";
                }
                else if (s.MinStatusValue==0)
                {
                    s.MinStatus = "Red";
                }
                else if (s.MinStatusValue< 0.3333334)
                {
                    s.MinStatus = "DarkOrange";
                }
                else
                {
                    s.MinStatus = "DarkGreen";
                }
            }

            if (comboBox1.Text == "Красный")
            {
                ShortResList = ShortResList.Where(x => x.MinStatus == "Red").ToList();
            }
            else if (comboBox1.Text == "Черный")
            {
                ShortResList = ShortResList.Where(x => x.MinStatus == "Black").ToList();
            }
            else if (comboBox1.Text == "Зеленый")
            {
                ShortResList = ShortResList.Where(x => x.MinStatus == "DarkGreen").ToList();
            }
            else if (comboBox1.Text == "Желтый")
            {
                ShortResList = ShortResList.Where(x => x.MinStatus == "DarkOrange").ToList();
            }
            else  //all colors doing nothing
            { }

            if (textBoxCode.Text.Trim() != "")
            {
                ShortResList = ShortResList.Where(x => x.Code.ToLower().Contains(textBoxCode.Text.Trim().ToLower())).ToList();
            }

            if (textBoxNaim.Text.Trim() != "")
            {
                ShortResList = ShortResList.Where(x => x.Name1.ToLower().Contains(textBoxNaim.Text.Trim().ToLower())).ToList();
            }


            Cursor = Cursors.Default;
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void FormNavData_FormClosed(object sender, FormClosedEventArgs e)
        {
            fm.Enabled = true;
        }

        private void textBoxCode_TextChanged(object sender, EventArgs e)
        {
            LoadData();
            ShowData();


            /* bs.Filter = "";
           //  dataGridViewChild.DataSource = null;
           //  dataGridViewChild.Columns.Clear();

             if (textBoxCode.Text.Trim()!="")
             {
                 bs.Filter = " Code like '%" + textBoxCode.Text.Trim() + "%' ";
             }

             if (textBoxNaim.Text.Trim()!="")
             {
                 if (bs.Filter.Length>0)
                 {
                     bs.Filter =bs.Filter+ " and Name1 like '%" + textBoxNaim.Text.Trim() + "%' ";
                 }
                 else
                 {
                     bs.Filter = " Name1 like '%" + textBoxNaim.Text.Trim() + "%' ";
                 }
             }
           //  dataGridViewChild.DataSource = bs;*/
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            LoadData();
            ShowData();
        }

        private void dataGridViewParent_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewChild.DataSource = null;
            dataGridViewChild.Columns.Clear();

            ResListDetails = new List<StockBalancesDetails>();

            if (dataGridViewParent.RowCount > 0)
            {
                string Code = dataGridViewParent.CurrentRow.Cells["Code"].Value.ToString();
              //  string Name1 = dataGridViewParent.CurrentRow.Cells["Name1"].Value.ToString();
                string Group = dataGridViewParent.CurrentRow.Cells["Group"].Value.ToString();

                var F = ResList.Where(x => (x.MaterialCode == Code) &&(x.MatGroup==Group)).ToList();

                foreach (var f in F)
                {
                    StockBalancesDetails sbd = new StockBalancesDetails
                    {
                        Lot = f.LotName,
                        Quantity =Convert.ToInt32(Math.Round( (double)f.Quantity,0)),
                        Storage = f.Storage + "(" + f.StorageType + ")",
                        ProdDate = ((DateTime)f.ProdDate).ToString("dd.MM.yyyy") + @"/" + ((DateTime)f.BBFDate).ToString("dd.MM.yyyy"),
                        Status = f.st,
                        TestQuality = f.TestQuality + "(" + f.TestOnStorage + ")",
                        Comment = f.Comment
                    };
                    ResListDetails.Add(sbd);
                }
            }

            dataGridViewChild.DataSource = ResListDetails;
            dataGridViewChild.Columns[0].HeaderText = "Партия";
            dataGridViewChild.Columns[0].Width = 120;
            dataGridViewChild.Columns[1].HeaderText = "Кол-во";
            dataGridViewChild.Columns[1].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dataGridViewChild.Columns[1].DefaultCellStyle.Format = "N0";
            dataGridViewChild.Columns[2].HeaderText = "Склад/тип";
            dataGridViewChild.Columns[2].Width = 250;
            dataGridViewChild.Columns[3].HeaderText = "Дата произ-ва/Срок годности";
            dataGridViewChild.Columns[3].Width = 150;
            dataGridViewChild.Columns[4].HeaderText = "Статус";
            dataGridViewChild.Columns[5].HeaderText = "Тест качества";
            dataGridViewChild.Columns[5].Width = 150;
            dataGridViewChild.Columns[6].HeaderText = "Комментарий";
            dataGridViewChild.Columns[6].Width = 300;
            dataGridViewChild.Columns[7].Visible = false;
            dataGridViewChild.Columns[8].Visible = false;
            dataGridViewChild.Columns[9].Visible = false;

            dataGridViewChild.ReadOnly = true;
            dataGridViewChild.AllowUserToAddRows = false;
        }

        private void dataGridViewParent_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (dataGridViewParent.RowCount>0)
            {
                foreach (DataGridViewRow r in dataGridViewParent.Rows)
                {
                    r.DefaultCellStyle.ForeColor = Color.FromName(r.Cells["MinStatus"].Value.ToString());
                }
            }
        }

        private void dataGridViewChild_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (dataGridViewChild.RowCount>0)
            {
                foreach (DataGridViewRow r in dataGridViewChild.Rows)
                {
                    if (r.Cells[4].Value.ToString()=="OK!")
                    {
                        r.DefaultCellStyle.ForeColor = Color.DarkGreen;
                    }
                    else if (r.Cells[4].Value.ToString() == @"<1/3")
                    {
                        r.DefaultCellStyle.ForeColor = Color.DarkOrange;
                    }
                    else if ((r.Cells[4].Value.ToString() == "<10")||(r.Cells[4].Value.ToString() == "ПР!"))
                    {
                        r.DefaultCellStyle.ForeColor = Color.Red;
                    }
                    else
                    {
                        r.DefaultCellStyle.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void comboBox1_SelectedIndexChanged(object sender, EventArgs e)
        {
            LoadData();
            ShowData();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            var rList = fm.tdb.fn_select_StockBalancesWithComments().ToList();

            if (rList.Count > 0)
            {
                Cursor = Cursors.WaitCursor;

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];
                wSheet.Name = "Лист1";
                Excel1.Range range1;

                int CurRow = 2;

                //Headers
                wSheet.Cells[1, 1] = "Склад";
                wSheet.Cells[1, 2] = "Код";
                wSheet.Cells[1, 3] = "Наименование";
                wSheet.Cells[1, 4] = "MatGroup";
                wSheet.Cells[1, 5] = "Лот";
                wSheet.Cells[1, 6] = "LotDescription";
                wSheet.Cells[1, 7] = "Количество";
                wSheet.Cells[1, 8] = "Владелец";
                wSheet.Cells[1, 9] = "ТестКачества";
                wSheet.Cells[1, 10] = "ТипСклада";
                wSheet.Cells[1, 11] = "ТестКачестваГр";
                wSheet.Cells[1, 12] = "ТестНаСкладе";
                wSheet.Cells[1, 13] = "ProdDate";
                wSheet.Cells[1, 14] = "BBFDate";
                wSheet.Cells[1, 15] = "dpr";
                wSheet.Cells[1, 16] = "bbf";
                wSheet.Cells[1, 17] = "difdates";
                wSheet.Cells[1, 18] = "difcurdates";
                wSheet.Cells[1, 19] = "st";
                wSheet.Cells[1, 20] = "MaterialCode";
                wSheet.Cells[1, 21] = "LotNo";
                wSheet.Cells[1, 22] = "Comment";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 22]];
                range1.Font.Bold = true;
                range1.HorizontalAlignment = HorizontalAlignment.Center;

                foreach (var r in rList)
                {
                    wSheet.Cells[CurRow, 1] = r.Storage;
                    wSheet.Cells[CurRow, 2] = r.MaterialCode;
                    wSheet.Cells[CurRow, 3] = r.MaterialName;
                    wSheet.Cells[CurRow, 4] = r.MatGroup;
                    wSheet.Cells[CurRow, 5] = r.LotName;
                    wSheet.Cells[CurRow, 6] = r.LotDescription;
                    wSheet.Cells[CurRow, 7] =(double) r.Quantity;
                    wSheet.Cells[CurRow, 8] = r.MaterialOwner;
                    wSheet.Cells[CurRow, 9] = r.TestQuality;
                    wSheet.Cells[CurRow, 10] = r.StorageType;
                    wSheet.Cells[CurRow, 11] = r.TestQualityGr;
                    wSheet.Cells[CurRow, 12] = r.TestOnStorage;
                    try
                    {
                        wSheet.Cells[CurRow, 13] = (DateTime)r.ProdDate;
                    }
                    catch (Exception xx)
                    {                    }
                    try
                    {
                        wSheet.Cells[CurRow, 14] = (DateTime)r.BBFDate;
                    }
                    catch (Exception xx)
                    { }
                    wSheet.Cells[CurRow, 15] = (int) r.dpr;
                    wSheet.Cells[CurRow, 16] = (int)r.bbf;
                    wSheet.Cells[CurRow, 17] = (int)r.difdates;
                    wSheet.Cells[CurRow, 18] = (int)r.difcurdates;
                    wSheet.Cells[CurRow, 19] = r.st;
                    wSheet.Cells[CurRow, 20] = r.MatCode;
                    wSheet.Cells[CurRow, 21] = r.LotNo;
                    wSheet.Cells[CurRow, 22] = r.Comment;

                    CurRow = CurRow + 1;
                }
                CurRow = CurRow - 1;

                range1= wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CurRow, 1]];
                range1.NumberFormat = "@";
                range1 = wSheet.Range[wSheet.Cells[2, 20], wSheet.Cells[CurRow, 20]];
                range1.NumberFormat = "@";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[CurRow, 22]];
                range1.EntireColumn.AutoFit();

                wSheet = null;
                wBook = null;
                range1 = null;
                //   xlApp.Quit();
                xlApp = null;
            }
            else
            {
                MessageBox.Show("Данные отсутсвуют в базе!","Сообщение системы");
            }

            Cursor = Cursors.Default;
            GC.Collect();
            MessageBox.Show("Данные успешно выгружены");
        }
    }
}
