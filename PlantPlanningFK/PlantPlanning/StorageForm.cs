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
    public partial class StorageForm : Form
    {
        FormMain fm;

        BindingSource bs;
        DataTable dt1;

        public StorageForm(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            dt1 = new DataTable();
            LoadData();
        }

        private void LoadData()
        {
            Cursor = Cursors.WaitCursor;

            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();

            dt1.Columns.Clear();
            dt1.Rows.Clear();

            dt1.Columns.Add("Склад");
            dt1.Columns.Add("Код");
            dt1.Columns.Add("Наименование");
            dt1.Columns.Add("Мат. группа");
            dt1.Columns.Add("Лот");
            dt1.Columns.Add("Описание лота");
            dt1.Columns.Add("Количество");
            dt1.Columns.Add("Владелец");
            dt1.Columns.Add("Тест качества");
            dt1.Columns.Add("Тип склада");
            dt1.Columns.Add("Тест качества группа");
            dt1.Columns.Add("Тест на складе");
            dt1.Columns.Add("Дата производства");
            dt1.Columns.Add("Срок годности");
            dt1.Columns.Add("dpr");
            dt1.Columns.Add("bbf");
            dt1.Columns.Add("difdates");
            dt1.Columns.Add("difcurdates");

            List<string> StorageNames = fm.StorageNamesList;        // fm.MesLocList.Select(x => x.LocationName).Distinct().ToList();
            var StockList = fm.tdb.StockBalancesMes.Where(x => x.DT == DateTime.Today.Date && (x.WorkDate.Hour >= 6 || x.WorkDate.Hour <= 10) && x.MaterialCode == "1031004635").ToList();

         //   var StockList = fm.edb.f_spMaterialLot_GetAllByWarehouseAndZoneFromNav_IT6WithoutCommentMody().ToList();
            StockList = StockList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();

            string[] NewRow;
            string ProductCode = "";
            string ProductName = "";
            double TotalQuantity = 0;
            double Quantity = 0;
            DateTime ProdDate;
            DateTime BBFDate;
            Int64 dpr;
            Int64 bbf;
            Int64 difdates;
            Int64 difcurdates;

            for (int i = 0; i < StockList.Count; i++)
            {
                if (ProductCode != StockList[i].MaterialCode)
                {//insert row

                    if (ProductCode != "")
                    {
                        NewRow = new string[]
                        {"Итого",
                        ProductCode,
                        ProductName,
                        "",
                        "",
                        "",
                        TotalQuantity.ToString("0")
                        };
                        dt1.Rows.Add(NewRow);
                    }

                    TotalQuantity = 0;
                    ProductCode = StockList[i].MaterialCode;
                    ProductName = StockList[i].MaterialName;
                }

                Quantity = (double)StockList[i].Quantity;
                TotalQuantity = TotalQuantity + Quantity;
                ProdDate = (DateTime)StockList[i].ProdDate;
                BBFDate = (DateTime)StockList[i].BBFDate;
                dpr = (Int64)StockList[i].dpr;
                bbf = (Int64)StockList[i].bbf;
                difdates = (Int64)StockList[i].difdates;
                difcurdates = (Int64)StockList[i].difcurdates;

                NewRow = new string[]
                { StockList[i].Storage,
                  StockList[i].MaterialCode,
                  StockList[i].MaterialName,
                  StockList[i].MatGroup,
                  StockList[i].LotName,
                  StockList[i].LotDescription,
                  Quantity.ToString("0.0"),
                  StockList[i].MaterialOwner,
                  StockList[i].TestQuality,
                  StockList[i].StorageType,
                  StockList[i].TestQualityGr,
                  StockList[i].TestOnStorage,
                  ProdDate.ToString("dd.MM.yyyy"),
                  BBFDate.ToString("dd.MM.yyyy"),
                  dpr.ToString(),
                  bbf.ToString(),
                  difdates.ToString(),
                  difcurdates.ToString()
                };
                dt1.Rows.Add(NewRow);
            }

            if (ProductCode != "")
            {
                NewRow = new string[]
                {"Итого",
                        ProductCode,
                        ProductName,
                        "",
                        "",
                        "",
                        TotalQuantity.ToString("0")
                };
                dt1.Rows.Add(NewRow);
            }

            bs = new BindingSource();
            bs.DataSource = dt1;
            dataGridView1.DataSource = bs;

            dataGridView1.Columns[0].Width = 200;
            dataGridView1.Columns[1].Width = 120;
            dataGridView1.Columns[2].Width = 350;
            dataGridView1.Columns[4].Width = 150;
            dataGridView1.Columns[5].Width = 350;
            dataGridView1.Columns[8].Width = 150;
            dataGridView1.Columns[10].Width = 150;
            dataGridView1.Columns[11].Width = 150;


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            Cursor = Cursors.Default;
        }

        private void StorageForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            if (dataGridView1.RowCount > 0)
            {
                Cursor = Cursors.WaitCursor;
                Excel1.Application xlApp = new Excel1.Application();
                object misValue = System.Reflection.Missing.Value;
                xlApp.Visible = true;
                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)xlApp.Worksheets.get_Item(1);
                Excel1.Range range1;

                for (int i = 0; i < dt1.Columns.Count; i++)
                {
                    wSheet.Cells[1, i + 1] = dt1.Columns[i].ColumnName;
                    range1 = wSheet.Cells[1, i + 1] as Excel1.Range;
                    range1.ColumnWidth = dataGridView1.Columns[i].Width / 8;
                }

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dt1.Columns.Count]];
                range1.Interior.Color = System.Drawing.ColorTranslator.ToOle(Color.Silver);
                //   range1.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;

                for (int i = 0; i < dataGridView1.Rows.Count; i++)
                {
                    for (int j = 0; j < dataGridView1.ColumnCount; j++)
                    {
                        wSheet.Cells[i + 2, j + 1] = dataGridView1.Rows[i].Cells[j].Value.ToString();
                    }
                }

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[dataGridView1.Rows.Count + 1, dataGridView1.Columns.Count]];
                range1.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;
                range1.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel1.XlLineStyle.xlContinuous;

                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].Weight = 4;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].Weight = 4;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].Weight = 4;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;
                range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].Weight = 4;
                range1.WrapText = true;

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();

                Cursor = Cursors.Default;
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            tbLotCode.Text = "";
            tbLotName.Text = "";
            tbProductCode.Text = "";
            tbProductName.Text = "";
        }

        private void tbProductCode_TextChanged(object sender, EventArgs e)
        {
            bs.Filter = "";

            if (tbProductCode.Text.Trim() != "")
            {
                bs.Filter = " [Код] like '%" + tbProductCode.Text.Trim() + "%' ";
            }

            if (tbProductName.Text.Trim() != "")
            {
                if (bs.Filter.Length > 0)
                {
                    bs.Filter = bs.Filter + " and  [Наименование] like '%" + tbProductName.Text.Trim() + "%' ";
                }
                else
                {
                    bs.Filter = " [Наименование] like '%" + tbProductName.Text.Trim() + "%' ";
                }
            }

            if (tbLotCode.Text.Trim() != "")
            {
                if (bs.Filter.Length > 0)
                {
                    bs.Filter = bs.Filter + " and  [Лот] like '%" + tbLotCode.Text.Trim() + "%' ";
                }
                else
                {
                    bs.Filter = " [Лот] like '%" + tbLotCode.Text.Trim() + "%' ";
                }
            }

            if (tbLotName.Text.Trim() != "")
            {
                if (bs.Filter.Length > 0)
                {
                    bs.Filter = bs.Filter + " and  [Описание лота] like '%" + tbLotName.Text.Trim() + "%' ";
                }
                else
                {
                    bs.Filter = " [Описание лота] like '%" + tbLotName.Text.Trim() + "%' ";
                }
            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
             Font f1 = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Bold);

            foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                if (r.Cells[0].Value.ToString().Contains("Итого"))
                {
                    r.DefaultCellStyle.Font = f1;
                }

            }
        }

        private void StorageForm_Resize(object sender, EventArgs e)
        {
            button2.Left = (this.Width - button2.Width) / 2;
        }
    }
}
