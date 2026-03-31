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
    public partial class FormReport : Form
    {
        FormMain fm;
        ConsurmptionForm cf;

        public FormReport(FormMain FM, ConsurmptionForm CF)
        {
            InitializeComponent();

            fm = FM;
            cf = CF;


            if (fm.tdb.tApplication.ToList().Count>0)
            {
                button2.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
            }

            if (fm.tdb.tStorageTask.ToList().Count > 0)
            {
                button1.Enabled = true;
            }
            else
            {
                button1.Enabled = false;
            }
        }

        private void FormReport_FormClosed(object sender, FormClosedEventArgs e)
        {
            cf.Enabled = true;
            GC.Collect();
        }

        private void button3_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            Excel1.Application xlApp;
            Excel1.Workbook wBook;
            Excel1.Worksheet wSheet;
            Excel1.Range range1;
            string fName = Application.StartupPath + @"\Shablons\MZZ.xlsx";
            int CurRow = 2;
            object misValue = System.Reflection.Missing.Value;
            bool fl;

            var Header = fm.tdb.SystemTable.ToList();

            if (Header.Count > 0)
            {
                xlApp = new Excel1.Application();
                xlApp.Visible = true;
                try
                {
                    wBook = xlApp.Workbooks.Open(fName);
                    fl = true;
                }
                catch (Exception xx)
                {
                    wBook = xlApp.Workbooks.Add(misValue);
                    fl = false;
                }
                xlApp.DisplayAlerts = false;
                wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                if (fl == false)
                {
                    wSheet.Cells[1, 1] = "Юр.лицо";
                    wSheet.Cells[1, 2] = "Для объекта";
                    wSheet.Cells[1, 3] = "ЦФО Код";
                    wSheet.Cells[1, 4] = "ЦФО согласования";
                    wSheet.Cells[1, 5] = "Адрес Доставки";
                    wSheet.Cells[1, 6] = "Необходимый срок поставки";
                    wSheet.Cells[1, 7] = "Тип материала";
                    wSheet.Cells[1, 8] = "Товар/Услуга Код";
                    wSheet.Cells[1, 9] = "Кол-во";
                    wSheet.Cells[1, 10] = "Ед.изм";
                    wSheet.Cells[1, 11] = "Цена";
                    wSheet.Cells[1, 12] = "Валюта Код";


                    range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 12]];
                    range1.Cells.Font.Size = 12;
                    range1.Cells.Font.Bold = true;
                    range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                    range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignTop;

                    range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    range1.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    range1.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;
                }

                var data = fm.tdb.tApplication.ToList();

                range1 = wSheet.Range[wSheet.Cells[2, 8], wSheet.Cells[data.Count + 1, 8]];
                range1.NumberFormat = "@";

                foreach (var d in data)
                {
                    wSheet.Cells[CurRow, 1] = Header[0].CompName;
                    wSheet.Cells[CurRow, 2] = Header[0].ObjectCode;
                    wSheet.Cells[CurRow, 3] = Header[0].CFOCode;
                    wSheet.Cells[CurRow, 4] = Header[0].CFOSolut;
                    wSheet.Cells[CurRow, 5] = Header[0].Address;
                    wSheet.Cells[CurRow, 6] = d.DateWork.ToString("dd.MM.yyyy");

                    /*  var MG = fm.mList.Where(x => x.MaterialCode == d.MaterialCode).ToList();
                      if (MG.Count > 0)
                      {
                          wSheet.Cells[CurRow, 7] = MG[0].MaterialGroup;
                      }
                      else
                      {*/
                    wSheet.Cells[CurRow, 7] = "Материал";
                    //}
                    wSheet.Cells[CurRow, 8] = d.MaterialCode;
                    wSheet.Cells[CurRow, 9] = d.Quantity;
                    wSheet.Cells[CurRow, 11] = 10;

                    CurRow = CurRow + 1;
                }

                if (fl == false)
                {
                    range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[CurRow, 12]];
                    range1.EntireColumn.AutoFit();
                }

                wSheet = null;
                wBook = null;
                xlApp = null;
            }
            else
            {
                MessageBox.Show("Отсутствуют данные! Меню-> Данные->Настройки заявок", "Сообщение системы");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            var Stor = fm.tdb.tStorageTask.ToList();

            if (Stor.Count > 0)
            {
                Cursor = Cursors.WaitCursor;

                Excel1.Application xlApp;
                Excel1.Workbook wBook;
                Excel1.Worksheet wSheet;
                Excel1.Range range1;
                string fName = Application.StartupPath + @"\Shablons\SBP.xlsm";
                int CurRow = 2;
                object misValue = System.Reflection.Missing.Value;
                bool fl;

                int CurrRow = 6;

                xlApp = new Excel1.Application();
                xlApp.Visible = true;
                try
                {
                    wBook = xlApp.Workbooks.Open(fName);
                    fl = true;
                }
                catch (Exception xx)
                {
                    wBook = xlApp.Workbooks.Add(misValue);
                    fl = false;
                }
                xlApp.DisplayAlerts = false;
                wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                range1 = wSheet.Range[wSheet.Cells[CurrRow, 1], wSheet.Cells[CurrRow + Stor.Count - 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i<Stor.Count; i++)
                {
                    wSheet.Cells[CurrRow, 1] = Stor[i].MaterialCode;
                    wSheet.Cells[CurrRow, 3] = DateTime.Now.ToString("HH:mm");
                    wSheet.Cells[CurrRow, 4] = Stor[i].Source;
                    wSheet.Cells[CurrRow, 5] = Stor[i].Dest;
                    wSheet.Cells[CurrRow, 6] = Stor[i].Quantity;

                    CurrRow = CurrRow + 1;
                }

                CurrRow = CurrRow - 1;

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[CurrRow , 6]];
                fm.SetBorders(range1, 1, true);

                Cursor = Cursors.Default;
            }
            else
            {
                MessageBox.Show("", "Сообщение системы");
            }
        }
    }
}
