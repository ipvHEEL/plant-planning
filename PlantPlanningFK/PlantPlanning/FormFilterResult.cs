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
using PlantPlanning.Context;

namespace PlantPlanning
{
    public partial class FormFilterResult : Form
    {
        ConsurmptionForm cf;
        FormMain fm;
        List<Material> CFMList;
        DataTable dt1;
        List<DailyResult> MaterialData;
        BindingSource bs;
        DateTime LimitDate;

        public Int32 AxX;
        public Int32 AxY;

        public List<string> CFMListGold;
        public List<string> CFMListGray;
        public List<Colorres> CFMListMinus;
        public List<string> CFMListRed;
        public List<Colorres> CFMListPlus;
        public List<string> CFMListGreen;
        public List<Colorres> CFMListLine;
        public List<string> CFMListPink;
        public List<Colorres> CFMListLinePlus;
        public List<string> CFMListGoldens;

        public List<UserSelectedCell> USCList;

        string CurrCode = "";

        Int32 RowInd = 0;
        Int32 ColInd = 0;

        public FormFilterResult(ConsurmptionForm CF, FormMain FM, DateTime lDate)
        {
            InitializeComponent();
            fm = FM;
            cf = CF;
            CFMList = new List<Material>();
            this.Text = "Компоненты " + cf.comboBoxFiltering.Text;

            cbMaterialDelay.Checked = cf.cbMaterialDelay.Checked;
            cbUseMaterialPlanning.Checked = cf.cbUseMaterialPlanning.Checked;
            cbUsingMaterial.Checked = cf.cbUsingMaterial.Checked;
            cbAddDay.Checked = cf.cbAddDay.Checked;
            cbDontShowAfter.Checked = cf.cbDontShowAfter.Checked;
            cbUsingZero.Checked = cf.cbUsingZero.Checked;
            cbViews.Text = cf.cbViews.Text;
            LimitDate = lDate;

            USCList = new List<UserSelectedCell>();

            LoadData();
            ShowData();
            //  DrawDGV();
        }

        private void DrawDGV()
        {
          
        }

        private void LoadData()
        {
            button1_a.BackColor = Color.LightGray; //    Color.DarkGray; #FFFF00 
            button1_b.BackColor = Color.LightPink;    // Color.LightPink;
            button1_c.BackColor = Color.LightYellow;      //Color.Gold;
            button1_d.BackColor = Color.LightCyan;     // Color.SteelBlue;
            button1_e.BackColor = Color.LightGreen;    //Color.Lime;
            button1_f.BackColor = Color.RosyBrown;
            button1_g.BackColor = Color.MistyRose;

            List<string> MatCodeList = cf.FilteredMaterialDataOut.Select(x => x.ProductCode).ToList();
            MaterialData = new List<DailyResult>();

            CFMList = cf.CFMList.Where(x => MatCodeList.Contains(x.MaterialCode)).ToList();
        }

        private void ShowData()
        {
            Cursor = Cursors.WaitCursor;

            //очистить таблицы
            dgvPlanned.Rows.Clear();
            dgvRecipes.Rows.Clear();
            dgvFactis.Rows.Clear();
            dgvStorages.Rows.Clear();
            dgvComment.DataSource = null;
            dgvComment.Rows.Clear();

            dataGridViewMain.DataSource = null;
            dataGridViewMain.Rows.Clear();
            dataGridViewMain.Columns.Clear();
            bs = new BindingSource();
            dt1 = new DataTable();

            CurrCode = "";

            string[] NewString;
            char RowSplitter = '|';
            string sVal = "";
            DateTime CurrDate;
            bool fl;
            Int32 J;
            double Value = 0;
            Int32 Cols;
            List<DateTime> DTList1 = new List<DateTime>();
            List<string> AskList = new List<string>();
            DTList1 = cf.DTList;
            if (cbDontShowAfter.Checked == true)  //cf.
            {
                var XLD = fm.ExcelDataList[cf.comboBoxFiltering.SelectedIndex];
                //  DateTime LimitDate = XLD.DateWork;
                DTList1 = (from d in DTList1
                           where d <= LimitDate
                           select d).ToList();
            }

            CFMListGold = new List<string>();
            CFMListGray = new List<string>();
            CFMListLine = new List<Colorres>(); //[
            CFMListPink = new List<string>();//[
            CFMListPlus = new List<Colorres>(); //+
            CFMListGreen = new List<string>();//+
            CFMListMinus = new List<Colorres>();//-
            CFMListRed = new List<string>(); //-
            CFMListLinePlus = new List<Colorres>();  //[+
            CFMListGoldens = new List<string>();  //[+
            USCList = new List<UserSelectedCell>();

            dt1.Columns.Clear();
            dt1.Rows.Clear();

            dt1.Columns.Add("Код материала");  //0
            dt1.Columns.Add("Наименование");  //1
            dt1.Columns.Add("Класс материала"); //2
            dt1.Columns.Add("Кратность поставки");  //3
            dt1.Columns.Add("Мин срок пост. дней");//4
            dt1.Columns.Add("Страховой запас");//5
            dt1.Columns.Add("Исполнитель");//6
            dt1.Columns.Add("Поставщик");//7
            //Month need insert
            foreach (var ddd in cf.WorkMonthList)
            {
                dt1.Columns.Add("Потребность на " + ddd.ToString("MM.yyyy"));
            }
            dt1.Columns.Add("Уже потреблено");//8
            dt1.Columns.Add("Остатки склады");//9
            dt1.Columns.Add("Остатки пр-во");//10
            dt1.Columns.Add("Внешние остатки");//11
            dt1.Columns.Add("Остатки по НАВ");//12
            dt1.Columns.Add("Заблокировано: Внешние");//13
            dt1.Columns.Add("Заблокировано: МЕС");//14
            dt1.Columns.Add("Старые заявки НАВ");//15

            CurrDate = DateTime.Today.Date;

            if (cbViews.Text == "Остаток")
            {
                foreach (DateTime dt in DTList1)
                {
                    dt1.Columns.Add("Остаток на " + dt.ToString("dd.MM.yyyy"));
                }
            }
            else if (cbViews.Text == "Потребление")
            {
                foreach (DateTime dt in DTList1)
                {
                    dt1.Columns.Add("Потребление на " + dt.ToString("dd.MM.yyyy"));
                }
            }
            else if (cbViews.Text == "Дефицит")
            {
                foreach (DateTime dt in DTList1)
                {
                    dt1.Columns.Add("Дефицит на " + dt.ToString("dd.MM.yyyy"));
                }
            }
            dt1.Columns.Add("Zerrrro");

            MaterialData = CFMList.Select(x => new DailyResult
            {
                ProductCode = x.MaterialCode,
                LineName = x.MaterialName,
                Quantity = 1,
                MaterialType = x.MaterialGroup.Trim()
            }).ToList();
            MaterialData = MaterialData.Distinct().ToList();

            for (int i = 0; i < CFMList.Count; i++)
            {
                sVal = CFMList[i].MaterialCode + RowSplitter + CFMList[i].MaterialName + RowSplitter +
                         CFMList[i].MaterialGroup + RowSplitter + CFMList[i].MaterialMultiplyString + RowSplitter +
                         CFMList[i].WaitingDaysString + RowSplitter + CFMList[i].StorageQuantityString + RowSplitter +
                         CFMList[i].Responsible + RowSplitter + CFMList[i].Sender + RowSplitter;

                for (int j = 0; j < cf.WorkMonthList.Count; j++)
                {
                    if ((Math.Abs(CFMList[i].OutList[j].Quantity) <= 1) && (Math.Abs(CFMList[i].OutList[j].Quantity) > 0))
                    {
                        sVal = sVal + CFMList[i].OutList[j].Quantity.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].OutList[j].Quantity.ToString("N0") + RowSplitter;
                    }
                }

                if ((Math.Abs(CFMList[i].CurrentConsumption) <= 1) && (Math.Abs(CFMList[i].CurrentConsumption) > 0))  //8
                {
                    sVal = sVal + CFMList[i].CurrentConsumption.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + CFMList[i].CurrentConsumption.ToString("N0") + RowSplitter;
                }

                if (cbMaterialDelay.Checked == false)
                {
                    if ((Math.Abs(CFMList[i].RawStorage) <= 1) && (Math.Abs(CFMList[i].RawStorage) > 0))  //9
                    {
                        sVal = sVal + CFMList[i].RawStorage.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawStorage.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawEnterprise) <= 1) && (Math.Abs(CFMList[i].RawEnterprise) > 0))  //10
                    {
                        sVal = sVal + CFMList[i].RawEnterprise.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawEnterprise.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawNav1) <= 1) && (Math.Abs(CFMList[i].RawNav1) > 0))  //11
                    {
                        sVal = sVal + CFMList[i].RawNav1.ToString("N1") + RowSplitter; //+ " (";
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawNav1.ToString("N0") + RowSplitter; // + " (";
                    }

                    if ((Math.Abs(CFMList[i].RawNav2) <= 1) && (Math.Abs(CFMList[i].RawNav2) > 0))  //12
                    {
                        sVal = sVal + CFMList[i].RawNav2.ToString("N1")  + RowSplitter;  //+ ") "
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawNav2.ToString("N0")  + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawMesOut1) <= 1) && (Math.Abs(CFMList[i].RawMesOut1) > 0))  //13
                    {
                        sVal = sVal + CFMList[i].RawMesOut1.ToString("N1") + RowSplitter;// + ") ";
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawMesOut1.ToString("N0") + RowSplitter; //+ " (";
                    }

                    if ((Math.Abs(CFMList[i].RawMesOut2) <= 1) && (Math.Abs(CFMList[i].RawMesOut2) > 0))  //14
                    {
                        sVal = sVal + CFMList[i].RawMesOut2.ToString("N1") + RowSplitter;  // + ")" 
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawMesOut2.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].OldTaskNav) <= 1) && (Math.Abs(CFMList[i].OldTaskNav) > 0)) //15
                    {
                        sVal = sVal + CFMList[i].OldTaskNav.ToString("N1") + RowSplitter;   //0-15
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].OldTaskNav.ToString("N0") + RowSplitter;    
                    }
                }
                else
                {
                    if ((Math.Abs(CFMList[i].RawStorageUD) <= 1) && (Math.Abs(CFMList[i].RawStorageUD) > 0))
                    {
                        sVal = sVal + CFMList[i].RawStorageUD.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawStorageUD.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawEnterpriseUD) <= 1) && (Math.Abs(CFMList[i].RawEnterpriseUD) > 0))
                    {
                        sVal = sVal + CFMList[i].RawEnterpriseUD.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawEnterpriseUD.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawNav1UD) <= 1) && (Math.Abs(CFMList[i].RawNav1UD) > 0))
                    {
                        sVal = sVal + CFMList[i].RawNav1UD.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawNav1UD.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawNav2UD) <= 1) && (Math.Abs(CFMList[i].RawNav2UD) > 0))
                    {
                        sVal = sVal + CFMList[i].RawNav2UD.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawNav2UD.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawMesOut2UD) <= 1) && (Math.Abs(CFMList[i].RawMesOut2UD) > 0))
                    {
                        sVal = sVal + CFMList[i].RawMesOut2UD.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawMesOut2UD.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].RawMesOut1UD) <= 1) && (Math.Abs(CFMList[i].RawMesOut1UD) > 0))
                    {
                        sVal = sVal + CFMList[i].RawMesOut1UD.ToString("N1") + +RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].RawMesOut1UD.ToString("N0") + RowSplitter;
                    }

                    if ((Math.Abs(CFMList[i].OldTaskNavUD) <= 1) && (Math.Abs(CFMList[i].OldTaskNavUD) > 0))
                    {
                        sVal = sVal + CFMList[i].OldTaskNavUD.ToString("N1") + RowSplitter;   //0-15
                    }
                    else
                    {
                        sVal = sVal + CFMList[i].OldTaskNavUD.ToString("N0") + RowSplitter;   //0-15
                    }

                }

                if ((CFMList[i].WaitingDaysString == "") || (CFMList[i].MaterialMultiplyString == ""))
                {
                    CFMListGray.Add(CFMList[i].MaterialCode);
                }

                if (cbViews.Text == "Остаток")  //cf.
                {
                    Value = CFMList[i].CurrentConsumption + CFMList[i].RawNav1;

                    if (cbMaterialDelay.Checked == false)
                    {
                        foreach (var sl in CFMList[i].RawStorageList)
                        {
                            Value = Value + sl.Quantity;
                        }

                    /*    foreach (var el in CFMList[i].RawEnterpriseList)
                        {
                            Value = Value + el.Quantity;
                        }*/
                    }
                    else
                    {
                        foreach (var sl in CFMList[i].RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                        {
                            Value = Value + sl.Quantity;
                        }

                       /* foreach (var el in CFMList[i].RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                        {
                            Value = Value + el.Quantity;
                        }*/
                    }

                    Cols = 16 + cf.WorkMonthList.Count;
                    var cListData = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode)).ToList();
                    var xlData = cf.XLDataList.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                    var AppList = fm.tdb.tApplication.ToList();

                    var adata = AppList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();    // fm.tdb.tApplication.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();       // cf.AppList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();

                    foreach (DateTime dt in DTList1)
                    {
                        //income
                        double T2 = 0;
                        double T1 = 0;

                        var FDTList = DTList1.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();

                        if (cbUseMaterialPlanning.Checked == true)
                        {
                            List<Raws> tData2 = new List<Raws>();

                            if (cbAddDay.Checked == false)
                            {
                                tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                            }
                            else
                            {
                                tData2 = CFMList[i].TaskNavList2.Where(x => x.AddDate == dt).ToList();
                            }

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                            {
                                tData2 = CFMList[i].TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                            }

                            foreach (var td2 in tData2)
                            {
                                T2 = T2 + td2.Quantity;
                            }

                            List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                            if (cbAddDay.Checked == false)
                            {
                                Cdata = cListData.Where(x => x.ExpDT.Date == dt.Date).ToList();    //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                            }
                            else
                            {
                                Cdata = cListData.Where(x => x.ExpDT.Date.AddDays(1) == dt.Date).ToList();       //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                            }

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                            {
                                Cdata = cListData.Where(x => ((DateTime)x.AlterDate).Date == dt.Date).ToList();     //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                            }

                            foreach (var cd in Cdata)
                            {
                                T2 = T2 + cd.Quantity;
                            }

                            if (T2 > 0)
                            {
                                if (!AskList.Contains(CFMList[i].MaterialCode))
                                {
                                    AskList.Add(CFMList[i].MaterialCode);
                                }

                                if (Value < 0)
                                {
                                    Value = 0;
                                }
                            }

                            Value = Value + T2;

                            //ask in progress
                            List<Raws> tData1 = new List<Raws>();

                            if (cbAddDay.Checked == false)
                            {
                                tData1 = CFMList[i].TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                            }
                            else
                            {
                                tData1 = CFMList[i].TaskNavList1.Where(x => x.AddDate == dt).ToList();
                            }

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                            {
                                tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                            }

                            foreach (var td1 in tData1)
                            {
                                T1 = T1 + td1.Quantity;
                            }

                            if (T1 > 0)
                            {
                                if (!AskList.Contains(CFMList[i].MaterialCode))
                                {
                                    AskList.Add(CFMList[i].MaterialCode);
                                }

                               /* if (Value < 0)
                                {
                                    Value = 0;
                                }*/
                            }

                         //   Value = Value + T1;
                        }

                        //credit
                        if (cbUsingMaterial.Checked == true)
                        {
                            var dll = xlData.Where(x => x.DateWork.Date == dt).ToList();     // cf.XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == dt)).ToList();
                            foreach (var dl in dll)
                            {
                                Value = Value - dl.Quantity;
                            }
                        }

                        //future planning Asks
                        double Q = 0;
                       // var adata = cf.AppList;    // fm.tdb.tApplication.ToList();
                        //  if (adata.Count > 0)
                       // {
                          //  adata = adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                            //    if (adata.Count > 0)
                         //   {
                              var  bdata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                if ((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                {
                                  //  adata = cf.AppList;
                                  //  adata = adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                                    bdata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                }

                                foreach (var ad in bdata)
                                {
                                    Q = Q + ad.Quantity;
                                }

                                USCList.Add(new UserSelectedCell
                                {
                                    RowIndex = i,
                                    ColIndex = Cols,
                                    Quantity = Q
                                });
                                //  Value = Value + Q;
                        //    }
                      //  }

                        if (Value < 0)
                        {
                          //  if (dt.Subtract(DateTime.Today.Date).Days < 60)
                            {
                                if (AskList.Contains(CFMList[i].MaterialCode))
                                {
                                    for (int k = AskList.Count - 1; k >= 0; k--)
                                    {
                                        if (AskList[k] == CFMList[i].MaterialCode)
                                        {
                                            AskList.RemoveAt(k);
                                            k = 0;
                                        }
                                    }
                                }

                                //  CFMBool[i] = false;
                                if (!CFMListGold.Contains(CFMList[i].MaterialCode))  //-
                                {
                                    CFMListGold.Add(CFMList[i].MaterialCode);
                                    CFMListRed.Add(CFMList[i].MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = CFMList[i].MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListMinus.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListMinus.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                                    try
                                    {
                                        Coll[0].ColNo.Add(Cols);
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = CFMList[i].MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                    }
                                }
                            }
                        }

                        if ((Math.Abs(Value) <= 1) && (Value != 0))
                        {
                            sVal = sVal + Value.ToString("N1") + " ";
                        }
                        else
                        {
                            sVal = sVal + Value.ToString("N0") + " ";
                        }

                        if (T2 != 0)
                        {
                            if ((Math.Abs(T2) <= 1) && (T2 != 0))
                            {
                                sVal = sVal + " (" + T2.ToString("N1") + ") ";
                            }
                            else
                            {
                                sVal = sVal + " (" + T2.ToString("N0") + ") ";
                            }
                            if (!CFMListGreen.Contains(CFMList[i].MaterialCode))  //+
                            {
                                CFMListGreen.Add(CFMList[i].MaterialCode);
                                Colorres Col = new Colorres
                                {
                                    ProdCode = CFMList[i].MaterialCode,
                                    ColNo = new List<int>()
                                };
                                Col.ColNo.Add(Cols);
                                CFMListPlus.Add(Col);
                            }
                            else
                            {
                                var Coll = CFMListPlus.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                                try
                                {
                                    Coll[0].ColNo.Add(Cols); ;
                                }
                                catch (Exception xx)
                                {
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = CFMList[i].MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListPlus.Add(Col);
                                }
                            }
                        }

                        if (T1 != 0)
                        {
                            if ((Math.Abs(T1) <= 1) && (T1 != 0))
                            {
                                sVal = sVal + " [" + T1.ToString("N1") + "] ";
                            }
                            else
                            {
                                sVal = sVal + " [" + T1.ToString("N0") + "] ";
                            }

                            if (!CFMListPink.Contains(CFMList[i].MaterialCode))  //[
                            {
                                CFMListPink.Add(CFMList[i].MaterialCode);
                                Colorres Col = new Colorres
                                {
                                    ProdCode = CFMList[i].MaterialCode,
                                    ColNo = new List<int>()
                                };
                                Col.ColNo.Add(Cols);
                                CFMListLine.Add(Col);
                            }
                            else
                            {
                                var Coll = CFMListLine.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                                try
                                {
                                    Coll[0].ColNo.Add(Cols);
                                }
                                catch (Exception xx)
                                {
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = CFMList[i].MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListLine.Add(Col);
                                }
                            }

                            if (CFMListGreen.Contains(CFMList[i].MaterialCode))  //[+
                            {
                                if (!CFMListGoldens.Contains(CFMList[i].MaterialCode))
                                {
                                    CFMListGoldens.Add(CFMList[i].MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = CFMList[i].MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListLinePlus.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListLinePlus.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                                    try
                                    {
                                        Coll[0].ColNo.Add(Cols);
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = CFMList[i].MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListLinePlus.Add(Col);
                                    }
                                }
                            }
                        }

                        if (Q != 0)
                        {
                            //Q = 0;
                            bdata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();
                            if ((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                            {
                                bdata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                            }

                            foreach (var b in bdata)
                            {
                                if ((Math.Abs(b.Quantity) <= 1) && (b.Quantity != 0))
                                {
                                    sVal = sVal + " {" + b.Quantity.ToString("N1") + "} ";
                                }
                                else
                                {
                                    sVal = sVal + " {" + b.Quantity.ToString("N0") + "} ";
                                }
                            }
                        }

                        sVal = sVal + "" + RowSplitter;
                        Cols = Cols + 1;
                    }
                }  //остаток
                else if (cbViews.Text == "Потребление")
                {
                    var xlData = cf.XLDataList.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();

                    foreach (DateTime dt in DTList1)
                    {
                        var xlD = xlData.Where(x => x.DateWork.Date == dt).ToList();

                        if (xlD.Count > 0)
                        {
                            double summ = xlD.Sum(x => x.Quantity);
                            /*if ((Math.Abs(xlD[0].Quantity) <= 1) && (xlD[0].Quantity != 0))
                            {
                                sVal = sVal + xlD[0].Quantity.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + xlD[0].Quantity.ToString("N0") + RowSplitter;
                            }*/
                            if ((Math.Abs(summ) <= 1) && (summ != 0))
                            {
                                sVal = sVal + summ.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + summ.ToString("N0") + RowSplitter;
                            }
                        }
                        else
                        {
                            sVal = sVal + "" + RowSplitter;
                        }
                    }
                }
                else  //Дефицит
                {
                    Value = CFMList[i].CurrentConsumption + CFMList[i].RawNav1;

                    if (cbMaterialDelay.Checked == false)
                    {
                        foreach (var sl in CFMList[i].RawStorageList)
                        {
                            Value = Value + sl.Quantity;
                        }

                       /* foreach (var el in CFMList[i].RawEnterpriseList)
                        {
                            Value = Value + el.Quantity;
                        }*/
                    }
                    else
                    {
                        foreach (var sl in CFMList[i].RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                        {
                            Value = Value + sl.Quantity;
                        }

                       /* foreach (var el in CFMList[i].RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                        {
                            Value = Value + el.Quantity;
                        }*/
                    }

                    Cols = 16 + cf.WorkMonthList.Count;
                    var cListData = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode)).ToList();
                    var xlData = cf.XLDataList.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                    var adata = fm.tdb.tApplication.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList(); // cf.AppList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();

                    foreach (DateTime dt in DTList1)
                    {
                        var FDTList = DTList1.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();
                        //income
                        List<Raws> tData2 = new List<Raws>();
                        if (cbAddDay.Checked == false)
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                        }
                        else
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.AddDate == dt).ToList();
                        }

                        double T2 = 0;
                        if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                        }

                        foreach (var td2 in tData2)
                        {
                            T2 = T2 + td2.Quantity;
                        }

                        List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                        if (cbAddDay.Checked == false)
                        {
                            Cdata = cListData.Where(x => x.ExpDT.Date == dt.Date).ToList();      //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                        }
                        else
                        {
                            Cdata = cListData.Where(x => x.ExpDT.Date.AddDays(1) == dt.Date).ToList();      //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                        }

                        if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = cListData.Where(x => ((DateTime)x.AlterDate).Date == dt.Date).ToList();     //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                        }

                        foreach (var cd in Cdata)
                        {
                            T2 = T2 + cd.Quantity;
                        }

                        if (T2 > 0)
                        {
                            if (Value < 0)
                            {
                                Value = 0;
                            }

                            if (!AskList.Contains(CFMList[i].MaterialCode))
                            {
                                AskList.Add(CFMList[i].MaterialCode);
                            }
                        }

                        Value = Value + T2;

                        //ask in progress
                        List<Raws> tData1 = new List<Raws>();

                        if (cbAddDay.Checked == false)
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                        }
                        else
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.AddDate == dt).ToList();
                        }
                        double T1 = 0;

                        if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                        }

                        foreach (var td1 in tData1)
                        {
                            T1 = T1 + td1.Quantity;
                        }

                        if (T1 > 0)
                        {
                            if (!AskList.Contains(CFMList[i].MaterialCode))
                            {
                                AskList.Add(CFMList[i].MaterialCode);
                            }

                            /*if (Value < 0)
                            {
                                Value = 0;
                            }*/
                        }

                       // Value = Value + T1;

                        //credit
                        var dll = xlData.Where(x => x.DateWork.Date == dt).ToList();    // cf.XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == dt)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value - dl.Quantity;
                        }

                        //future planning Asks
                        double Q = 0;
                      //  var adata = cf.AppList;    // fm.tdb.tApplication.ToList();
                        //  if (adata.Count > 0)
                        {
                           // adata = adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                            //    if (adata.Count > 0)
                            {
                                var bdata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                if ((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                {
                                   // adata = cf.AppList;
                                   // adata = adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                                    bdata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                }

                                foreach (var ad in bdata)
                                {
                                    Q = Q + ad.Quantity;
                                }

                                USCList.Add(new UserSelectedCell
                                {
                                    RowIndex = i,
                                    ColIndex = Cols,
                                    Quantity = Q
                                });
                                //  Value = Value + Q;
                            }
                        }

                        if (Value < 0)
                        {
                          //   if (dt.Subtract(DateTime.Today.Date).Days < 60)
                            {
                                if (AskList.Contains(CFMList[i].MaterialCode))
                                {
                                    for (int k = AskList.Count - 1; k >= 0; k--)
                                    {
                                        if (AskList[k] == CFMList[i].MaterialCode)
                                        {
                                            AskList.RemoveAt(k);
                                            k = 0;
                                        }
                                    }
                                }

                                //  CFMBool[i] = false;
                                if (!CFMListGold.Contains(CFMList[i].MaterialCode))  //-
                                {
                                    CFMListGold.Add(CFMList[i].MaterialCode);
                                    CFMListRed.Add(CFMList[i].MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = CFMList[i].MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListMinus.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListMinus.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                                    try
                                    {
                                        Coll[0].ColNo.Add(Cols);
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = CFMList[i].MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                    }
                                }
                            }

                            if ((Math.Abs(Value) <= 1) && (Value != 0))
                            {
                                sVal = sVal + Value.ToString("N1") + " ";
                            }
                            else
                            {
                                sVal = sVal + Value.ToString("N0") + " ";
                            }
                        }

                        sVal = sVal + "" + RowSplitter;
                        Cols = Cols + 1;
                    }
                }

                NewString = sVal.Split(RowSplitter);
                dt1.Rows.Add(NewString);
            }

            CFMListGold = CFMListGold.Except(AskList).ToList();

            bs = new BindingSource();
            bs.DataSource = dt1;
            dataGridViewMain.DataSource = bs;

            dataGridViewMain.Columns[0].Width = 80;
            dataGridViewMain.Columns[1].Width = 300;
            dataGridViewMain.Columns[2].Width = 120;
            dataGridViewMain.Columns[3].Width = 50;
            dataGridViewMain.Columns[4].Width = 50;
            dataGridViewMain.Columns[5].Width = 50;
            dataGridViewMain.Columns[6].Width = 100;
            dataGridViewMain.Columns[7].Width = 100;
            dataGridViewMain.Columns[dataGridViewMain.ColumnCount - 1].Visible = false;

            for (int i = 8; i < 15; i++)
            {
                dataGridViewMain.Columns[i + cf.WorkMonthList.Count].DefaultCellStyle.BackColor = Color.Gainsboro;
                dataGridViewMain.Columns[i + cf.WorkMonthList.Count].HeaderCell.Style.Font = new Font(dataGridViewMain.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f, FontStyle.Bold);
            }

            dataGridViewMain.Columns[15 + cf.WorkMonthList.Count].HeaderCell.Style.Font = new Font(dataGridViewMain.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f, FontStyle.Bold);

            //   DGV_UpdateColumnZero();
            //   DGV_Get_CellColor();

            for (int i = 0; i < 5; i++)
            {
                dataGridViewMain.Columns[i].Frozen = true;
            }

            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;
            dataGridViewMain.DoubleBuffered = true;
            dataGridViewMain.Update();
            //cellcolor

           /*  foreach (DataGridViewColumn c in dataGridViewMain.Columns)
             {
                 c.SortMode = DataGridViewColumnSortMode.NotSortable;
             }*/

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Cursor = Cursors.Default;
        }

        private void ClearDGV()
        {
            dataGridViewMain.ClearSelection();
            dataGridViewMain.CurrentCell = null;

            dgvFactis.Rows.Clear();
            dgvPlanned.Rows.Clear();
            dgvRecipes.Rows.Clear();
            dgvStorages.Rows.Clear();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void FormFilterResult_FormClosed(object sender, FormClosedEventArgs e)
        {
            cf.FilteredMaterialDataOut.Clear();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            cf.Enabled = true;
        }

        private void FormFilterResult_Shown(object sender, EventArgs e)
        {
            DrawDGV();
        }

        private void dataGridViewMain_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            RowInd = e.RowIndex;
            ColInd = e.ColumnIndex;

            Rectangle rect = dataGridViewMain.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            AxX = rect.X +/* fm.Left +*/ this.Left;
            AxY = rect.Y +/* fm.Top + this.panel1.Height + */this.Top + 80;

            if ((RowInd != -1) /*&& (ColIndex != -1)*/)
            {
                string ProductCode = dataGridViewMain.Rows[RowInd].Cells[0].Value.ToString().Trim();
                DateTime CurrDate = DateTime.Today.Date;
                DateTime DateStart = CurrDate.AddHours(8);
                DateTime DateEnd = DateStart.AddDays(1);
                DateTime PlanDate;
                string State = "";

                List<StockBalancesResult> sbrMesList = new List<StockBalancesResult>();
                List<StockBalancesResult> sbrNavList = new List<StockBalancesResult>();
                List<StockBalancesResult> sbrItogList = new List<StockBalancesResult>();
                List<StockBalancesPlant> sbpList = new List<StockBalancesPlant>();
                List<NavPurchaseTable> sListA = new List<NavPurchaseTable>();
                List<NavPurchaseTable> sListB = new List<NavPurchaseTable>();

                string[] dataString;
                bool fl = false;
                double Quantity = 0;
                double Total = 0;
                string LName = "";
                string CurrStr;
                DateTime dt;
                List<DateTime> DDList = new List<DateTime>();

                if (ColInd > 0)
                {
                    if ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Остаток на")) || (dataGridViewMain.Columns[ColInd].HeaderText.Contains("Потребление на")) || ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Дефицит на"))))
                    {
                        LName = dataGridViewMain.Columns[ColInd].HeaderText;
                        LName = LName.Substring(LName.Length - 11, 11).Trim();
                        fl = true;

                        CurrDate = fm.ConvertDataToDate(LName);
                        DateStart = CurrDate.AddHours(8);
                        DateEnd = DateStart.AddDays(1);
                        Quantity = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[ColInd].Value.ToString().Trim());
                    }
                    else
                    {
                        CurrDate = DateTime.Today.Date;
                    }
                }
                else
                {
                    CurrDate = DateTime.Today.Date;
                }

                var cfm1 = CFMList.Where(x => x.MaterialCode == ProductCode).ToList();

                List<DateTime> DList = new List<DateTime>();
                DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();
                DList = DList.Distinct().OrderBy(x => x.Date).ToList();

                DateTime StartDate1;
                DateTime EndDate1;

                try
                {
                    StartDate1 = DList.Min();
                    EndDate1 = DList.Max();
                }
                catch (Exception xx)
                {
                    StartDate1 = DateTime.Today.Date;
                    EndDate1 = StartDate1.AddDays(1);
                }

                var CurProduct = fm.lRec.Where(x => x.MaterialCode == ProductCode).ToList();
                DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();

                if (CurrCode == ProductCode)
                {  // показать остатки - потребление  в меню

                }
                else
                {
                    //очистить таблицы
                    dgvPlanned.Rows.Clear();
                    dgvRecipes.Rows.Clear();
                    dgvFactis.Rows.Clear();
                    dgvStorages.Rows.Clear();
                    dgvComment.DataSource = null;
                    dgvComment.Rows.Clear();
                    //вывести данные для текущей таблицы
                    switch (tabControl2.SelectedIndex)
                    {
                        case 1:
                            dgvPlanned.Rows.Clear();
                            dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };
                            dgvPlanned.Rows.Add(dataString);
                            var E1 = fm.ExcelDataList.Where(x => DDList.Contains(x.DateWork) && x.Quantity != 0).ToList();

                            for (int i=0; i<DDList.Count; i++) //  foreach (DateTime dt1 in DDList)   //DList
                            {
                                //planned
                                //  if (dt1 >= DateTime.Today.Date)
                                {
                                    fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;  //Cells["Класс материала"]
                                    var ELDListDate = E1.Where(x => (x.DateWork == DDList[i])).ToList();  // fm.ExcelDataList.Where(x => (x.DateWork == DDList[i]) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //dt1

                                     for (int j=0; j<ELDListDate.Count; j++)   //foreach (var ELD in ELDListDate)
                                    {
                                        var CP = CurProduct.Where(x => (x.ProductCode == ELDListDate[j].ProdCode) && (x.LineNumber == ELDListDate[j].Line) && (x.WorkDate <= DDList[i])).ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
                                        fn_select_RecipesView_Result SRR;
                                        try
                                        {
                                            //  CP = CP.Where(x => x.WorkDate <= dt1).OrderByDescending(x => x.WorkDate).ToList();
                                            SRR = CP.OrderByDescending(x => x.WorkDate).First();      // CP = CP.Where(x => x.WorkDate <= dt1).OrderByDescending(x => x.WorkDate).ToList();  //x.Starting_Date
                                        }
                                        catch (Exception xx)
                                        {
                                            SRR = null;
                                        }

                                        if (SRR != null)    // (CP.Count > 0)
                                        {
                                            //раскрутка рецептуры в обратном направлении! Продукт-Линия-совпадает ли дата?
                                            var PRL = fm.ProdRecipeList.Where(x => x.ProductCode == ELDListDate[j].ProdCodeStr).ToList();
                                            if (PRL.Count > 0)
                                            {
                                                var LL = PRL[0].RecLineList.Where(x => x.LineNumber == ELDListDate[j].Line).ToList();
                                                if (LL.Count > 0)
                                                {
                                                    var DL = LL[0].RecList.Where(x => x.WorkDateStart <= DDList[i]).OrderByDescending(x => x.WorkDateStart).ToList();  //CurrDate
                                                    if (DL.Count > 0)
                                                    {
                                                        if (DL[0].WorkDateStart <= SRR.WorkDate)
                                                        {
                                                            if (ELDListDate[j].RePack == true)
                                                            {
                                                                if (fl == true)
                                                                {
                                                                    if ((double)SRR.Quantity > 0)
                                                                    {
                                                                        dataString = new string[] {
                                                                    DDList[i].ToString("dd.MM.yyyy"),  //  CurrDate
                                                                    ELDListDate[j].Line,
                                                                    ELDListDate[j].Quantity.ToString("N0"),
                                                                    ((double) SRR.Quantity*1000).ToString("N0"),   //CP[0].Quantity
                                                                    (((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                                                                    ELDListDate[j].ProdCodeStr,
                                                                    ELDListDate[j].AlterProdName
                                                            };

                                                                        dgvPlanned.Rows.Add(dataString);
                                                                    }
                                                                }
                                                            }
                                                            else
                                                            {
                                                                if ((double)SRR.Quantity > 0)
                                                                {
                                                                    dataString = new string[] {
                                                                DDList[i].ToString("dd.MM.yyyy"),
                                                               ELDListDate[j].Line,
                                                                ELDListDate[j].Quantity.ToString("N0"),
                                                                ((double) SRR.Quantity*1000).ToString("N0"),
                                                                (((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                                                                ELDListDate[j].ProdCodeStr,
                                                                ELDListDate[j].AlterProdName
                                                            };

                                                                    dgvPlanned.Rows.Add(dataString);
                                                                }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    foreach (DataGridViewRow r in dgvPlanned.Rows)
                                    {
                                        if (r.Cells[0].Value.ToString() != "")
                                        {
                                            r.Cells[0].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                            r.Cells[1].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                            r.Cells[2].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                            r.Cells[3].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                        }
                                    }
                                    ELDListDate = null;
                                }                       // dt1
                            }  //foreach

                            break;
                        case 0:

                            dgvFactis.Rows.Clear();
                            dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };
                            dgvFactis.Rows.Add(dataString);
                            {
                                try
                                {
                                    StartDate1 = DList.Min();
                                    EndDate1 = DList.Max();
                                }
                                catch (Exception xx)
                                {
                                    StartDate1 = DateTime.Today.Date;
                                    EndDate1 = StartDate1.AddDays(1);
                                }

                                fl = dataGridViewMain.CurrentRow.Cells[2].Value.ToString().ToLower() == "тара" ? true : false;
                                //  var MuaSection = fm.MUAList.Where(x => (x.DateTime >= DateStart) && (x.DateTime < DateEnd)).ToList();    

                                var FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, StartDate1).OrderBy(x => x.LotForErp).ToList();      //   DateTime.Today.Date.AddYears(-2)               //fm.edb.fn_select_MaterialUsingByProduct(StartDate1, EndDate1, ProductCode).OrderBy(x => x.DateStart).ToList();

                                List<string> LotForErp = FactData.Select(x => x.LotForErp).Distinct().ToList();
                                List<fn_select_ProductionByCode_Result> FactData1 = new List<fn_select_ProductionByCode_Result>();

                                foreach (string lfe in LotForErp)
                                {
                                    var fd = FactData.Where(x => x.LotForErp == lfe).ToList();

                                    if (fd.Count > 0)
                                    {
                                        var fd1 = fd[0];

                                        for (int i = 1; i < fd.Count; i++)
                                        {
                                            fd1.Quantity = fd1.Quantity + fd[i].Quantity;
                                            fd1.RawQuant = fd1.RawQuant + fd[i].RawQuant;
                                        }
                                        fd1.EndDate = fd[fd.Count - 1].EndDate;
                                    }
                                }

                                foreach (var f in FactData)
                                {
                                    // if (f.DateStart >= DateStart)
                                    {
                                        dataString = new string[]
                                        {
                                    ((double)f.RawQuant).ToString("N3"),
                                    ((DateTime)f.StartDate).ToString("dd.MM.yyyy HH;mm"),
                                    ((DateTime)f.EndDate).ToString("dd.MM.yyyy HH:mm"),
                                    f.LotForErp,
                                    ((double)f.Quantity).ToString("N2"),
                                    f.MaterialCode,
                                    f.SystemName+" "+ f.MaterialName
                                        };

                                        dgvFactis.Rows.Add(dataString);
                                    }
                                }

                                foreach (DataGridViewRow r in dgvFactis.Rows)
                                {
                                    if (r.Cells[0].Value.ToString() != "")
                                    {
                                        r.Cells[0].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                        r.Cells[1].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                        r.Cells[2].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                        r.Cells[3].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                    }
                                }

                                FactData = null;
                            }  //fact
                            break;
                        case 2:
                            dgvRecipes.Rows.Clear();
                            //    Quantity = fm.ConvertStringToDouble(dataGridViewMain.CurrentRow.cel.Value.ToString().Trim());
                            var PRListA = fm.ProdRecipeList.Where(x => x.ProductCode == ProductCode).ToList();     //PRList.Where(x => x.ProductCode == ProductCode).ToList();
                            if (PRListA.Count > 0)
                            {
                                dataString = new string[] { "", "Рецепты с разбиением по линиям, действующие с " + CurrDate.ToString("dd.MM.yyyy") + ":" };//+ System.Environment.NewLine + System.Environment.NewLine;
                                dgvRecipes.Rows.Add(dataString);
                            }
                            foreach (var prA in PRListA)
                            {
                                var PRListLines = prA.RecLineList;

                                foreach (var PRL in PRListLines)
                                {
                                    var Recipesl = PRL.RecList.Where(x => (x.WorkDateStart <= CurrDate) && (x.WorkDateEnd >= CurrDate)).ToList();
                                    if (Recipesl.Count > 0)
                                    {
                                        dataString = new string[] { "", PRL.LineNumber };
                                        dgvRecipes.Rows.Add(dataString);
                                    }

                                    foreach (var RL in Recipesl)
                                    {
                                        var RecList = RL.rList;

                                        foreach (var rrr in RecList)
                                        {
                                            dataString = new string[] { rrr.MaterialCode,
                                                                rrr.MaterialName,
                                                                rrr.MaterialQuantity.ToString("N6"),
                                                                (Quantity * rrr.MaterialQuantity).ToString("N3") };
                                            dgvRecipes.Rows.Add(dataString);
                                        }
                                    }
                                }
                            }  //recipes
                            break;
                        case 3:
                            dgvUsing.Rows.Clear();
                            {
                              //  var LR = fm.lRec1.Where(x => x.MaterialCode == ProductCode).ToList();
                                List<fn_select_RecipesView_Result> LResult = new List<fn_select_RecipesView_Result>();
                                List<string> Slist = CurProduct.Select(x => x.LineNumber).Distinct().ToList();

                                foreach (var s in Slist)
                                {
                                    try
                                    {
                                        var LR1 = CurProduct.Where(x => x.LineNumber == s && x.WorkDate <= DateTime.Today.Date).ToList();

                                        StartDate1 = ((DateTime)LR1.Max(x => x.WorkDate));
                                    }
                                    catch (Exception xx)
                                    {
                                        StartDate1 = DateTime.Today.Date;
                                    }

                                    var LR2 = CurProduct.Where(x => x.LineNumber == s && x.WorkDate >= StartDate1).ToList();
                                    LResult.AddRange(LR2);
                                }

                                LResult = LResult.OrderBy(x => x.WorkDate).ThenBy(x => x.LineNumber).ToList();
                                dgvUsing.Rows.Clear();

                                foreach (var lr in LResult)
                                {
                                    if (lr.Quantity > 0)
                                    {
                                        double Q = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[16 + cf.WorkMonthList.Count].Value.ToString().Trim());

                                        dataString = new string[]
                                        {
                                            ((DateTime)lr.WorkDate).ToString("yyyy-MM-dd"),
                                             lr.LineNumber,
                                            ((double)lr.Quantity).ToString("N6"),
                                            ((double)(Q/lr.Quantity)).ToString("N2"),
                                             lr.ProductCode,
                                            lr.ProductName
                                        };
                                    }
                                    else
                                    {
                                        dataString = new string[]
                                        {
                                            ((DateTime)lr.WorkDate).ToString("yyyy-MM-dd"),
                                                lr.LineNumber,
                                                ((double)lr.Quantity).ToString("N6"),
                                                "---",
                                            lr.ProductCode,
                                            lr.ProductName
                                        };
                                    }
                                    dgvUsing.Rows.Add(dataString);
                                }
                            }
                            break;
                        case 4:
                            dgvComment.DataSource = null;
                            dgvComment.Rows.Clear();

                            var ComData = fm.tdb.fn_select_CommentDataID(ProductCode).ToList();
                            //   dgvComment.DataSource = ComData;

                            foreach (var CD in ComData)
                            {
                                dataString = new string[]
                                {
                            CD.IDC.ToString(),
                            CD.MaterialCodeC,
                            CD.MaterialNameC,
                            ((DateTime)CD.DateRecordC).ToString("yyyy-MM-dd"),
                            CD.IsActualC.ToString(),
                            CD.UserName,
                            CD.CommentC
                                };
                                dgvComment.Rows.Add(dataString);
                            }

                            dgvComment.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                            dgvComment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                            break;

                        default:
                            break;
                    }//switch

                    {  //и для складов
                        //storages
                        var ListAA = cfm1[0].RawStorageList;

                        if (cbMaterialDelay.Checked == false)
                        {
                            ListAA = ListAA.Where(x => x.Quantity != 0).ToList();
                        }
                        else
                        {
                            ListAA = ListAA.Where(x => (x.BBFDate >= DateTime.Today.Date) && (x.Quantity != 0)).ToList();
                        }

                        foreach (var aa in ListAA)
                        {
                            if (aa.BBFDate.Year <= 2000)
                            {
                                aa.Color = "Violet";
                            }
                            else
                            {
                                DateTime pBDzate = aa.BBFDate;

                                if ((pBDzate - DateTime.Today.Date).Days > 30)
                                {
                                    aa.Color = "Green";
                                }
                                else if ((pBDzate - DateTime.Today.Date).Days > 20)
                                {
                                    aa.Color = "Magenta";
                                }
                                else if ((pBDzate - DateTime.Today.Date).Days > 0)
                                {
                                    aa.Color = "DarkOrange";
                                }
                                else
                                {
                                    aa.Color = "Red";
                                }
                            }
                        }

                        if (ListAA.Count > 0)
                        {
                            dataString = new string[]
                            {
                                "Остатки склады:",
                                "",
                                "",
                                "",
                                "",
                                ""
                            };
                            dgvStorages.Rows.Add(dataString);

                            foreach (var aa in ListAA.OrderBy(x => x.BBFDate))
                            {
                                dataString = new string[]
                                 {
                                    aa.Storage,
                                    aa.LotName,
                                    aa.Quantity.ToString("N1"),
                                    aa.ProdDate.ToString("dd.MM.yyyy") ,    //ss.ValidFrom.ToString("dd.MM.yyyy"),
                                    aa.BBFDate.ToString("dd.MM.yyyy"),    // ss.ValidTo.ToString("dd.MM.yyyy"),
                                    aa.Color            //ss.Status
                                 };
                                dgvStorages.Rows.Add(dataString);
                            }

                            if (ListAA.Count > 1)
                            {
                                dataString = new string[]
                                {
                                "",
                                "Итого остатки:",
                                ListAA.Sum(x=>x.Quantity).ToString("N1"),
                                "",
                                "",
                                ""
                                };
                                dgvStorages.Rows.Add(dataString);
                            }

                            dataString = new string[]
                             {
                                "",
                                "",
                                "",
                                "",
                                "",
                                ""
                            };
                            dgvStorages.Rows.Add(dataString);
                        }

                        //plant
                        if (cfm1.Count > 0)
                        {
                            var EL = cfm1[0].RawEnterpriseList;

                            if (cbMaterialDelay.Checked == false)
                            {
                                EL = EL.Where(x => x.Quantity != 0).ToList();
                            }
                            else
                            {
                                EL = EL.Where(x => (x.BBFDate >= DateTime.Today.Date) && (x.Quantity != 0)).ToList();
                            }

                            foreach (var el in EL)
                            {
                                if (el.BBFDate.Year <= 2000)
                                {
                                    el.Color = "Violet";
                                }
                                else
                                {
                                    DateTime pBDzate = el.BBFDate;

                                    if ((pBDzate - DateTime.Today.Date).Days > 30)
                                    {
                                        el.Color = "Green";
                                    }
                                    else if ((pBDzate - DateTime.Today.Date).Days > 20)
                                    {
                                        el.Color = "Magenta";
                                    }
                                    else if ((pBDzate - DateTime.Today.Date).Days > 0)
                                    {
                                        el.Color = "DarkOrange";
                                    }
                                    else
                                    {
                                        el.Color = "Red";
                                    }
                                }
                            }

                            if (EL.Count > 0)
                            {
                                dataString = new string[]
                                 {
                                "Остатки производство:",
                                "",
                                "",
                                "",
                                "",
                                ""
                                 };
                                dgvStorages.Rows.Add(dataString);

                                foreach (var el in EL.OrderBy(x => x.BBFDate))
                                {
                                    dataString = new string[]
                                     {
                                    el.Storage,
                                    el.LotName,
                                    el.Quantity.ToString("N1"),
                                    el.ProdDate.ToString("dd.MM.yyyy") ,    //ss.ValidFrom.ToString("dd.MM.yyyy"),
                                    el.BBFDate.ToString("dd.MM.yyyy"),    // ss.ValidTo.ToString("dd.MM.yyyy"),
                                    el.Color            //ss.Status
                                     };
                                    dgvStorages.Rows.Add(dataString);
                                }

                                if (EL.Count > 1)
                                {
                                    dataString = new string[]
                                    {
                                "",
                                "Итого остатки:",
                                EL.Sum(x=>x.Quantity).ToString("N1"),
                                 "",
                                 "",
                                 ""
                                    };
                                    dgvStorages.Rows.Add(dataString);
                                }

                                dataString = new string[]
                                 {
                                "",
                                "",
                                "",
                                "",
                                "",
                                ""
                                };
                                dgvStorages.Rows.Add(dataString);
                            }
                        }

                        //tasks  //data from cfm?

                        if (cbMaterialDelay.Checked == false)
                        {
                            sListA = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();
                            sListB = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim() != "")).OrderBy(x => x.PlanOperDate).ToList();
                        }
                        else
                        {
                            sListA = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.PlanOperDate.Date >= DateTime.Today.Date) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();  //cfm1[0].TaskNavList1.Where(x => x.BBFDate >= DateTime.Today.Date).OrderBy(x => x.BBFDate).ToList();  
                            sListB = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.PlanOperDate.Date >= DateTime.Today.Date) && (x.StatusMZP.Trim() != "")).OrderBy(x => x.PlanOperDate).ToList();   //cfm1[0].TaskNavList2.Where(x => x.BBFDate >= DateTime.Today.Date).OrderBy(x => x.BBFDate).ToList();   
                        }

                        var sListBB = sListB.Where(x => /*(x.PlanOperDate.Date >= DateTime.Today.Date) && */(x.StatusMZP.Trim().ToLower() == "заказано")).ToList();

                        //containers
                        var cList = fm.ContList.Where(x => (x.MaterialCode == ProductCode)).OrderBy(x => x.ExpDT).ToList();
                        //  cList = cList.Where(x => x.ExpDT >= DateTime.Today.Date).ToList();

                        foreach (var c in cList)
                        {
                            NavPurchaseTable npt = new NavPurchaseTable
                            {
                                PlanOperDate = c.ExpDT,
                                PlanDate = c.ExpDT,
                                PlanQuantity = c.Quantity,
                                FactQuantity = 0,
                                OrderNymber = c.LotNo,
                                MZP = "",
                                Initiator = "",
                                MaterialCode = c.MaterialCode,
                                StatusMZP = c.LotDescription
                            };
                            sListBB.Add(npt);
                        }

                        if (sListA.Count > 0)
                        {
                            dataString = new string[]
                             {
                                "Запланированные заявки навижн:",
                                "",
                                "",
                                "",
                                "",
                                ""
                            };
                            dgvStorages.Rows.Add(dataString);

                            foreach (var s in sListA)
                            {
                                PlanDate = s.PlanOperDate;
                                if ((PlanDate - DateTime.Today.Date).Days > 30)
                                {
                                    State = "Green";
                                }
                                else
                                {
                                    PlanDate = s.PlanOperDate;
                                    if ((PlanDate - DateTime.Today.Date).Days > 30)
                                    {
                                        State = "Green";
                                    }
                                    else if ((PlanDate - DateTime.Today.Date).Days > 20)
                                    {
                                        State = "Magenta";
                                    }
                                    else if ((PlanDate - DateTime.Today.Date).Days > 0)
                                    {
                                        State = "DarkOrange";
                                    }
                                    else
                                    {
                                        State = "Red";
                                    }
                                }

                                dataString = new string[]
                                 {
                                    s.OrderNymber,
                                    s.PlanOperDate.ToString("dd.MM.yyyy")+" "+s.Initiator+" "+s.StatusMZP,
                                    (s.PlanQuantity-s.FactQuantity).ToString("N1"),   //FactQuantity   
                                    s.PlanOperDate.ToString("dd.MM.yyyy"),
                                    s.PlanOperDate.ToString("dd.MM.yyyy"),
                                    State
                                 };
                                dgvStorages.Rows.Add(dataString);
                            }

                            if (sListA.Count > 1)
                            {
                                dataString = new string[]
                                {
                            "",
                            "ИТОГО:",
                            sListA.Sum(x=>x.PlanQuantity-x.FactQuantity).ToString("N1"),
                            "",
                            "",
                            ""
                                };
                                dgvStorages.Rows.Add(dataString);
                            }

                            dataString = new string[]
                            {
                                "",
                                "",
                                "",
                                "",
                                "",
                                ""
                            };
                            dgvStorages.Rows.Add(dataString);
                        }

                        if (sListBB.Count > 0)
                        {
                            dataString = new string[]
                             {
                                "Рабочие заявки навижн:",
                                "",
                                "",
                                "",
                                "",
                                ""
                            };
                            dgvStorages.Rows.Add(dataString);

                            foreach (var s in sListBB)
                            {
                                if (s.PlanOperDate.Year <= 2000)
                                {
                                    State = "Violet";
                                }
                                else
                                {
                                    PlanDate = s.PlanOperDate;
                                    if ((PlanDate - DateTime.Today.Date).Days > 30)
                                    {
                                        State = "Green";
                                    }
                                    else if ((PlanDate - DateTime.Today.Date).Days > 20)
                                    {
                                        State = "Magenta";
                                    }
                                    else if ((PlanDate - DateTime.Today.Date).Days > 0)
                                    {
                                        State = "DarkOrange";
                                    }
                                    else
                                    {
                                        State = "Red";
                                    }
                                }

                                dataString = new string[]
                                 {
                                    s.MZP.Trim()==""? s.OrderNymber:  s.OrderNymber+"/"+s.MZP,
                                    s.PlanOperDate.ToString("dd.MM.yyyy")+" "+s.Initiator+" "+s.StatusMZP,
                                    (s.PlanQuantity-s.FactQuantity).ToString("N1"),  //FactQuantity    
                                    s.PlanOperDate.ToString("dd.MM.yyyy"),
                                    s.PlanOperDate.ToString("dd.MM.yyyy"),
                                    State
                                 };
                                dgvStorages.Rows.Add(dataString);
                            }

                            if (sListBB.Count > 1)
                            {
                                dataString = new string[]
                                {
                            "",
                            "ИТОГО:",
                            sListBB.Sum(x=>x.PlanQuantity-x.FactQuantity).ToString("N1"),
                            "",
                            "",
                            ""
                                };
                                dgvStorages.Rows.Add(dataString);
                            }

                            dataString = new string[]
                            {
                                "",
                                "",
                                "",
                                "",
                                "",
                                ""
                            };
                            dgvStorages.Rows.Add(dataString);
                        }

                        foreach (DataGridViewRow r in dgvStorages.Rows)
                        {
                            if (r.Cells[3].Value.ToString() != "")
                            {
                                r.Cells[0].ToolTipText = "Произведен: " + r.Cells[3].Value.ToString() + "; Годен до: " + r.Cells[4].Value.ToString();
                                r.Cells[1].ToolTipText = "Произведен: " + r.Cells[3].Value.ToString() + "; Годен до: " + r.Cells[4].Value.ToString();
                                r.Cells[2].ToolTipText = "Произведен: " + r.Cells[3].Value.ToString() + "; Годен до: " + r.Cells[4].Value.ToString();
                            }

                        }
                    }//storages, plant and tasks
                    //показать меню

                    //обновить переменную
                    CurrCode = ProductCode;
                }

                //MenuStrip
                contextMenuStripData.Items.Clear();

             //   if (dataGridViewMain.Rows[e.RowIndex].Cells["Класс материала"].Value.ToString().ToLower() == "гп")
                {
                    ToolStripMenuItem m1;
                    Int32 Offset = cf.WorkMonthList.Count;
                    cf.UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                    if (ColInd > 15 + Offset)
                    {
                        if ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Остаток на")) || (dataGridViewMain.Columns[ColInd].HeaderText.Contains("Потребление на")) || ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Дефицит на"))))
                        {
                            LName = dataGridViewMain.Columns[e.ColumnIndex].HeaderText;
                            LName = LName.Substring(LName.Length - 11, 11).Trim();
                            CurrDate = fm.ConvertDataToDate(LName);
                        }
                        else
                        {
                            CurrDate = DateTime.Today.Date;
                        }


                        var ExData = cf.XLDataList.Where(x => (x.ProdCode == cf.UserSelectedCode) && (x.DateWork == CurrDate)).ToList(); //fm.ExcelList.Where(x => (x.ProductCode == UserSelectedCode) && (x.DateWork == dt)).ToList();

                        foreach (var ed in ExData)
                        {
                            //  foreach (var xd in ed.EDList)
                            {
                                LName = ed.Quantity.ToString("N1") + " => " + ed.Line;
                                if (ed.RePack == true)
                                {
                                    LName = LName + " П/УП";
                                }

                                contextMenuStripData.Items.Add(LName);
                            }
                        }

                        if (ExData.Count > 1)
                        {
                            LName = "__________";
                            contextMenuStripData.Items.Add(LName);
                            LName = "Итого: " + (ExData.Sum(x => x.Quantity)).ToString("N1");
                            contextMenuStripData.Items.Add(LName);
                        }
                    }
                    else if (e.ColumnIndex == (8 + Offset))
                    {
                      //  var CM = CFMList.Where(x => x.MaterialCode == cf.UserSelectedCode).ToList();

                        if (cfm1.Count > 0)
                        {
                            foreach (var cl in cfm1[0].ConsurmptionList)
                            {
                                m1 = new ToolStripMenuItem(cl.Storage + " => " + cl.Quantity.ToString("N1"));
                                contextMenuStripData.Items.Add(m1);
                            }
                        }
                    }
                    else if (e.ColumnIndex == (9 + Offset))
                    {
                        var ListAA = cfm1[0].RawStorageList;

                        if (cbMaterialDelay.Checked == true)
                        {
                            ListAA = ListAA.Where(x => x.BBFDate >= DateTime.Today.Date).ToList();
                        }

                        foreach (var sml in ListAA.OrderBy(x => x.ProdDate)) //sbrMesList)
                        {
                            m1 = new ToolStripMenuItem(sml.Storage + " [" + sml.LotName + "] => " + sml.Quantity.ToString("N1"));
                            m1.ToolTipText = "Выпущен: " + sml.ProdDate.ToString("dd.MM.yyyy") + "; Годен до: " + sml.BBFDate.ToString("dd.MM.yyyy"); // +"; К-во: "+sml.Count.ToString();
                            try
                            {
                                m1.ForeColor = Color.FromName(sml.Color);
                            }
                            catch (Exception xx)
                            {
                                m1.ForeColor = Color.Violet;
                            }

                            contextMenuStripData.Items.Add(m1);
                        }
                    }
                    else if (e.ColumnIndex == (10 + Offset))
                    {
                        if (cfm1.Count > 0)
                        {
                            var EL = cfm1[0].RawEnterpriseList;

                            if (cbMaterialDelay.Checked == false)
                            {
                                EL = EL.Where(x => x.Quantity != 0).ToList();
                            }
                            else
                            {
                                EL = EL.Where(x => (x.BBFDate >= DateTime.Today.Date) && (x.Quantity != 0)).ToList();
                            }

                            foreach (var el in EL.OrderBy(x => x.ProdDate))
                            {
                                ToolStripMenuItem m2 = new ToolStripMenuItem(el.Storage + " [" + el.LotName + "] => " + el.Quantity.ToString("N1"));
                                //      m2.ToolTipText = "Количество записей: " + sl1.Count.ToString();
                                m2.ToolTipText = "Выпущен: " + el.ProdDate.ToString("dd.MM.yyyy") + "; Годен до: " + el.BBFDate.ToString("dd.MM.yyyy");
                                try
                                {
                                    m2.ForeColor = Color.FromName(el.Color);
                                }
                                catch (Exception xx)
                                {
                                    m2.ForeColor = Color.Violet;
                                }

                                contextMenuStripData.Items.Add(m2);
                            }
                        }
                    }
                    else if (e.ColumnIndex == (15 + Offset))
                    {
                        var sListAA = sListA.Where(x => x.PlanOperDate < DateTime.Today.Date).ToList();
                        var sListAAA = sListAA.GroupBy(x => new { x.OrderNymber, x.MZP, x.PlanOperDate }).
                            Select(g => new PurchTable
                            {
                                PlanOperDate = g.Key.PlanOperDate,
                                OrderNymber = g.Key.OrderNymber,
                                PlanQuantity = g.Sum(x => x.PlanQuantity - x.FactQuantity),
                                Count = g.Count()
                            }).ToList();

                        if (sListAAA.Count > 0)
                        {
                            m1 = new ToolStripMenuItem("Запланированные заявки Навижн");
                            contextMenuStripData.Items.Add(m1);

                            foreach (var sb in sListAAA.OrderBy(x => x.PlanOperDate))
                            {
                                m1 = new ToolStripMenuItem(sb.OrderNymber + " => " + sb.PlanQuantity.ToString("N0"));
                                m1.ToolTipText = "Плановая дата поставки: " + sb.PlanOperDate.ToString("dd.MM.yyyy") + "; К-во: " + sb.Count;
                                m1.ForeColor = Color.DarkRed;

                                contextMenuStripData.Items.Add(m1);
                            }
                        }

                        var sListBBB = sListB.Where(x => x.PlanOperDate < DateTime.Today.Date).ToList();
                        var sListBBB1 = sListBBB.GroupBy(x => new { x.OrderNymber, x.MZP, x.PlanOperDate }).
                            Select(g => new PurchTable
                            {
                                PlanOperDate = g.Key.PlanOperDate,
                                LostDescr = g.Key.MZP,
                                OrderNymber = g.Key.OrderNymber,
                                PlanQuantity = g.Sum(x => x.PlanQuantity - x.FactQuantity),
                                Count = g.Count()
                            }).ToList();
                        if (sListBBB1.Count > 0)
                        {
                            m1 = new ToolStripMenuItem("Рабочие заявки Навижн");
                            contextMenuStripData.Items.Add(m1);

                            foreach (var sb in sListBBB1.OrderBy(x => x.PlanOperDate))
                            {
                                m1 = new ToolStripMenuItem(sb.OrderNymber + "/" + sb.LostDescr + " => " + sb.PlanQuantity.ToString("N0"));
                                m1.ToolTipText = "Плановая дата поставки: " + sb.PlanOperDate.ToString("dd.MM.yyyy") + "; К-во: " + sb.Count;
                                m1.ForeColor = Color.DarkRed;

                                contextMenuStripData.Items.Add(m1);
                            }
                        }

                        var cList = fm.ContList.Where(x => (x.MaterialCode == ProductCode)).OrderBy(x => x.ExpDT).ToList();
                     //   cList = cList.Where(x => x.ExpDT < DateTime.Today.Date).ToList();

                        var cListA = cList.GroupBy(x => new { x.LotNo, x.LotDescription, x.ExpDT }).
                            Select(g => new PurchTable
                            {
                                PlanOperDate = g.Key.ExpDT,
                                OrderNymber = g.Key.LotNo,
                                LostDescr = g.Key.LotDescription,
                                PlanQuantity = g.Sum(x => x.Quantity),
                                Count = g.Count()
                            }).ToList();

                        if (cListA.Count > 0)
                        {
                            m1 = new ToolStripMenuItem("Лоты/Контейнера");
                            contextMenuStripData.Items.Add(m1);

                            foreach (var c in cListA.OrderBy(x => x.PlanOperDate))
                            {
                                m1 = new ToolStripMenuItem(c.OrderNymber + "/" + c.LostDescr + " => " + c.PlanQuantity.ToString("N0"));
                                m1.ToolTipText = "Плановая дата поставки: " + c.PlanOperDate.ToString("dd.MM.yyyy") + "; К-во: " + c.Count;
                                m1.ForeColor = Color.DarkRed;

                                contextMenuStripData.Items.Add(m1);
                            }
                        }
                    }
                }
           
                contextMenuStripData.Show(AxX, AxY);

            }  //rowInd, ColInd

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Cursor = Cursors.Default;
        }

        private void dgvStorages_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            if (dgvStorages.RowCount > 0)
            {
                foreach (DataGridViewRow r in dgvStorages.Rows)
                {
                    Color cColor;
                    try
                    {
                        cColor = Color.FromName(r.Cells[5].Value.ToString());
                    }
                    catch (Exception xx)
                    {
                        cColor = Color.Black;
                    }

                    foreach (DataGridViewCell c in r.Cells)
                    {
                        c.Style.ForeColor = cColor;
                    }
                }
            }
        }

        private void button7_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.RowCount > 0)
            {
                Cursor = Cursors.WaitCursor;
                double Value;
                string sVal;
                double T1;
                double T2;
                double Q = 0;
                Int32 Cols = cf.WorkMonthList.Count;
                Int32 CurCol = 16 + cf.WorkMonthList.Count;
                DateTime WorkDate;
                List<DateTime> WorkDateList = new List<DateTime>();
                Int32 Interval;

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Остатки с МЗП";
                Excel1.Range range1;

                List<DateTime> DTList1 = new List<DateTime>();
                DTList1 = cf.DTList;
                if (cf.cbDontShowAfter.Checked == true)
                {
                    var XLD = fm.ExcelDataList[cf.comboBoxFiltering.SelectedIndex];
                    DateTime LimitDate = XLD.DateWork;
                    DTList1 = (from d in DTList1
                               where d <= LimitDate
                               select d).ToList();
                }

                for (int i = 0; i < dataGridViewMain.Columns.Count - 1; i++)
                {
                    wSheet.Cells[1, i + 1] = dataGridViewMain.Columns[i].HeaderText;
                }

                range1 = wSheet.Cells[1, 1] as Excel1.Range;
                range1.ColumnWidth = 15;
                range1 = wSheet.Cells[1, 2] as Excel1.Range;
                range1.ColumnWidth = 50;
                range1 = wSheet.Cells[1, 3] as Excel1.Range;
                range1.ColumnWidth = 30;

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dataGridViewMain.Columns.Count - 1]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CFMList.Count + 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i < CFMList.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = CFMList[i].MaterialCode;
                    wSheet.Cells[i + 2, 2] = CFMList[i].MaterialName;
                    wSheet.Cells[i + 2, 3] = CFMList[i].MaterialGroup;
                    wSheet.Cells[i + 2, 4] = CFMList[i].MaterialMultiply;
                    wSheet.Cells[i + 2, 5] = CFMList[i].WaitingDays;
                    wSheet.Cells[i + 2, 6] = CFMList[i].StorageQuantity;
                    wSheet.Cells[i + 2, 7] = CFMList[i].Responsible;
                    wSheet.Cells[i + 2, 8] = CFMList[i].Sender;

                    for (int j = 0; j < Cols; j++)
                    {
                        wSheet.Cells[i + 2, 9 + j] = Math.Round(CFMList[i].OutList[j].Quantity, 0);
                    }

                    wSheet.Cells[i + 2, 9 + Cols] = Math.Round(CFMList[i].CurrentConsumption, 0);
                    wSheet.Cells[i + 2, 10 + Cols] = Math.Round(CFMList[i].RawStorage, 0);
                    wSheet.Cells[i + 2, 11 + Cols] = Math.Round(CFMList[i].RawEnterprise, 0);
                    wSheet.Cells[i + 2, 12 + Cols] = CFMList[i].RawNav1.ToString("N0");
                    wSheet.Cells[i + 2, 13 + Cols] = CFMList[i].RawNav2.ToString("N0");
                    wSheet.Cells[i + 2, 14 + Cols] = CFMList[i].RawMesOut1.ToString("N0");
                    wSheet.Cells[i + 2, 15 + Cols] = CFMList[i].RawMesOut2.ToString("N0");
                    wSheet.Cells[i + 2, 16 + Cols] = Math.Round(CFMList[i].OldTaskNav, 0);

                    Value = CFMList[i].CurrentConsumption;

                    foreach (var sl in CFMList[i].RawStorageList)
                    {
                        Value = Value + sl.Quantity;
                    }

                  /*  foreach (var el in CFMList[i].RawEnterpriseList)
                    {
                        Value = Value + el.Quantity;
                    }*/

                    foreach (var r1 in CFMList[i].RawNav1List)
                    {
                        Value = Value + r1.Quantity;
                    }

                    /*  foreach (var r2 in CFMList[i].RawNav2List)
                      {
                          Value = Value + r2.Quantity;
                      }*/
                    var AppList = fm.tdb.tApplication.ToList();

                    var adata = AppList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList(); 
                    var DL = cf.XLDataList.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();
                    var CD = fm.ContList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();

                    for (int j = 0; j < DTList1.Count; j++)
                    {
                        T2 = 0;
                        var FDTList = DTList1.Where(x => (x.Month == DTList1[j].Month) && (x.Year == DTList1[j].Year)).ToList();
                        //income
                        var tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == DTList1[j].Date).ToList();
                        if ((DTList1[j].Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.AlterDate == DTList1[j]).ToList();
                        }

                        foreach (var td2 in tData2)
                        {
                            T2 = T2 + td2.Quantity;
                        }

                        var Cdata = CD.Where(x => x.ExpDT.Date == DTList1[j].Date).ToList();     //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == DTList1[j].Date)).ToList();
                        if ((DTList1[j].Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = CD.Where(x => ((DateTime)x.AlterDate).Date == DTList1[j].Date).ToList();   //fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == DTList1[j].Date)).ToList();
                        }

                        foreach (var cd in Cdata)
                        {
                            T2 = T2 + cd.Quantity;
                        }

                        if (T2>0)
                        {
                            if (Value<0)
                            {
                                Value = 0;
                            }
                        }
                        Value = Value + T2;
                        //ask in progress
                        var tData1 = CFMList[i].TaskNavList1.Where(x => x.BBFDate == DTList1[j].Date).ToList();
                        if ((DTList1[j].Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate == DTList1[j]).ToList();
                        }

                        T1 = 0;
                        foreach (var td1 in tData1)
                        {
                            T1 = T1 + td1.Quantity;
                        }
                       // Value = Value + T1;
                        //credit
                        var dll = DL.Where(x => x.DateWork.Date == DTList1[j].Date).ToList();   // cf.XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList1[j].Date)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value - dl.Quantity;
                        }

                        Q = 0;
                        var bdata = adata.Where(x => x.DateWork.Date == DTList1[j].Date).ToList();    //cf.AppList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList1[j].Date)).ToList();
                        if ((bdata.Count == 0) && (FDTList.Count == 1) && (DTList1[j].Day == 1))
                        {
                           // adata = cf.AppList;
                           // adata = adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                            bdata = adata.Where(x => (DateTime)x.AlterDate == DTList1[j].Date).ToList();
                        }

                        foreach (var ad in bdata)
                        {
                            Q = Q + ad.Quantity;
                        }

                       /* if (Value < 0)
                        {
                            wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(Value, 0);
                            range1 = wSheet.Cells[i + 2, CurCol + j + 1] as Excel1.Range;
                            range1.Font.Bold = true;
                            range1.Font.Color = Color.Red;    //Color.DarkRed;
                        }
                        else
                        {*/
                            if ((T1 > 0) || (T2 > 0) || (Q > 0))
                            {
                                sVal = Value.ToString("N0");

                                if (T1 > 0)
                                {
                                    sVal = sVal + " [" + T1.ToString("N0") + "]";
                                }
                                if (T2 > 0)  //|| (T1 > 0)
                                {
                                    sVal = sVal + " (" + (T2).ToString("N0") + ")";  // + T1
                                }
                                if (Q > 0)
                                {
                                    foreach (var bd in bdata)
                                    {
                                        sVal = sVal + " {" + bd.Quantity.ToString("N0") + "}";
                                    }
                                }
                                wSheet.Cells[i + 2, CurCol + j + 1] = sVal;
                            }
                            else
                            {
                                wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(Value, 0);
                            }

                        if (Value < 0)
                        {
                            range1 = wSheet.Cells[i + 2, CurCol + j + 1] as Excel1.Range;
                            range1.Font.Bold = true;
                            range1.Font.Color = Color.Red;  //Color.DarkRed
                        }
                        //  }
                    }
                }  //остаток

                wSheet = (Excel1.Worksheet)wBook.Sheets.Add(After: wBook.ActiveSheet);
                wSheet.Name = "Остатки без МЗП";

                for (int i = 0; i < dataGridViewMain.Columns.Count - 1; i++)
                {
                    wSheet.Cells[1, i + 1] = dataGridViewMain.Columns[i].HeaderText;
                }

                range1 = wSheet.Cells[1, 1] as Excel1.Range;
                range1.ColumnWidth = 20;
                range1 = wSheet.Cells[1, 2] as Excel1.Range;
                range1.ColumnWidth = 50;
                range1 = wSheet.Cells[1, 3] as Excel1.Range;
                range1.ColumnWidth = 30;

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dataGridViewMain.Columns.Count - 1]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CFMList.Count + 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i < CFMList.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = CFMList[i].MaterialCode;
                    wSheet.Cells[i + 2, 2] = CFMList[i].MaterialName;
                    wSheet.Cells[i + 2, 3] = CFMList[i].MaterialGroup;
                    wSheet.Cells[i + 2, 4] = CFMList[i].MaterialMultiply;
                    wSheet.Cells[i + 2, 5] = CFMList[i].WaitingDays;
                    wSheet.Cells[i + 2, 6] = CFMList[i].StorageQuantity;
                    wSheet.Cells[i + 2, 7] = CFMList[i].Responsible;
                    wSheet.Cells[i + 2, 8] = CFMList[i].Sender;

                    for (int j = 0; j < Cols; j++)
                    {
                        wSheet.Cells[i + 2, 9 + j] = Math.Round(CFMList[i].OutList[j].Quantity, 0);
                    }

                    wSheet.Cells[i + 2, 9 + Cols] = Math.Round(CFMList[i].CurrentConsumption, 0);
                    wSheet.Cells[i + 2, 10 + Cols] = Math.Round(CFMList[i].RawStorage, 0);
                    wSheet.Cells[i + 2, 11 + Cols] = Math.Round(CFMList[i].RawEnterprise, 0);
                    wSheet.Cells[i + 2, 12 + Cols] = CFMList[i].RawNav1.ToString("N0");
                    wSheet.Cells[i + 2, 13 + Cols] = CFMList[i].RawNav2.ToString("N0");
                    wSheet.Cells[i + 2, 14 + Cols] = CFMList[i].RawMesOut1.ToString("N0");
                    wSheet.Cells[i + 2, 15 + Cols] = CFMList[i].RawMesOut2.ToString("N0");
                    wSheet.Cells[i + 2, 16 + Cols] = Math.Round(CFMList[i].OldTaskNav, 0);

                    Value = CFMList[i].CurrentConsumption;

                    foreach (var sl in CFMList[i].RawStorageList)
                    {
                        Value = Value + sl.Quantity;
                    }

                /*    foreach (var el in CFMList[i].RawEnterpriseList)
                    {
                        Value = Value + el.Quantity;
                    }*/

                    foreach (var r1 in CFMList[i].RawNav1List)
                    {
                        Value = Value + r1.Quantity;
                    }

                    /*  foreach (var r2 in CFMList[i].RawNav2List)
                      {
                          Value = Value + r2.Quantity;
                      }*/

                    var DL = cf.XLDataList.Where(x => x.ProdCode == CFMList[i].MaterialCode).ToList();

                    for (int j = 0; j < DTList1.Count; j++)
                    {
                        //credit
                        var dll = DL.Where(x => x.DateWork.Date == DTList1[j].Date).ToList();    //cf.XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList1[j].Date)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value - dl.Quantity;
                        }

                        wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(Value, 0);

                        if (Value < 0)
                        {
                            range1 = wSheet.Cells[i + 2, CurCol + j + 1] as Excel1.Range;
                            range1.Font.Bold = true;
                            range1.Font.Color = Color.Red; //Color.DarkRed;
                        }
                    }
                }  //остаток2
                wSheet = (Excel1.Worksheet)wBook.Sheets.Add(After: wBook.ActiveSheet);
                wSheet.Name = "Дефицит";

                for (int i = 0; i < dataGridViewMain.Columns.Count - 1; i++)
                {
                    wSheet.Cells[1, i + 1] = dataGridViewMain.Columns[i].HeaderText;
                }

                range1 = wSheet.Cells[1, 1] as Excel1.Range;
                range1.ColumnWidth = 20;
                range1 = wSheet.Cells[1, 2] as Excel1.Range;
                range1.ColumnWidth = 50;
                range1 = wSheet.Cells[1, 3] as Excel1.Range;
                range1.ColumnWidth = 30;

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dataGridViewMain.Columns.Count - 1]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CFMList.Count + 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i < CFMList.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = CFMList[i].MaterialCode;
                    wSheet.Cells[i + 2, 2] = CFMList[i].MaterialName;
                    wSheet.Cells[i + 2, 3] = CFMList[i].MaterialGroup;
                    wSheet.Cells[i + 2, 4] = CFMList[i].MaterialMultiply;
                    wSheet.Cells[i + 2, 5] = CFMList[i].WaitingDays;
                    wSheet.Cells[i + 2, 6] = CFMList[i].StorageQuantity;
                    wSheet.Cells[i + 2, 7] = CFMList[i].Responsible;
                    wSheet.Cells[i + 2, 8] = CFMList[i].Sender;

                    for (int j = 0; j < Cols; j++)
                    {
                        wSheet.Cells[i + 2, 9 + j] = Math.Round(CFMList[i].OutList[j].Quantity, 0);
                    }

                    wSheet.Cells[i + 2, 9 + Cols] = Math.Round(CFMList[i].CurrentConsumption, 0);
                    wSheet.Cells[i + 2, 10 + Cols] = Math.Round(CFMList[i].RawStorage, 0);
                    wSheet.Cells[i + 2, 11 + Cols] = Math.Round(CFMList[i].RawEnterprise, 0);
                    wSheet.Cells[i + 2, 12 + Cols] = CFMList[i].RawNav1.ToString("N0");
                    wSheet.Cells[i + 2, 13 + Cols] = CFMList[i].RawNav2.ToString("N0");
                    wSheet.Cells[i + 2, 14 + Cols] = CFMList[i].RawMesOut1.ToString("N0");
                    wSheet.Cells[i + 2, 15 + Cols] = CFMList[i].RawMesOut2.ToString("N0");
                    wSheet.Cells[i + 2, 16 + Cols] = Math.Round(CFMList[i].OldTaskNav, 0);

                    Value = CFMList[i].CurrentConsumption;

                    foreach (var sl in CFMList[i].RawStorageList)
                    {
                        Value = Value + sl.Quantity;
                    }

                  /*  foreach (var el in CFMList[i].RawEnterpriseList)
                    {
                        Value = Value + el.Quantity;
                    }*/

                    foreach (var r1 in CFMList[i].RawNav1List)
                    {
                        Value = Value + r1.Quantity;
                    }

                    /*  foreach (var r2 in CFMList[i].RawNav2List)
                      {
                          Value = Value + r2.Quantity;
                      }*/

                    for (int j = 0; j < DTList1.Count; j++)
                    {
                        T2 = 0;
                        var FDTList = DTList1.Where(x => (x.Month == DTList1[j].Month) && (x.Year == DTList1[j].Year)).ToList();
                        //income
                        var tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == DTList1[j].Date).ToList();
                        if ((DTList1[j].Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.AlterDate == DTList1[j]).ToList();
                        }

                        foreach (var td2 in tData2)
                        {
                            T2 = T2 + td2.Quantity;
                        }

                        var Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == DTList1[j].Date)).ToList();
                        if ((DTList1[j].Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == DTList1[j].Date)).ToList();
                        }

                        foreach (var cd in Cdata)
                        {
                            T2 = T2 + cd.Quantity;
                        }

                        if (T2 > 0)
                        {
                            if (Value < 0)
                            {
                                Value = 0;
                            }
                        }
                        Value = Value + T2;

                        //ask in progress
                        var tData1 = CFMList[i].TaskNavList1.Where(x => x.BBFDate == DTList1[j].Date).ToList();
                        if ((DTList1[j].Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate == DTList1[j]).ToList();
                        }

                        T1 = 0;
                        foreach (var td1 in tData1)
                        {
                            T1 = T1 + td1.Quantity;
                        }
                      //  Value = Value + T1;

                        //credit
                        var dll = cf.XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList1[j].Date)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value - dl.Quantity;
                        }

                        if (Value < 0)
                        {
                            wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(-1 * Value, 0);
                        }
                        else
                        { }
                    }
                }  //дефицит

                wSheet = (Excel1.Worksheet)wBook.Sheets.Add(After: wBook.ActiveSheet);
                wSheet.Name = "Потребление";

                for (int i = 0; i < dataGridViewMain.Columns.Count - 1; i++)
                {
                    wSheet.Cells[1, i + 1] = dataGridViewMain.Columns[i].HeaderText;
                }

                range1 = wSheet.Cells[1, 1] as Excel1.Range;
                range1.ColumnWidth = 20;
                range1 = wSheet.Cells[1, 2] as Excel1.Range;
                range1.ColumnWidth = 50;
                range1 = wSheet.Cells[1, 3] as Excel1.Range;
                range1.ColumnWidth = 30;

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dataGridViewMain.Columns.Count - 1]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CFMList.Count + 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i < CFMList.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = CFMList[i].MaterialCode;
                    wSheet.Cells[i + 2, 2] = CFMList[i].MaterialName;
                    wSheet.Cells[i + 2, 3] = CFMList[i].MaterialGroup;
                    wSheet.Cells[i + 2, 4] = CFMList[i].MaterialMultiply;
                    wSheet.Cells[i + 2, 5] = CFMList[i].WaitingDays;
                    wSheet.Cells[i + 2, 6] = CFMList[i].StorageQuantity;
                    wSheet.Cells[i + 2, 7] = CFMList[i].Responsible;
                    wSheet.Cells[i + 2, 8] = CFMList[i].Sender;

                    for (int j = 0; j < Cols; j++)
                    {
                        wSheet.Cells[i + 2, 9 + j] = Math.Round(CFMList[i].OutList[j].Quantity, 0);
                    }

                    wSheet.Cells[i + 2, 9 + Cols] = Math.Round(CFMList[i].CurrentConsumption, 0);
                    wSheet.Cells[i + 2, 10 + Cols] = Math.Round(CFMList[i].RawStorage, 0);
                    wSheet.Cells[i + 2, 11 + Cols] = Math.Round(CFMList[i].RawEnterprise, 0);
                    wSheet.Cells[i + 2, 12 + Cols] = CFMList[i].RawNav1.ToString("N0");
                    wSheet.Cells[i + 2, 13 + Cols] = CFMList[i].RawNav2.ToString("N0");
                    wSheet.Cells[i + 2, 14 + Cols] = CFMList[i].RawMesOut1.ToString("N0");
                    wSheet.Cells[i + 2, 15 + Cols] = CFMList[i].RawMesOut2.ToString("N0");
                    wSheet.Cells[i + 2, 16 + Cols] = Math.Round(CFMList[i].OldTaskNav, 0);

                    for (int j = 0; j < DTList1.Count; j++)
                    {
                        Value = 0;
                        //credit
                        var dll = cf.XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList1[j].Date)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value + dl.Quantity;
                        }

                        if (Value > 0)
                        {
                            wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(Value, 0);
                        }
                    }
                }  //потребление

                wSheet = (Excel1.Worksheet)wBook.Sheets.Add(After: wBook.ActiveSheet);
                wSheet.Name = "ОПР";

                WorkDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                WorkDate = WorkDate.AddMonths(-1);
                WorkDateList.Add(WorkDate);
                WorkDate = WorkDate.AddMonths(-1);
                WorkDateList.Insert(0, WorkDate);
                WorkDate = WorkDate.AddMonths(-1);
                WorkDateList.Insert(0, WorkDate);

                wSheet.Cells[1, 1] = "Код Материала";
                wSheet.Cells[1, 2] = "Наименование";
                wSheet.Cells[1, 3] = "Срок поставки";
                wSheet.Cells[1, 4] = "Кратность поставки";
                wSheet.Cells[1, 5] = "Месячное потребление " + WorkDateList[0].ToString("MM.yyyy");
                wSheet.Cells[1, 6] = "Месячное потребление " + WorkDateList[1].ToString("MM.yyyy");
                wSheet.Cells[1, 7] = "Месячное потребление " + WorkDateList[2].ToString("MM.yyyy");
                wSheet.Cells[1, 8] = "Среднее";
                wSheet.Cells[1, 9] = "Остатки Склад";
                wSheet.Cells[1, 10] = "Остатки Производство";
                wSheet.Cells[1, 11] = "Остатки Внешние";
                wSheet.Cells[1, 12] = "Итого, Остаток";
                wSheet.Cells[1, 13] = "Остатков хватит, дней";

                range1 = wSheet.Cells[1, 1] as Excel1.Range;
                range1.ColumnWidth = 20;
                range1 = wSheet.Cells[1, 2] as Excel1.Range;
                range1.ColumnWidth = 50;

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 13]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[fm.OPRList.Count + 1, 1]];
                range1.NumberFormat = "@";

                var Clist = CFMList.Where(x => fm.OPRList.Contains(x.MaterialCode)).ToList();
                WorkDate = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                Interval = DateTime.Today.Date.Subtract(WorkDate).Days;

                for (int i = 0; i < Clist.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = Clist[i].MaterialCode;
                    wSheet.Cells[i + 2, 2] = Clist[i].MaterialName;
                    wSheet.Cells[i + 2, 3] = Clist[i].WaitingDays;
                    wSheet.Cells[i + 2, 4] = Clist[i].StorageQuantity;

                    var OList = Clist[i].OutList.Where(x => x.ProdDate == WorkDateList[0]).ToList();
                    Value = 0;

                    if (OList.Count > 0)
                    {
                        Value = OList.Sum(x => x.Quantity);
                        wSheet.Cells[i + 2, 5] = Value;
                    }

                    OList = Clist[i].OutList.Where(x => x.ProdDate == WorkDateList[1]).ToList();
                    if (OList.Count > 0)
                    {
                        Value = Value + OList.Sum(x => x.Quantity);
                        wSheet.Cells[i + 2, 6] = OList.Sum(x => x.Quantity);
                    }

                    OList = Clist[i].OutList.Where(x => x.ProdDate == WorkDateList[2]).ToList();
                    if (OList.Count > 0)
                    {
                        Value = Value + OList.Sum(x => x.Quantity);
                        wSheet.Cells[i + 2, 7] = OList.Sum(x => x.Quantity);
                    }

                    T2 = 0;
                    Value = Value / 3.0;
                    wSheet.Cells[i + 2, 8] = Math.Round(Value);
                    wSheet.Cells[i + 2, 9] = Math.Round(Clist[i].RawStorage, 0);
                    T2 = Clist[i].RawStorage;
                    wSheet.Cells[i + 2, 10] = Math.Round(Clist[i].RawEnterprise, 0);
                    T2 = T2 + Clist[i].RawEnterprise;
                    wSheet.Cells[i + 2, 11] = Math.Round(Clist[i].RawNav1);
                    T2 = T2 + Clist[i].RawNav1;
                    wSheet.Cells[i + 2, 12] = Math.Round(T2, 0);

                    Value = (T2 * 30 / Value) - Interval;
                    wSheet.Cells[i + 2, 13] = Math.Round(Value, 0);
                }//OPR

                //  List<string> OPRList = new List<string>();

                var OPRListNames = fm.OPRList.Except(Clist.Select(x => x.MaterialCode).ToList()).ToList();
                Int32 Offer = Clist.Count + 5;
                var OprList = fm.tdb.PlanOPZList.Where(x => (bool)x.IsBase != true).ToList();    // fm.helg.PlanOPZList.ToList();
                OprList = OprList.Where(x => OPRListNames.Contains(x.MatCode)).ToList();

                range1 = wSheet.Range[wSheet.Cells[Offer, 1], wSheet.Cells[OprList.Count + Offer + 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i < OprList.Count; i++)
                {
                    wSheet.Cells[i + Offer, 1] = OprList[i].MatCode;
                    wSheet.Cells[i + Offer, 2] = OprList[i].MatName;

                    T2 = 0;
                    var dData = fm.StockBaklanceList.Where(x => x.WorkDate.Date == DateTime.Today.Date && x.MaterialCode == OprList[i].MatCode).ToList();
                    T2 = dData.Sum(x => (double)x.Quantity);  //storage
                    wSheet.Cells[i + Offer, 9] = T2;

                    var dData1 = fm.StockBalancesPlantList.Where(x => x.WorkDate.Date == DateTime.Today.Date && x.MaterialCode == OprList[i].MatCode).ToList();
                    T2 = T2 + dData1.Sum(x => (double)x.Quantity);
                    wSheet.Cells[i + Offer, 10] = dData1.Sum(x => x.Quantity);  //Plant

                    var dData2 = fm.StockBalancesMesList.Where(x => x.WorkDate.Date == DateTime.Today.Date && x.MaterialCode == OprList[i].MatCode).ToList();
                    T2 = T2 + dData2.Sum(x => (double)x.Quantity);
                    wSheet.Cells[i + Offer, 11] = dData2.Sum(x => x.Quantity);  //mes

                    wSheet.Cells[i + Offer, 12] = T2;
                }

                wSheet = null;
                wBook = null;
                range1 = null;
                //   xlApp.Quit();
                xlApp = null;

                MessageBox.Show("Отчет сформирован", "Сообщение системы");
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Cursor = Cursors.Default;
        }

        private void cbMaterialDelay_CheckedChanged(object sender, EventArgs e)
        {
            panel4.Enabled = false;
            ShowData();
            DrawDGV();
            panel4.Enabled = true;
        }

        private void cbUseMaterialPlanning_CheckedChanged(object sender, EventArgs e)
        {
            panel4.Enabled = false;
            ShowData();
            DrawDGV();
            panel4.Enabled = true;
        }

        private void cbUsingMaterial_CheckedChanged(object sender, EventArgs e)
        {
            panel4.Enabled = false;
            ShowData();
            DrawDGV();
            panel4.Enabled = true;
        }

        private void cbAddDay_CheckedChanged(object sender, EventArgs e)
        {
            panel4.Enabled = false;
            ShowData();
            DrawDGV();
            panel4.Enabled = true;
        }

        private void cbDontShowAfter_CheckedChanged(object sender, EventArgs e)
        {
            panel4.Enabled = false;
            ShowData();
            DrawDGV();
            panel4.Enabled = true;
        }

        private void cbViews_SelectedIndexChanged(object sender, EventArgs e)
        {
            panel4.Enabled = false;
            ShowData();
            DrawDGV();
            panel4.Enabled = true;
        }

        private void dataGridViewMain_Sorted(object sender, EventArgs e)
        {
            DrawDGV();
        }

        private void dataGridViewMain_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //DrawDGV();
        }

        private void dataGridViewMain_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            Int32 Offset = cf.WorkMonthList.Count;
            //cells values
            Font f1 = new Font(dataGridViewMain.DefaultCellStyle.Font, FontStyle.Bold);
            //1-2 cells
            for (int j = 0; j < CFMList.Count; j++)
            {
                string PC = dataGridViewMain.Rows[j].Cells[0].Value.ToString();   // CFMList[j].MaterialCode;

                if (cf.URSCodeList.Contains(PC))                 //(dataGridViewMain.Rows[j].Cells["Код материала"].Value.ToString().ToLower() == URSList[i].MaterialCode.ToLower())
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = cf.SColor;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = cf.SColor;
                }
                else if (cf.URMCodeList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = cf.MColor;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = cf.MColor;
                }
                else if (fm.OPRList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightCyan;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightCyan;
                }
                else if (CFMListGold.Contains(PC))  //cf.
                {
                    if (CFMList[j].MaterialGroup != "ГП") //(CFMBool[j] == false) 
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightYellow;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightYellow;  //Gold
                    }
                }
                else if (CFMListGray.Contains(PC))  //cf.
                {
                    if (CFMList[j].MaterialGroup != "ГП")
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGray;   //.DarkGray;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGray;
                    }
                }

                if (CFMListRed.Contains(PC)) //-  cf.
                {
                    var CM = CFMListMinus.Where(x => x.ProdCode == PC).ToList();  //cf.
                    if (CM.Count > 0)
                    {
                        foreach (var c in CM[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c/* + Offset*/].Style.ForeColor = Color.Red;  // Color.DarkRed;
                            }
                            catch (Exception xx)
                            { }
                            /* try
                             {
                                 dataGridViewMain.Rows[j].Cells[c].Style.Font = f1;
                             }
                             catch (Exception xx)
                             { }*/
                        }
                    }
                }

                if (CFMListGreen.Contains(PC))  //+  cf.
                {
                    var CP = CFMListPlus.Where(x => x.ProdCode == PC).ToList(); //cf.
                    if (CP.Count > 0)
                    {
                        foreach (var c in CP[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightGreen;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                if (CFMListPink.Contains(PC))   //cf.
                {
                    var CL = CFMListLine.Where(x => x.ProdCode == PC).ToList();  //cf.
                    if (CL.Count > 0)
                    {
                        foreach (var c in CL[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightPink;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                if (CFMListGoldens.Contains(PC))   //cf.
                {
                    var CLP = CFMListLinePlus.Where(x => x.ProdCode == PC).ToList();  // cf.
                    if (CLP.Count > 0)
                    {
                        foreach (var c in CLP[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightYellow;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                if (cf.TaskList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGreen;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGreen;
                }
                else if (cf.MesList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.MistyRose;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.MistyRose;
                }


                //9/10 color
                try
                {
                    var cfm = CFMList.Where(x => x.MaterialCode == PC).First();
                    if (cf.cbMaterialDelay.Checked == false)
                    {
                        dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1);
                        dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2);
                    }
                    else
                    {
                        dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1UD);
                        dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2UD);
                    }
                }
                catch (Exception xxx)
                { }

                /* for (int i = 8; i <= 13; i++)
                 {
                     dataGridViewMain.Columns[i + Offset].HeaderCell.Style.Font = new Font(dataGridViewMain.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f, FontStyle.Bold);
                 }*/
            }
            Cursor = Cursors.Default;
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] dataString;
            bool fl;

            Cursor = Cursors.WaitCursor;

            List<DateTime> DList = new List<DateTime>();
            DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();
            DList = DList.Distinct().OrderBy(x => x.Date).ToList();
            string ProductCode = dataGridViewMain.Rows[RowInd].Cells[0].Value.ToString().Trim();
            double Quantity = 0;
            DateTime DateStart;
            DateTime DateEnd;

            DateTime StartDate1;
            DateTime EndDate1;
            DateTime CurrDate = DateTime.Today.Date;

            if (ColInd > 0)
            {
                if ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Остаток на")) || (dataGridViewMain.Columns[ColInd].HeaderText.Contains("Потребление на")) || ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Дефицит на"))))
                {
                    string LName = dataGridViewMain.Columns[ColInd].HeaderText;
                    LName = LName.Substring(LName.Length - 11, 11).Trim();
                    fl = true;

                    CurrDate = fm.ConvertDataToDate(LName);
                    DateStart = CurrDate.AddHours(8);
                    DateEnd = DateStart.AddDays(1);
                    Quantity = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[ColInd].Value.ToString().Trim());
                }
                else
                {
                    CurrDate = DateTime.Today.Date;
                }
            }
            else
            {
                CurrDate = DateTime.Today.Date;
            }

            var CurProduct = fm.lRec.Where(x => x.MaterialCode == ProductCode).ToList();
            List<DateTime> DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();

            switch (tabControl2.SelectedIndex)
            {
                case 1:
                    if (dgvPlanned.RowCount <= 1)
                    {
                        dgvPlanned.Rows.Clear();
                        dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };
                        dgvPlanned.Rows.Add(dataString);
                        var E1 = fm.ExcelDataList.Where(x => DDList.Contains(x.DateWork) && x.Quantity != 0).ToList();

                        for (int i=0; i<DDList.Count; i++)  // foreach (DateTime dt1 in DDList)   //DList
                        {
                            //planned
                            //  if (dt1 >= DateTime.Today.Date)
                            {
                                fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;  //Cells["Класс материала"]
                                var ELDListDate = E1.Where(x => (x.DateWork == DDList[i])).ToList();   //fm.ExcelDataList.Where(x => (x.DateWork == dt1) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //dt1

                                 for (int j=0; j< ELDListDate.Count; j++)   // foreach (var ELD in ELDListDate)
                                {
                                    var CP = CurProduct.Where(x => (x.ProductCode == ELDListDate[j].ProdCode) && (x.LineNumber == ELDListDate[j].Line) && (x.WorkDate <= DDList[i])).ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
                                    fn_select_RecipesView_Result SRR;
                                    try
                                    {
                                        //  CP = CP.Where(x => x.WorkDate <= dt1).OrderByDescending(x => x.WorkDate).ToList();
                                        SRR = CP.OrderByDescending(x => x.WorkDate).First();      // CP = CP.Where(x => x.WorkDate <= dt1).OrderByDescending(x => x.WorkDate).ToList();  //x.Starting_Date
                                    }
                                    catch (Exception xx)
                                    {
                                        SRR = null;
                                    }

                                    if (SRR != null)    // (CP.Count > 0)
                                    {
                                        //раскрутка рецептуры в обратном направлении! Продукт-Линия-совпадает ли дата?
                                        var PRL = fm.ProdRecipeList.Where(x => x.ProductCode == ELDListDate[j].ProdCodeStr).ToList();
                                        if (PRL.Count > 0)
                                        {
                                            var LL = PRL[0].RecLineList.Where(x => x.LineNumber == ELDListDate[j].Line).ToList();
                                            if (LL.Count > 0)
                                            {
                                                var DL = LL[0].RecList.Where(x => x.WorkDateStart <= DDList[i]).OrderByDescending(x => x.WorkDateStart).ToList();  //CurrDate
                                                if (DL.Count > 0)
                                                {
                                                    if (DL[0].WorkDateStart == SRR.WorkDate)
                                                    {
                                                        if (ELDListDate[j].RePack == true)
                                                        {
                                                            if (fl == true)
                                                            {
                                                                if ((double)SRR.Quantity > 0)
                                                                {
                                                                    dataString = new string[] {
                                                                    DDList[i].ToString("dd.MM.yyyy"),  //  CurrDate
                                                                    ELDListDate[j].Line,
                                                                    ELDListDate[j].Quantity.ToString("N0"),
                                                                    ((double) SRR.Quantity*1000).ToString("N0"),   //CP[0].Quantity
                                                                    (((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                                                                    ELDListDate[j].ProdCodeStr,
                                                                    ELDListDate[j].AlterProdName
                                                                    };

                                                                    dgvPlanned.Rows.Add(dataString);
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            if ((double)SRR.Quantity > 0)
                                                            {
                                                                dataString = new string[] {
                                                                DDList[i].ToString("dd.MM.yyyy"),
                                                                ELDListDate[j].Line,
                                                               ELDListDate[j].Quantity.ToString("N0"),
                                                                ((double) SRR.Quantity*1000).ToString("N0"),
                                                                (((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                                                               ELDListDate[j].ProdCodeStr,
                                                               ELDListDate[j].AlterProdName
                                                                };

                                                                dgvPlanned.Rows.Add(dataString);
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }

                                foreach (DataGridViewRow r in dgvPlanned.Rows)
                                {
                                    if (r.Cells[0].Value.ToString() != "")
                                    {
                                        r.Cells[0].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                        r.Cells[1].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                        r.Cells[2].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                        r.Cells[3].ToolTipText = "Количество: " + r.Cells[3].Value.ToString() + "; Норма на тонну: " + r.Cells[4].Value.ToString() + " Итого расход: " + r.Cells[5].Value.ToString();
                                    }
                                }
                                ELDListDate = null;
                            }                       // dt1
                        }  //foreach
                    }
                    break;
                case 0:
                    if (dgvFactis.RowCount <= 1)
                    {
                        dgvFactis.Rows.Clear();
                        dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };
                        dgvFactis.Rows.Add(dataString);
                        {
                            try
                            {
                                StartDate1 = DList.Min();
                                EndDate1 = DList.Max();
                            }
                            catch (Exception xx)
                            {
                                StartDate1 = DateTime.Today.Date;
                                EndDate1 = StartDate1.AddDays(1);
                            }

                            fl = dataGridViewMain.CurrentRow.Cells[2].Value.ToString().ToLower() == "тара" ? true : false;
                            //  var MuaSection = fm.MUAList.Where(x => (x.DateTime >= DateStart) && (x.DateTime < DateEnd)).ToList();    

                            var FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, StartDate1).OrderBy(x => x.LotForErp).ToList();      //   DateTime.Today.Date.AddYears(-2)               //fm.edb.fn_select_MaterialUsingByProduct(StartDate1, EndDate1, ProductCode).OrderBy(x => x.DateStart).ToList();

                            List<string> LotForErp = FactData.Select(x => x.LotForErp).Distinct().ToList();
                            List<fn_select_ProductionByCode_Result> FactData1 = new List<fn_select_ProductionByCode_Result>();

                            foreach (string lfe in LotForErp)
                            {
                                var fd = FactData.Where(x => x.LotForErp == lfe).ToList();

                                if (fd.Count > 0)
                                {
                                    var fd1 = fd[0];

                                    for (int i = 1; i < fd.Count; i++)
                                    {
                                        fd1.Quantity = fd1.Quantity + fd[i].Quantity;
                                        fd1.RawQuant = fd1.RawQuant + fd[i].RawQuant;
                                    }
                                    fd1.EndDate = fd[fd.Count - 1].EndDate;
                                }
                            }

                            foreach (var f in FactData)
                            {
                                // if (f.DateStart >= DateStart)
                                {
                                    dataString = new string[]
                                    {
                                    ((double)f.RawQuant).ToString("N3"),
                                    ((DateTime)f.StartDate).ToString("dd.MM.yyyy HH;mm"),
                                    ((DateTime)f.EndDate).ToString("dd.MM.yyyy HH:mm"),
                                    f.LotForErp,
                                    ((double)f.Quantity).ToString("N2"),
                                    f.MaterialCode,
                                    f.SystemName+" "+ f.MaterialName
                                    };

                                    dgvFactis.Rows.Add(dataString);
                                }
                            }

                            foreach (DataGridViewRow r in dgvFactis.Rows)
                            {
                                if (r.Cells[0].Value.ToString() != "")
                                {
                                    r.Cells[0].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                    r.Cells[1].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                    r.Cells[2].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                    r.Cells[3].ToolTipText = "Произведено: " + r.Cells[4].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                                }
                            }

                            FactData = null;
                        }  //fact
                    }
                    break;
                case 2:
                    if (dgvRecipes.Rows.Count <= 1)
                    {
                        dgvRecipes.Rows.Clear();
                        //    Quantity = fm.ConvertStringToDouble(dataGridViewMain.CurrentRow.cel.Value.ToString().Trim());
                        var PRListA = fm.ProdRecipeList.Where(x => x.ProductCode == ProductCode).ToList();     //PRList.Where(x => x.ProductCode == ProductCode).ToList();
                        if (PRListA.Count > 0)
                        {
                            dataString = new string[] { "", "Рецепты с разбиением по линиям, действующие с " + CurrDate.ToString("dd.MM.yyyy") + ":" };//+ System.Environment.NewLine + System.Environment.NewLine;
                            dgvRecipes.Rows.Add(dataString);
                        }
                        foreach (var prA in PRListA)
                        {
                            var PRListLines = prA.RecLineList;

                            foreach (var PRL in PRListLines)
                            {
                                var Recipesl = PRL.RecList.Where(x => (x.WorkDateStart <= CurrDate) && (x.WorkDateEnd >= CurrDate)).ToList();
                                if (Recipesl.Count > 0)
                                {
                                    dataString = new string[] { "", PRL.LineNumber };
                                    dgvRecipes.Rows.Add(dataString);
                                }

                                foreach (var RL in Recipesl)
                                {
                                    var RecList = RL.rList;

                                    foreach (var rrr in RecList)
                                    {
                                        dataString = new string[] { rrr.MaterialCode,
                                                                rrr.MaterialName,
                                                                rrr.MaterialQuantity.ToString("N6"),
                                                                (Quantity * rrr.MaterialQuantity).ToString("N3") };
                                        dgvRecipes.Rows.Add(dataString);
                                    }
                                }
                            }
                        }  //recipes
                    }
                    break;
                case 3:
                    if (dgvUsing.RowCount <= 1)
                    {
                      //  var LR = fm.lRec1.Where(x => x.MaterialCode == ProductCode).ToList();
                        List<fn_select_RecipesView_Result> LResult = new List<fn_select_RecipesView_Result>();
                        List<string> Slist = CurProduct.Select(x => x.LineNumber).Distinct().ToList();

                        foreach (var s in Slist)
                        {
                            try
                            {
                                var LR1 = CurProduct.Where(x => x.LineNumber == s && x.WorkDate <= DateTime.Today.Date).ToList();

                                StartDate1 = ((DateTime)LR1.Max(x => x.WorkDate));
                            }
                            catch (Exception xx)
                            {
                                StartDate1 = DateTime.Today.Date;
                            }

                            var LR2 = CurProduct.Where(x => x.LineNumber == s && x.WorkDate >= StartDate1).ToList();
                            LResult.AddRange(LR2);
                        }

                        LResult = LResult.OrderBy(x => x.WorkDate).ThenBy(x => x.LineNumber).ToList();
                        dgvUsing.Rows.Clear();

                        foreach (var lr in LResult)
                        {
                            if (lr.Quantity > 0)
                            {
                                double Q = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[16 + cf.WorkMonthList.Count].Value.ToString().Trim());

                                dataString = new string[]
                                {
                                    ((DateTime)lr.WorkDate).ToString("yyyy-MM-dd"),
                                    lr.LineNumber,
                                    ((double)lr.Quantity).ToString("N6"),
                                    ((double)(Q/lr.Quantity)).ToString("N2"),
                                    lr.ProductCode,
                                    lr.ProductName
                                };
                            }
                            else
                            {
                                dataString = new string[]
                                {
                                   ((DateTime)lr.WorkDate).ToString("yyyy-MM-dd"),
                                    lr.LineNumber,
                                    ((double)lr.Quantity).ToString("N6"),
                                    "---",
                                    lr.ProductCode,
                                    lr.ProductName
                                };
                            }
                            dgvUsing.Rows.Add(dataString);
                        }
                    }
                    break;
                case 4:
                    dgvComment.DataSource = null;
                    dgvComment.Rows.Clear();

                    var ComData = fm.tdb.fn_select_CommentDataID(ProductCode).ToList();
                    //   dgvComment.DataSource = ComData;

                    foreach (var CD in ComData)
                    {
                        dataString = new string[]
                        {
                            CD.IDC.ToString(),
                            CD.MaterialCodeC,
                            CD.MaterialNameC,
                            ((DateTime)CD.DateRecordC).ToString("yyyy-MM-dd"),
                            CD.IsActualC.ToString(),
                            CD.UserName,
                            CD.CommentC
                        };
                        dgvComment.Rows.Add(dataString);
                    }

                    dgvComment.DefaultCellStyle.WrapMode = DataGridViewTriState.True;
                    dgvComment.AutoSizeRowsMode = DataGridViewAutoSizeRowsMode.AllCells;
                    break;

                default:
                    break;
            }//switch
            Cursor = Cursors.Default;
        }
    }
}
