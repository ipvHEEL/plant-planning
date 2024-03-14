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
using System.Threading;
//using PlantPlanning.Models;

namespace PlantPlanning
{
    public partial class ConsurmptionForm : Form
    {
        FormMain fm;
        FormLineList fll;
        FormRequest fr;
        FormFilterResult ffr;
        FormTaskClear ftc;
        FormComment fc;

        public DateTime StartDate;
        public DateTime EndDate;

        public BindingSource bs;
        public DataTable dt1;

        public List<tApplication> AppList;
        public List<tUserRecordSelected> URSList;
        public List<tMercuryRecordSelected> URMList;
        public List<string> URSCodeList;
        public List<string> URMCodeList;
        public List<tColorSelection> CSList;
        public List<string> ProductNamesList;
        public List<string> MaterialNamesList;
        public List<lData> XLDataList;
        public List<DailyResult> MaterialData;
        public List<DailyResult> FilteredMaterialDataIn;
        public List<DailyResult> FilteredMaterialDataOut;

        public List<StockBalances> SBList;
        public List<StockBalancesResult> SBRList;  //ResultListData!!!!!
        public List<StockBalancesMes> SBMList;
        public List<StockBalancesPlant> SBPList;
        public List<UserSelectedCell> USCList;
        public List<Material> CFMList;
        public List<DateTime> DTList;
        // string[] CFMBool;
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
        public List<DateTime> WorkMonthList;
        public List<string> TaskList;
        public List<string> MesList;
        public List<ProductRecipes> TBP;
        public Int32 CurrIndex;
        public Int32 SelRowIndes;
        public Int32 ColIndex;
        public Int32 RowIndex;
        public List<string[]> RowList;
        public List<fn_select_RecipesView_Result> lRec1;

        public List<ProductRecipes> PRList;

        public Color SColor = Color.LightPink;   //.LightPink;  Color color = (Color)ColorConverter.ConvertFromString("#FFDFD991");
        public Color MColor = Color.RosyBrown; //Mercury!

        public UserSelectedCell USC;

        public Int32 AxX;
        public Int32 AxY;

        public string UserSelectedCode = "";
        public string UserSelectedProdName = "";
        public DateTime UserSelectedDateTime = DateTime.Today.Date;
        bool DayNight;
        public bool DataLoaded = false;

        private Cache memoryCache;
        private int LowPageIndex;
        private int HighPageIndex;
        private int OldLowPageIndex;
        private int OldHighPageIndex;

        public int ColumnSortedIndex = 1;
        public string ColumnSortedDirection = "ASC";
        private string CurrCode = "";

        Int32 RowInd = 0;
        Int32 ColInd = 0;
  

        public ConsurmptionForm(FormMain FM, FormLineList FLL)
        {
            InitializeComponent();
            fm = FM;
            fll = FLL;
            label6.Visible = false;
            label5.Visible = false;
            button6.Visible = false;
            button5.Visible = false;
            button9.Visible = false;
            string sVal = "";
            string sValUD = "";
            char RowSplitter = '|';

            CurrIndex = 0;
            SelRowIndes = 0;
            USCList = new List<UserSelectedCell>();
            DayNight = FLL.checkBox1.Checked;

            LoadData();  //workMonthList!
            GetStorage();

            DateTime XlDataStart = fm.ExcelDataList.Min(x => x.DateWork);
            XlDataStart = new DateTime(XlDataStart.Year, XlDataStart.Month, 1);
            DateTime XlDataEnd = fm.ExcelDataList.Max(x => x.DateWork);
            XlDataEnd = new DateTime(XlDataEnd.Year, XlDataEnd.Month, 1);
            XlDataEnd = XlDataEnd.AddMonths(1).AddDays(-1);

            fll.Close();
            fm.Enabled = false;
            AppList = fm.tdb.tApplication.ToList();

            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;
            cbViews.Text = cbViews.Items[0].ToString();
            dgvStorages.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvFactis.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            comboBox1.Items.Clear();
            comboBox1.Items.Add("ГП/ПФ");
            comboBox1.Items.Add("Мясосырье");
            comboBox1.Items.Add("Отходы");
            comboBox1.Items.Add("Специи/ингредиенты");
            comboBox1.Items.Add("Тара");

            //карта
            var ConsList = fm.ConsDataList;  //   fm.edb.fn_select_FactConsurmption().ToList();
            CFMList = fm.mList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).OrderBy(x => x.MaterialCode).ToList();
            //   CFMList = CFMList.Where(x => x.MaterialCode == "8СИ00127").ToList();
          
            //заявки
            TaskList = new List<string>();
            var SL = fm.StorageList.Where(x => (x.DateWork.Date == DateTime.Today.Date) && (x.DateWork.Hour >= 6) && (x.DateWork.Hour <= 10)).ToList();
            TaskList = SL.Select(x => x.MaterialCode).Distinct().ToList();



            for (int i = TaskList.Count - 1; i >= 0; i--)
            {
                var cfm = CFMList.Where(x => x.MaterialCode == TaskList[i]).ToList();

                if (cfm.Count > 0)
                {
                    TaskList.RemoveAt(i);
                }
            }

            foreach (string s in TaskList)
            {
                var sList = SL.Where(x => x.MaterialCode == s).ToList();      // fm.StorageList.Where(x => (x.DateWork.Date == DateTime.Today.Date) && (x.DateWork.Hour >= 6) && (x.DateWork.Hour <= 10) &&(x.MaterialCode == s)).ToList();
                if (sList.Count > 0)
                {
                  //  var sListOld = sList.Where(x => x.PlanOperDate < DateTime.Today.Date).ToList();
                  //  var sListNew = sList.Where(x => (x.PlanOperDate >= DateTime.Today.Date) && ((x.StatusMZP.ToLower().Trim() == "заказано") || (x.StatusMZP.ToLower().Trim() == ""))).ToList();

                    var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                    Material M = new Material();

                    M.MaterialCode = s;
                    M.MaterialName = sList[0].MaterialName;
                    if (MList.Count > 0)
                    {
                        M.MaterialGroup = MList[0].MaterialGroup;
                        M.MaterialCategory = MList[0].MaterialCategory;
                        M.MaterialMultiply = MList[0].MaterialMultiply;
                        M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                        M.WaitingDays = MList[0].WaitingDays;
                        M.WaitingDaysString = MList[0].WaitingDaysString;
                        M.StorageQuantity = MList[0].StorageQuantity;
                        M.StorageQuantityString = MList[0].StorageQuantityString;
                        M.Responsible = MList[0].Responsible;
                        M.Sender = MList[0].Sender;
                        M.DontShow = true;
                    }
                    else
                    {
                        M.MaterialGroup = "Н/Д";
                        M.MaterialCategory = "Н/Д";
                        M.MaterialMultiply = 0;
                        M.MaterialMultiplyString = "Н/Д";
                        M.WaitingDays = 1000;
                        M.WaitingDaysString = "Н/Д";
                        M.StorageQuantity = 0;
                        M.Responsible = "Н/Д";
                        M.StorageQuantityString = "Н/Д";
                        M.Sender = "Н/Д";
                        M.DontShow = true;
                    }

                    M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                    M.OldTaskNavUD = 0;

                    CFMList.Add(M);
                }
            }

            //остатки Mes
            var mml = fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            MesList = mml.Select(x => x.MaterialCode).Distinct().ToList();
            //  mml = mml.Where(x => x.MaterialCode == "8СИ00088").ToList();

            for (int i = MesList.Count - 1; i >= 0; i--)
            {
                var cfm = CFMList.Where(x => x.MaterialCode == MesList[i]).ToList();

                if (cfm.Count > 0)
                {
                    MesList.RemoveAt(i);
                }
            }

            foreach (string s in MesList)
            {
                var mesList = mml.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (mesList.Count > 0)
                {
                    var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                    Material M = new Material();

                    M.MaterialCode = s;
                    M.MaterialName = mesList[0].MaterialName;
                    if (MList.Count > 0)
                    {
                        M.MaterialGroup = MList[0].MaterialGroup;
                        M.MaterialCategory = MList[0].MaterialCategory;
                        M.MaterialMultiply = MList[0].MaterialMultiply;
                        M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                        M.WaitingDays = MList[0].WaitingDays;
                        M.WaitingDaysString = MList[0].WaitingDaysString;
                        M.StorageQuantity = MList[0].StorageQuantity;
                        M.StorageQuantityString = MList[0].StorageQuantityString;
                        M.Responsible = MList[0].Responsible;
                        M.Sender = MList[0].Sender;
                        M.DontShow = true;
                    }
                    else
                    {
                        M.MaterialGroup = "Н/Д";
                        M.MaterialCategory = "Н/Д";
                        M.MaterialMultiply = 0;
                        M.MaterialMultiplyString = "Н/Д";
                        M.WaitingDays = 1000;
                        M.WaitingDaysString = "Н/Д";
                        M.StorageQuantity = 0;
                        M.Responsible = "Н/Д";
                        M.StorageQuantityString = "Н/Д";
                        M.Sender = "Н/Д";
                        M.DontShow = true;
                    }

                    M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                    M.OldTaskNavUD = 0;

                    CFMList.Add(M);
                }                
            }

            //остатки НАВ
            var sbl = fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.Quantity > 0)).ToList();

            var sbp = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.Quantity > 0)).ToList();  //остатки пр-во

            // plant  4, 6, 8
            var sbpNonZero = sbp.Where(x => ((x.MaterialGroupID == 4) || (x.MaterialGroupID == 5) || (x.MaterialGroupID == 8)) ).ToList();

            List<string> PlanDataList = sbpNonZero.Select(x => x.MaterialCode).Distinct().ToList();

            for (int i= PlanDataList.Count-1; i>=0; i--)
            {
                var cfm = CFMList.Where(x => x.MaterialCode == PlanDataList[i]).ToList();

                if (cfm.Count > 0)
                {
                    PlanDataList.RemoveAt(i);
                }
            }

            foreach (string s in PlanDataList)
            {
                var pdList = sbpNonZero.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (pdList.Count > 0)
                {
                    var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                    Material M = new Material();

                    M.MaterialCode = s;
                    M.MaterialName = pdList[0].MaterialName;
                    if (MList.Count > 0)
                    {
                        M.MaterialGroup = MList[0].MaterialGroup;
                        M.MaterialCategory = MList[0].MaterialCategory;
                        M.MaterialMultiply = MList[0].MaterialMultiply;
                        M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                        M.WaitingDays = MList[0].WaitingDays;
                        M.WaitingDaysString = MList[0].WaitingDaysString;
                        M.StorageQuantity = MList[0].StorageQuantity;
                        M.StorageQuantityString = MList[0].StorageQuantityString;
                        M.Responsible = MList[0].Responsible;
                        M.Sender = MList[0].Sender;
                        M.DontShow = true;
                    }
                    else
                    {
                        M.MaterialGroup = "Н/Д";
                        M.MaterialCategory = "Н/Д";
                        M.MaterialMultiply = 0;
                        M.MaterialMultiplyString = "Н/Д";
                        M.WaitingDays = 1000;
                        M.WaitingDaysString = "Н/Д";
                        M.StorageQuantity = 0;
                        M.Responsible = "Н/Д";
                        M.StorageQuantityString = "Н/Д";
                        M.Sender = "Н/Д";
                        M.DontShow = true;
                    }

                    M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                    M.OldTaskNavUD = 0;

                    CFMList.Add(M);
                    MesList.Add(s);
                }
            }

            //plANT 7 (335, 336, 338, 346, 349, 375)
            sbpNonZero = sbp.Where(x => (x.MaterialGroupID == 7) && 
                ((x.MaterialTypeID == 335) || (x.MaterialTypeID == 336) || (x.MaterialTypeID == 338) || (x.MaterialTypeID == 346) || (x.MaterialTypeID == 359) || (x.MaterialTypeID == 375))).ToList();

            PlanDataList = sbpNonZero.Select(x => x.MaterialCode).Distinct().ToList();

            for (int i = PlanDataList.Count - 1; i >= 0; i--)
            {
                var cfm = CFMList.Where(x => x.MaterialCode == PlanDataList[i]).ToList();

                if (cfm.Count > 0)
                {
                    PlanDataList.RemoveAt(i);
                }
            }

            foreach (string s in PlanDataList)
            {
                var pdList = sbpNonZero.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (pdList.Count > 0)
                {
                    var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                    Material M = new Material();

                    M.MaterialCode = s;
                    M.MaterialName = pdList[0].MaterialName;
                    if (MList.Count > 0)
                    {
                        M.MaterialGroup = MList[0].MaterialGroup;
                        M.MaterialCategory = MList[0].MaterialCategory;
                        M.MaterialMultiply = MList[0].MaterialMultiply;
                        M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                        M.WaitingDays = MList[0].WaitingDays;
                        M.WaitingDaysString = MList[0].WaitingDaysString;
                        M.StorageQuantity = MList[0].StorageQuantity;
                        M.StorageQuantityString = MList[0].StorageQuantityString;
                        M.Responsible = MList[0].Responsible;
                        M.Sender = MList[0].Sender;
                        M.DontShow = true;
                    }
                    else
                    {
                        M.MaterialGroup = "Н/Д";
                        M.MaterialCategory = "Н/Д";
                        M.MaterialMultiply = 0;
                        M.MaterialMultiplyString = "Н/Д";
                        M.WaitingDays = 1000;
                        M.WaitingDaysString = "Н/Д";
                        M.StorageQuantity = 0;
                        M.Responsible = "Н/Д";
                        M.StorageQuantityString = "Н/Д";
                        M.Sender = "Н/Д";
                        M.DontShow = true;
                    }

                    M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                    M.OldTaskNavUD = 0;

                    CFMList.Add(M);
                    MesList.Add(s);
                }
            }

            //9 plant ( 175, 389)  
            //   sbpNonZero = sbp.Where(x => (x.MaterialGroupID == 9) && 
            //      ((x.MaterialTypeID == 175) || (x.MaterialTypeID == 389) )).ToList();
            //not like N'Отход%' and m.Name not like N'Репроцесс%'

            var NonZero = fm.edb.pr_GetAlarmWaste().Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            PlanDataList = NonZero.Select(x => x.MaterialCode).Distinct().ToList();

            for (int i = PlanDataList.Count - 1; i >= 0; i--)
            {
                var cfm = CFMList.Where(x => x.MaterialCode == PlanDataList[i]).ToList();

                if (cfm.Count > 0)
                {
                    PlanDataList.RemoveAt(i);
                }
            }

            foreach (string s in PlanDataList)
            {
                var pdList = NonZero.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (pdList.Count > 0)
                {
                    var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                    Material M = new Material();

                    M.MaterialCode = s;
                    M.MaterialName = pdList[0].MaterialName;
                    if (MList.Count > 0)
                    {
                        M.MaterialGroup = MList[0].MaterialGroup;
                        M.MaterialCategory = MList[0].MaterialCategory;
                        M.MaterialMultiply = MList[0].MaterialMultiply;
                        M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                        M.WaitingDays = MList[0].WaitingDays;
                        M.WaitingDaysString = MList[0].WaitingDaysString;
                        M.StorageQuantity = MList[0].StorageQuantity;
                        M.StorageQuantityString = MList[0].StorageQuantityString;
                        M.Responsible = MList[0].Responsible;
                        M.Sender = MList[0].Sender;
                        M.DontShow = true;
                    }
                    else
                    {
                        M.MaterialGroup = "Н/Д";
                        M.MaterialCategory = "Н/Д";
                        M.MaterialMultiply = 0;
                        M.MaterialMultiplyString = "Н/Д";
                        M.WaitingDays = 1000;
                        M.WaitingDaysString = "Н/Д";
                        M.StorageQuantity = 0;
                        M.Responsible = "Н/Д";
                        M.StorageQuantityString = "Н/Д";
                        M.Sender = "Н/Д";
                        M.DontShow = true;
                    }

                    M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                    M.OldTaskNavUD = 0;

                    CFMList.Add(M);

                    MesList.Add(s);
                }
            }        

            #region Hide
            //остатки Plant
            /*  List<string>  MesList1 = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).Select(x => x.MaterialCode).Distinct().ToList();

              for (int i = MesList1.Count - 1; i >= 0; i--)
              {
                  var cfm = CFMList.Where(x => x.MaterialCode == MesList1[i]).ToList();

                  if (cfm.Count > 0)
                  {
                      MesList1.RemoveAt(i);
                  }
              }

              foreach (string s in MesList1)
              {
                  var pList = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();
                  if (pList.Count > 0)
                  {
                      var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                      Material M = new Material();

                      M.MaterialCode = s;
                      M.MaterialName = pList[0].MaterialName;
                      if (MList.Count > 0)
                      {
                          M.MaterialGroup = MList[0].MaterialGroup;
                          M.MaterialCategory = MList[0].MaterialCategory;
                          M.MaterialMultiply = MList[0].MaterialMultiply;
                          M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                          M.WaitingDays = MList[0].WaitingDays;
                          M.WaitingDaysString = MList[0].WaitingDaysString;
                          M.StorageQuantity = MList[0].StorageQuantity;
                          M.StorageQuantityString = MList[0].StorageQuantityString;
                          M.Responsible = MList[0].Responsible;
                          M.Sender = MList[0].Sender;
                      }
                      else
                      {
                          M.MaterialGroup = "Н/Д";
                          M.MaterialCategory = "Н/Д";
                          M.MaterialMultiply = 0;
                          M.MaterialMultiplyString = "Н/Д";
                          M.WaitingDays = 1000;
                          M.WaitingDaysString = "Н/Д";
                          M.StorageQuantity = 0;
                          M.Responsible = "Н/Д";
                          M.StorageQuantityString = "Н/Д";
                          M.Sender = "Н/Д";
                      }

                      M.RawStorageList = new List<Raws>();
                      M.RawEnterpriseList = new List<Raws>();
                      M.RawNav1List = new List<Raws>();
                      M.RawNav2List = new List<Raws>();
                      M.RawMesOut1List = new List<Raws>();
                      M.RawMesOut2List = new List<Raws>();
                      M.OldTaskNavList = new List<Raws>();
                      M.TaskNavList1 = new List<Raws>();
                      M.TaskNavList2 = new List<Raws>();
                      M.OutList = new List<Raws>();
                      M.StorageQuantList = new List<Raws>();
                      M.ConsurmptionList = new List<Raws>();

                      M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                      M.OldTaskNavUD = 0;

                      CFMList.Add(M);
                  }
              }

              MesList.AddRange(MesList1);

              //остатки
             List<string>  MesList2 = fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).Select(x => x.MaterialCode).Distinct().ToList();

              for (int i = MesList2.Count - 1; i >= 0; i--)
              {
                  var cfm = CFMList.Where(x => x.MaterialCode == MesList2[i]).ToList();

                  if (cfm.Count > 0)
                  {
                      MesList2.RemoveAt(i);
                  }
              }

              foreach (string s in MesList2)
              {
                  var bList = fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();
                  if (bList.Count > 0)
                  {
                      var MList = fm.mList.Where(x => x.MaterialCode == s).ToList();

                      Material M = new Material();

                      M.MaterialCode = s;
                      M.MaterialName = bList[0].MaterialName;
                      if (MList.Count > 0)
                      {
                          M.MaterialGroup = MList[0].MaterialGroup;
                          M.MaterialCategory = MList[0].MaterialCategory;
                          M.MaterialMultiply = MList[0].MaterialMultiply;
                          M.MaterialMultiplyString = MList[0].MaterialMultiplyString;
                          M.WaitingDays = MList[0].WaitingDays;
                          M.WaitingDaysString = MList[0].WaitingDaysString;
                          M.StorageQuantity = MList[0].StorageQuantity;
                          M.StorageQuantityString = MList[0].StorageQuantityString;
                          M.Responsible = MList[0].Responsible;
                          M.Sender = MList[0].Sender;
                      }
                      else
                      {
                          M.MaterialGroup = "Н/Д";
                          M.MaterialCategory = "Н/Д";
                          M.MaterialMultiply = 0;
                          M.MaterialMultiplyString = "Н/Д";
                          M.WaitingDays = 1000;
                          M.WaitingDaysString = "Н/Д";
                          M.StorageQuantity = 0;
                          M.Responsible = "Н/Д";
                          M.StorageQuantityString = "Н/Д";
                          M.Sender = "Н/Д";
                      }

                      M.RawStorageList = new List<Raws>();
                      M.RawEnterpriseList = new List<Raws>();
                      M.RawNav1List = new List<Raws>();
                      M.RawNav2List = new List<Raws>();
                      M.RawMesOut1List = new List<Raws>();
                      M.RawMesOut2List = new List<Raws>();
                      M.OldTaskNavList = new List<Raws>();
                      M.TaskNavList1 = new List<Raws>();
                      M.TaskNavList2 = new List<Raws>();
                      M.OutList = new List<Raws>();
                      M.StorageQuantList = new List<Raws>();
                      M.ConsurmptionList = new List<Raws>();

                      M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                      M.OldTaskNavUD = 0;

                      CFMList.Add(M);
                  }
              }

              MesList.AddRange(MesList2);*/
            #endregion

            //итоги
            foreach (var ml in CFMList)
            {
              //  if (ml.MaterialCode == "1030001707")
              //  { }

                var sbList8 = sbl.Where(x => x.MaterialCode == ml.MaterialCode).ToList();    // fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList();//остатки НАВ
                //   sbList8 = sbList8.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();

                var sbListMes8 = mml.Where(x => x.MaterialCode == ml.MaterialCode).ToList();    // fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList(); //остатки склады
                                                                                                //  sbListMes8=sbListMes8.Where(x => x.TestQuality != "3Блок").ToList();

                var sbpList8 = sbp.Where(x => x.MaterialCode == ml.MaterialCode).ToList();   // fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList();  //остатки пр-во
                sbpList8 = sbpList8.Where(x => fm.StorageNamesList.Contains(x.LineName)).ToList();

                var sbListIn = sbList8.Where(x => x.StorageType == "Собств").ToList();  //Собств по НАВ
                var sbListOut = sbList8.Where(x => x.StorageType != "Собств").ToList(); //Внеш

                var sbListInBlock = sbListIn.Where(x => x.TestQualityGr == "3Блок").ToList(); //забл собств
                var sbListOutBlock = sbListOut.Where(x => x.TestQualityGr == "3Блок").ToList();  //забл Внеш

                var sList = fm.StorageList.Where(x => x.MaterialCode == ml.MaterialCode).OrderBy(x => x.PlanOperDate).ToList();
                //containers
                var cList = fm.ContList.Where(x => (x.MaterialCode == ml.MaterialCode)).OrderBy(x => x.ExpDT).ToList();

                var sListOld = sList.Where(x => x.PlanOperDate < DateTime.Today.Date).ToList();  //старые заявки НАВ
                var sListNew = sList.Where(x => (x.PlanOperDate >= DateTime.Today.Date) && ((x.StatusMZP.ToLower().Trim() == "заказано") || (x.StatusMZP.ToLower().Trim() == ""))).ToList(); //Новые заявки

                ml.RawStorageList = new List<Raws>();
                ml.RawEnterpriseList = new List<Raws>();
                ml.RawNav1List = new List<Raws>();
                ml.RawNav2List = new List<Raws>();
                ml.RawMesOut1List = new List<Raws>();
                ml.RawMesOut2List = new List<Raws>();
                ml.OldTaskNavList = new List<Raws>();
                ml.TaskNavList1 = new List<Raws>();
                ml.TaskNavList2 = new List<Raws>();
                ml.OutList = new List<Raws>();
                ml.StorageQuantList = new List<Raws>();
                ml.ConsurmptionList = new List<Raws>();
                ml.ListA = new List<Raws>();
                ml.ListB = new List<Raws>();
                ml.ListC = new List<Raws>();
                ml.ListD = new List<Raws>();

                ml.ListLine = new List<int>();
                ml.ListPlus = new List<int>();

                ml.RawStorage = sbListMes8.Sum(x => x.Quantity * -1) - (double)sbListInBlock.Sum(x => x.Quantity);  /*- sbListOutBlock.Sum(x => x.Quantity)*/   // sbList8.Sum(x => x.Quantity) - ...
                ml.RawEnterprise = (double)sbpList8.Sum(x => x.Quantity);
                ml.RawNav1 = (double)sbListOut.Sum(x => x.Quantity) - (double)sbListOutBlock.Sum(x => x.Quantity);  //с учетом заблокированных!
                ml.RawNav2 = (double)sbListIn.Sum(x => x.Quantity) - (double)sbListInBlock.Sum(x => x.Quantity);
                ml.RawMesOut1 = (double)sbListOutBlock.Sum(x => x.Quantity);
                ml.RawMesOut2 = (double)sbListInBlock.Sum(x => x.Quantity);
                ml.OldTaskNav = sListOld.Sum(x => (x.PlanQuantity - x.FactQuantity));

                ml.RawStorageUD = sbListMes8.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity * -1) - (double)sbListInBlock.Sum(x => x.Quantity) /*- sbListOutBlock.Sum(x => x.Quantity)*/;   //sbList8.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity)
                ml.RawEnterpriseUD = (double)sbpList8.Where(x => x.ValidTo >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawNav1UD = (double)sbListOut.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity) - (double)sbListOutBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawNav2UD = (double)sbListIn.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity) - (double)sbListInBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawMesOut1UD = (double)sbListOutBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawMesOut2UD = (double)sbListInBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.OldTaskNavUD = 0;

                var CL = ConsList.Where(x => x.MaterialCode == ml.MaterialCode).ToList();

                if (CL.Count > 0)
                {
                    ml.CurrentConsumption = CL.Sum(x => (double)x.Quantity);
                }

                foreach (var cl in CL)
                {
                    Raws r = new Raws
                    {
                        Storage = cl.JobName,
                        Quantity = (double)cl.Quantity
                    };

                    ml.ConsurmptionList.Add(r);
                }

                //raw Storage
                foreach (var rs in sbListMes8)  //остатки склад     sbListMes8
                {
                    bool falg = false;

                    foreach (var sb in sbListInBlock)
                    {

                        if (sb.LotName == rs.MaterialLot)
                        {
                            falg = true;
                        }
                    }
   
                    if (falg==false)
                    {
                        Raws r = new Raws
                        {
                            Storage = rs.Storage,
                            LotName = rs.MaterialLot,
                            MaterialCode=rs.MaterialCode,
                            MaterialName=rs.MaterialName,
                            LotDescr=rs.MaterialLot,                                                      
                            BBFDate = rs.BBFDate,
                            ProdDate = rs.ProdDate,
                            Quantity = (double)rs.Quantity*(-1)
                        };

                        ml.RawStorageList.Add(r);
                    }
                }

                try
                {
                    DateTime DT = ml.RawStorageList.Min(x => x.BBFDate);
                    if (DT.Year <= 2000)
                    {
                        ml.Color1 = "Violet";
                    }
                    else
                    {
                        if ((DT - DateTime.Today.Date).Days > 30)
                        {
                            ml.Color1 = "Green";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 20)
                        {
                            ml.Color1 = "Magenta";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 0)
                        {
                            ml.Color1 = "DarkOrange";
                        }
                        else
                        {
                            ml.Color1 = "Red";
                        }
                    }
                }
                catch (Exception xxx)
                { }
                //UnDelay
                try
                {
                    DateTime DT = ml.RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList().Min(x => x.BBFDate);

                    if (DT.Year <= 2000)
                    {
                        ml.Color1UD = "Violet";
                    }
                    else
                    {
                        if ((DT - DateTime.Today.Date).Days > 30)
                        {
                            ml.Color1UD = "Green";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 20)
                        {
                            ml.Color1UD = "Magenta";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 0)
                        {
                            ml.Color1UD = "DarkOrange";
                        }
                        else
                        {
                            ml.Color1UD = "Red";
                        }
                    }
                }
                catch (Exception xxx)
                { }

                //raw EnterPrise
                foreach (var rs1 in sbpList8)  //остатки пр-во
                {
                    Raws r = new Raws
                    {
                        Storage = rs1.LineName,
                        LotName = rs1.MaterialLotName,
                        MaterialCode = rs1.MaterialCode,
                        MaterialName = rs1.MaterialName,
                        LotDescr = rs1.AlterLotName,
                        BBFDate = rs1.ValidTo,
                        ProdDate = rs1.ValidFrom,
                        Quantity = (double)rs1.Quantity
                    };

                    ml.RawEnterpriseList.Add(r);
                }

                try
                {
                    DateTime DT = ml.RawEnterpriseList.Min(x => x.BBFDate);

                    if (DT.Year <= 2000)
                    {
                        ml.Color2 = "Violet";
                    }
                    else
                    {
                        if ((DT - DateTime.Today.Date).Days > 30)
                        {
                            ml.Color2 = "Green";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 20)
                        {
                            ml.Color2 = "Magenta";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 0)
                        {
                            ml.Color2 = "DarkOrange";
                        }
                        else
                        {
                            ml.Color2 = "Red";
                        }
                    }
                }
                catch (Exception xxx)
                { }

                //UnDelay
                try
                {
                    DateTime DT = ml.RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList().Min(x => x.BBFDate);

                    if (DT.Year <= 2000)
                    {
                        ml.Color2UD = "Violet";
                    }
                    else
                    {
                        if ((DT - DateTime.Today.Date).Days > 30)
                        {
                            ml.Color2UD = "Green";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 20)
                        {
                            ml.Color2UD = "Magenta";
                        }
                        else if ((DT - DateTime.Today.Date).Days > 0)
                        {
                            ml.Color2UD = "DarkOrange";
                        }
                        else
                        {
                            ml.Color2UD = "Red";
                        }
                    }
                }
                catch (Exception xxx)
                { }

                // остатки внешние
                foreach (var rs2 in sbListOut)
                {
                    bool falg = false;

                    foreach (var sb in sbListOutBlock)
                    {
                        if (sb.LotName == rs2.LotName)
                        {
                            falg = true;
                        }
                    }

                    if (falg == false)
                    {
                        Raws r = new Raws
                        {
                            Storage = rs2.Storage,
                            LotName = rs2.LotName,
                            LotDescr = rs2.LotDescription,
                            BBFDate = rs2.BBFDate,
                            ProdDate = rs2.ProdDate,
                            Quantity = (double)rs2.Quantity
                        };

                        ml.RawNav1List.Add(r);
                    }
                }

                // остатки по НАВ
                foreach (var rs3 in sbListIn)
                {
                    bool falg = false;

                    foreach (var sb in sbListInBlock)
                    {
                        if (sb.LotName == rs3.LotName)
                        {
                            falg = true;
                        }
                    }

                    if (falg == false)
                    {
                        Raws r = new Raws
                        {
                            Storage = rs3.Storage,
                            LotName = rs3.LotName,
                            LotDescr = rs3.LotDescription,
                            BBFDate = rs3.BBFDate,
                            ProdDate = rs3.ProdDate,
                            Quantity = (double)rs3.Quantity
                        };

                        ml.RawNav2List.Add(r);
                    }
                }

                foreach (var rs4 in sbListOutBlock)// Blocked внешние
                {
                    Raws r = new Raws
                    {
                        Storage = rs4.Storage,
                        LotName = rs4.LotName,
                        LotDescr = rs4.LotDescription,
                        BBFDate = rs4.BBFDate,
                        ProdDate = rs4.ProdDate,
                        Quantity = (double)rs4.Quantity
                    };

                    ml.RawMesOut1List.Add(r);
                }

                foreach (var rs5 in sbListInBlock)// Blocked Собств
                {
                    Raws r = new Raws
                    {
                        Storage = rs5.Storage,
                        LotName = rs5.LotName,
                        LotDescr = rs5.LotDescription,
                        BBFDate = rs5.BBFDate,
                        ProdDate = rs5.ProdDate,
                        Quantity = (double)rs5.Quantity
                    };

                    ml.RawMesOut2List.Add(r);
                }

                foreach (var rs6 in sListOld)// старые заявки
                {
                    if (rs6.PlanQuantity > rs6.FactQuantity)
                    {
                        Raws r = new Raws
                        {
                            Storage = rs6.OrderNymber,
                            BBFDate = rs6.PlanOperDate,
                            ProdDate = rs6.PlanDate,
                            Quantity = rs6.PlanQuantity - rs6.FactQuantity,
                            Initiator = rs6.Initiator,
                            StatusMZP = rs6.StatusMZP,
                            MZP=rs6.MZP
                        };

                        if ((r.BBFDate - DateTime.Today.Date).Days > 30)
                        {
                            r.Color = "Green";
                        }
                        else if ((r.BBFDate - DateTime.Today.Date).Days > 20)
                        {
                            r.Color = "Magenta";
                        }
                        else if ((r.BBFDate - DateTime.Today.Date).Days > 0)
                        {
                            r.Color = "DarkOrange";
                        }
                        else
                        {
                            r.Color = "Red";
                        }

                        ml.OldTaskNavList.Add(r);

                        if (rs6.StatusMZP.Trim()=="")
                        {
                            ml.ListA.Add(r);
                        }
                    }
                }

                foreach (var rs7 in sListNew)// новые заявки
                {
                    if (rs7.PlanQuantity > rs7.FactQuantity)
                    {
                        Raws r = new Raws
                        {
                            Storage = rs7.OrderNymber,
                            BBFDate = rs7.PlanOperDate,
                            ProdDate = rs7.PlanDate,
                            Quantity = rs7.PlanQuantity-rs7.FactQuantity,
                            AlterDate = new DateTime(rs7.PlanOperDate.Year, rs7.PlanOperDate.Month, 1),
                            AddDate = rs7.PlanOperDate.AddDays(1),
                            Initiator = rs7.Initiator,
                            StatusMZP = rs7.StatusMZP,
                            MZP=rs7.MZP
                        };

                        if ((r.BBFDate - DateTime.Today.Date).Days > 30)
                        {
                            r.Color = "Green";
                        }
                        else if ((r.BBFDate - DateTime.Today.Date).Days > 20)
                        {
                            r.Color = "Magenta";
                        }
                        else if ((r.BBFDate - DateTime.Today.Date).Days > 0)
                        {
                            r.Color = "DarkOrange";
                        }
                        else
                        {
                            r.Color = "Red";
                        }

                        if (rs7.StatusMZP.Trim() == "")
                        {
                            ml.TaskNavList1.Add(r);
                            ml.ListA.Add(r);
                        }
                        else
                        {
                            ml.TaskNavList2.Add(r);
                            ml.ListB.Add(r);  //заказ
                            ml.ListD.Add(r);
                        }
                    }
                }

                foreach (var cl in cList)
                {
                    Raws r = new Raws
                    {
                        Storage=cl.LotNo,
                        StatusMZP=cl.LotDescription,
                        Quantity=cl.Quantity,
                        Initiator="",
                        MZP="",
                        BBFDate=cl.ExpDT,
                        ProdDate=cl.ExpDT
                    };

                    ml.ListC.Add(r);
                    ml.ListD.Add(r);
                }
            }

            while (XlDataStart < XlDataEnd)
            {
              DateTime  CurrDate = XlDataStart.AddMonths(1).AddDays(-1);

                foreach (var ml in CFMList)
                {
                    var xData = XLDataList.Where(x => (x.DateWork >= XlDataStart) && (x.DateWork <= CurrDate) && (x.ProdCode == ml.MaterialCode)).ToList();

                   /* if (ml.MaterialCode=="3000106749")
                    {
                        if (XlDataStart == Convert.ToDateTime("2022-10-01"))
                        { }
                    }*/

                  double  Value = xData.Sum(x => x.Quantity);
                    ml.OutList.Add(new Raws
                    {
                        ProdDate = XlDataStart,
                        Quantity = Value
                    });
                }

                WorkMonthList.Add(XlDataStart);
                XlDataStart = CurrDate.AddDays(1);
            }

            DateTime CurrDate1 = DateTime.Today.Date;

            DTList = (from d in DTList
                      where d >= CurrDate1
                      select d).ToList();
            //строки для сетки
            foreach (var ml in CFMList)
            {                
                ml.Value_ = ml.CurrentConsumption + ml.RawNav1 + ml.RawStorageList.Sum(x => x.Quantity) + ml.RawEnterpriseList.Sum(x => x.Quantity);
                ml.ValueUD = ml.CurrentConsumption + ml.RawNav1
                         + ml.RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity)
                         + ml.RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);  //исключить просрочку

                sVal = ml.MaterialCode + RowSplitter + ml.MaterialName + RowSplitter +
                      ml.MaterialGroup + RowSplitter + ml.MaterialMultiplyString + RowSplitter +
                      ml.WaitingDaysString + RowSplitter + ml.StorageQuantityString + RowSplitter +
                      ml.Responsible + RowSplitter + ml.Sender + RowSplitter;

                for (int j = 0; j < WorkMonthList.Count; j++)
                {
                    if ((Math.Abs(ml.OutList[j].Quantity) <= 1) && (Math.Abs(ml.OutList[j].Quantity) > 0))
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N0") + RowSplitter;
                    }
                }


                if ((Math.Abs(ml.CurrentConsumption) <= 1) && (Math.Abs(ml.CurrentConsumption) > 0)) //8
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N0") + RowSplitter;
                }

                sValUD = sVal;

                if ((Math.Abs(ml.RawStorage) <= 1) && (ml.RawStorage != 0)) //9
                {
                    sVal = sVal + ml.RawStorage.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + ml.RawStorage.ToString("N0") + RowSplitter;
                }

                if ((Math.Abs(ml.RawEnterprise) <= 1) && (ml.RawEnterprise != 0))  //10
                {
                    sVal = sVal + ml.RawEnterprise.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + ml.RawEnterprise.ToString("N0") + RowSplitter;
                }

                if ((Math.Abs(ml.RawNav1) <= 1) && (ml.RawNav1 != 0)) //внешние 11
                {
                    sVal = sVal + ml.RawNav1.ToString("N1") + RowSplitter;   // " (";
                }
                else
                {
                    sVal = sVal + ml.RawNav1.ToString("N0") + RowSplitter;  //" (";
                }

                if ((Math.Abs(ml.RawNav2) <= 1) && (ml.RawNav2 != 0))  //собств по НАВ   12
                {
                    sVal = sVal + ml.RawNav2.ToString("N1") + RowSplitter;  //+ ") "
                }
                else
                {
                    sVal = sVal + ml.RawNav2.ToString("N0") + RowSplitter;  //+ ") "
                }

                if ((Math.Abs(ml.RawMesOut1) <= 1) && (ml.RawMesOut1 != 0)) //забл внешние 13
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N1") + RowSplitter; // " (";
                }
                else
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N0") + RowSplitter; // " (";
                }

                if ((Math.Abs(ml.RawMesOut2) <= 1) && (ml.RawMesOut2 != 0))   // забл  МЕС  14
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N1") + RowSplitter;  // + ")" 
                }
                else
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N0") + RowSplitter;  //+ ")"
                }

                if ((Math.Abs(ml.OldTaskNav) <= 1) && (ml.OldTaskNav != 0))  //15 старые заявки НАВ
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N1") + RowSplitter;   //0-15
                }
                else
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N0") + RowSplitter;   //0-15
                }

                //UD
                if ((Math.Abs(ml.RawStorageUD) <= 1) && (ml.RawStorageUD != 0))
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N1") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N0") + RowSplitter;
                }

                if ((Math.Abs(ml.RawEnterpriseUD) <= 1) && (ml.RawEnterpriseUD != 0))
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N1") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N0") + RowSplitter;
                }

                if ((Math.Abs(ml.RawNav1UD) <= 1) && (ml.RawNav1UD != 0))
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N1") + RowSplitter;  //+ " ("
                }
                else
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N0") + RowSplitter; // +" ("
                }

                if ((Math.Abs(ml.RawNav2UD) <= 1) && (ml.RawNav2UD != 0))
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N1") + RowSplitter;  // + ") "
                }
                else
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N0") + RowSplitter;  // +") "
                }

                if ((Math.Abs(ml.RawMesOut2UD) <= 1) && (ml.RawMesOut2UD != 0))
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N1") + RowSplitter;  //+ " ("
                }
                else
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N0") + RowSplitter;  //+ " ("
                }

                if ((Math.Abs(ml.RawMesOut1UD) <= 1) && (ml.RawMesOut1UD != 0))
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N1") + RowSplitter;  //") " +
                }
                else
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N0") + RowSplitter;  //+ ") "
                }

                if ((Math.Abs(ml.OldTaskNavUD) <= 1) && (ml.OldTaskNavUD != 0))
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N1") + RowSplitter;   //0-15
                }
                else
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N0") + RowSplitter;   //0-15
                }

                ml.ShortString = sVal;
                ml.ShortStringUD = sValUD;

                if (ml.DontShow==true)
                {
                    string V;
                    string vUD;
                    ml.ValueS = "";
                    ml.ValueUDS = "";
                    ml.AlterValueS = "";
                    ml.AlterValueUDS = "";

                    double Q = ml.Value_;
                    double QUD = ml.ValueUD;

                    if ((Math.Abs(ml.Value_) <= 1) && (ml.Value_ != 0))
                    {
                        V = ml.Value_.ToString("N1");
                    }
                    else
                    {
                        V = ml.Value_.ToString("N0");
                    }

                    if ((Math.Abs(ml.ValueUD) <= 1) && (ml.ValueUD != 0))
                    {
                        vUD = ml.ValueUD.ToString("N1");
                    }
                    else
                    {
                        vUD = ml.ValueUD.ToString("N0");
                    }

                    foreach (DateTime dt in DTList)
                    {
                        ml.AlterValueS = ml.AlterValueS + vUD + "" + RowSplitter;
                        ml.AlterValueUDS = ml.AlterValueUDS + V + "" + RowSplitter;
                    }

                    for (int k=0; k<DTList.Count; k++)
                    {
                        double T2 = 0;
                        double T1 = 0;

                        DateTime dt = DTList[k];

                        var FDTList = DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();

                        if ((Math.Abs(Q) <= 1) && (Q != 0))
                        {
                            V = Q.ToString("N1");
                        }
                        else
                        {
                            V = Q.ToString("N0");
                        }

                        if ((Math.Abs(QUD) <= 1) && (QUD != 0))
                        {
                            vUD = QUD.ToString("N1");
                        }
                        else
                        {
                            vUD = QUD.ToString("N0");
                        }

                        ml.ValueS = ml.ValueS +  vUD;
                        ml.ValueUDS = ml.ValueUDS + V;

                        List<Raws> tData2 = new List<Raws>();  //asks
                        if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = ml.TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                        }
                        else
                        {
                            tData2 = ml.TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                        }
                        T2 = T2 + tData2.Sum(x => x.Quantity);

                        List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();
                        if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == ml.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                        }
                        else
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == ml.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                        }
                        T2 = T2 + Cdata.Sum(x => x.Quantity);

                        List<Raws> tData1 = new List<Raws>(); //ask in progress
                        if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = ml.TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                        }
                        else
                        {
                            tData1 = ml.TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                        }
                        T1 = T1 + tData1.Sum(x => x.Quantity);

                        if (T2>0)
                        {
                            foreach (var t in tData2)
                            {
                                if ((Math.Abs(t.Quantity) <= 1) && (t.Quantity != 0))
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N1") + ") ";
                                }
                                else
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N0") + ") ";
                                }
                            }

                            ml.Green = true;

                            try
                            {
                                ml.ListPlus.Add(k + 16 + WorkMonthList.Count);
                            }
                            catch (Exception xx)
                            {
                                ml.ListPlus = new List<int>();
                                ml.ListPlus.Add(k + 16 + WorkMonthList.Count);
                            }

                            Q = Q + T2;
                            QUD = QUD + T2;
                        }

                        if (T1>0)
                        {
                            foreach (var t in tData1)
                            {
                                if ((Math.Abs(t.Quantity) <= 1) && (t.Quantity != 0))
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N1") + "] ";
                                }
                                else
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N0") + "] ";
                                }
                            }

                            ml.Pink = true;

                            try
                            {
                                ml.ListLine.Add(k + 16 + WorkMonthList.Count);
                            }
                            catch (Exception xx)
                            {
                                ml.ListLine = new List<int>();
                                ml.ListLine.Add(k + 16 + WorkMonthList.Count);
                            }

                         //   Q = Q + T1;
                         //   QUD = QUD + T1;
                        }

                        ml.ValueS = ml.ValueS + RowSplitter;
                        ml.ValueUDS = ml.ValueUDS + RowSplitter;
                    }
                }
            }
            
            CFMList = CFMList.OrderBy(x => x.MaterialName).ToList();  //Where(x => x.MaterialCode == "8СИ00009").ToList();       //

        //    var cffm = CFMList.Where(x => x.MaterialCode == "3000168517").ToList();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            dt1 = new DataTable();
            dateTimePicker1.Value = DateTime.Today.Date;

        //    RecalculateData();
            ClearDGV();

            //  Thread thread = new Thread(GetData);
            //  thread.Start();
        }


        private void GetData()
        {
            /*  Invoke(new Action(() =>
              {
                  fm.label1.Visible = true;
                  fm.progressBar1.Visible = true;
                  fm.label1.Text = "Загрузка данных";
                  fm.progressBar1.Value = 22;
              }));

              Invoke(new Action(() =>
              {
                  LoadData();
              }));

              Invoke(new Action(() =>
              {
                  fm.label1.Text = "Данные со складов";
                  fm.progressBar1.Value = 44;
              }));

              Invoke(new Action(() =>
              {
                  GetStorage();
              }));

              Invoke(new Action(() =>
              {
                  fll.Close();
                  fm.Enabled = false;
                  AppList = fm.tdb.tApplication.ToList();
                  //  cbUseMaterialPlanning.Checked = true;
                  //  cbUsingMaterial.Checked = true;

                  dataGridViewMain.ReadOnly = true;
                  dataGridViewMain.AllowUserToAddRows = false;
                  cbViews.Text = cbViews.Items[0].ToString();
                  dgvStorages.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
                  dgvFactis.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
              }));

              Invoke(new Action(() =>
              {
                  fm.label1.Text = "Заявки на материалы";
                  fm.progressBar1.Value = 66;
              }));

            //  Invoke(new Action(() =>
            //  {
                  CFMList = fm.mList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).OrderBy(x => x.MaterialCode).ToList();
            //  }));

              var ConsList = fm.edb.fn_select_FactConsurmption().ToList();

              Invoke(new Action(() =>
              {
                  foreach (var ml in CFMList)
                  {             
                      var sbList8 = fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList();
                      //   sbList8 = sbList8.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();

                      var sbpList8 = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList();
                      sbpList8 = sbpList8.Where(x => fm.StorageNamesList.Contains(x.LineName)).ToList();

                      var sbListIn = sbList8.Where(x => x.StorageType == "Собств").ToList();
                      var sbListOut = sbList8.Where(x => x.StorageType != "Собств").ToList();

                      var sbListInBlock = sbListIn.Where(x => x.TestQualityGr == "3Блок").ToList();
                      var sbListOutBlock = sbListOut.Where(x => x.TestQualityGr == "3Блок").ToList();

                      var sList = fm.StorageList.Where(x => x.MaterialCode == ml.MaterialCode).ToList();
                      var sListOld = sList.Where(x => x.PlanOperDate < DateTime.Today.Date).ToList();
                      var sListNew = sList.Where(x => (x.PlanOperDate >= DateTime.Today.Date) && ((x.StatusMZP.ToLower().Trim() == "заказано") || (x.StatusMZP.ToLower().Trim() == ""))).ToList();

                      ml.RawStorageList.Clear();
                      ml.RawEnterpriseList.Clear();
                      ml.RawNav1List.Clear();
                      ml.RawNav2List.Clear();
                      ml.RawMesOut1List.Clear();
                      ml.RawMesOut2List.Clear();
                      ml.OldTaskNavList.Clear();
                      ml.TaskNavList1.Clear();
                      ml.TaskNavList2.Clear();
                      ml.OutList.Clear();
                      ml.StorageQuantList.Clear();
                      ml.ConsurmptionList.Clear();

                      ml.RawStorage = (double)(sbList8.Sum(x => x.Quantity) - sbListInBlock.Sum(x => x.Quantity) );
                      ml.RawEnterprise = (double)sbpList8.Sum(x => x.Quantity);
                      ml.RawNav1 = (double)sbListOut.Sum(x => x.Quantity);
                      ml.RawNav2 = (double)sbListIn.Sum(x => x.Quantity);
                      ml.RawMesOut1 = (double)sbListOutBlock.Sum(x => x.Quantity);
                      ml.RawMesOut2 = (double)sbListInBlock.Sum(x => x.Quantity);
                      ml.OldTaskNav = sListOld.Sum(x => x.FactQuantity);


                      ml.RawStorageUD = (double)(sbList8.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity) - sbListInBlock.Sum(x => x.Quantity) );
                      ml.RawEnterpriseUD = (double)sbpList8.Where(x => x.ValidTo >= DateTime.Today.Date).Sum(x => x.Quantity);
                      ml.RawNav1UD = (double)sbListOut.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                      ml.RawNav2UD = (double)sbListIn.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                      ml.RawMesOut1UD = (double)sbListOutBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                      ml.RawMesOut2UD = (double)sbListInBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                      ml.OldTaskNavUD = 0;


                      var CL = ConsList.Where(x => x.MaterialCode == ml.MaterialCode).ToList();

                      if (CL.Count > 0)
                      {
                          ml.CurrentConsumption = CL.Sum(x => (double)x.Quantity);
                      }

                      foreach (var cl in CL)
                      {
                          Raws r = new Raws
                          {
                              Storage = cl.JobName,
                              Quantity = (double)cl.Quantity
                          };

                          ml.ConsurmptionList.Add(r);
                      }

                      foreach (var rs in sbList8)  //остатки склад
                      {
                          if (sbListInBlock.Contains(rs))
                          {

                          }
                          else
                          {
                              Raws r = new Raws
                              {
                                  Storage = rs.Storage,
                                  LotName = rs.LotName,
                                  //    LotDescr=rs.LotDescription,
                                  BBFDate = rs.BBFDate,
                                  ProdDate = rs.ProdDate,
                                  Quantity = (double)rs.Quantity
                              };

                              ml.RawStorageList.Add(r);
                          }
                      }

                      try
                      {
                          DateTime DT = ml.RawStorageList.Min(x => x.BBFDate);

                          if ((DT - DateTime.Today.Date).Days > 30)
                          {
                              ml.Color1 = "Green";
                          }
                          else if ((DT - DateTime.Today.Date).Days > 0)
                          {
                              ml.Color1 = "DarkOrange";
                          }
                          else
                          {
                              ml.Color1 = "Red";
                          }

                      }
                      catch (Exception xxx)
                      { }
                      //UnDelay
                      try
                      {
                          DateTime DT = ml.RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList().Min(x => x.BBFDate);

                          if ((DT - DateTime.Today.Date).Days > 30)
                          {
                              ml.Color1UD = "Green";
                          }
                          else if ((DT - DateTime.Today.Date).Days > 0)
                          {
                              ml.Color1UD = "DarkOrange";
                          }
                          else
                          {
                              ml.Color1UD = "Red";
                          }

                      }
                      catch (Exception xxx)
                      { }


                      foreach (var rs1 in sbpList8)  //остатки пр-во
                      {
                          Raws r = new Raws
                          {
                              Storage = rs1.LineName,
                              LotName = rs1.MaterialLotName,
                              //  LotDescr = rs1.,
                              BBFDate = rs1.ValidTo,
                              ProdDate = rs1.ValidFrom,
                              Quantity = (double)rs1.Quantity
                          };

                          ml.RawEnterpriseList.Add(r);
                      }

                      try
                      {
                          DateTime DT = ml.RawEnterpriseList.Min(x => x.BBFDate);

                          if ((DT - DateTime.Today.Date).Days > 30)
                          {
                              ml.Color2 = "Green";
                          }
                          else if ((DT - DateTime.Today.Date).Days > 0)
                          {
                              ml.Color2 = "DarkOrange";
                          }
                          else
                          {
                              ml.Color2 = "Red";
                          }

                      }
                      catch (Exception xxx)
                      { }

                      foreach (var rs2 in sbListOut)// остатки внешние
                      {
                          Raws r = new Raws
                          {
                              Storage = rs2.Storage,
                              LotName = rs2.LotName,
                              LotDescr = rs2.LotDescription,
                              BBFDate = rs2.BBFDate,
                              ProdDate = rs2.ProdDate,
                              Quantity = (double)rs2.Quantity
                          };

                          ml.RawNav1List.Add(r);
                      }

                      foreach (var rs3 in sbListIn)// остатки по НАВ
                      {
                          Raws r = new Raws
                          {
                              Storage = rs3.Storage,
                              LotName = rs3.LotName,
                              LotDescr = rs3.LotDescription,
                              BBFDate = rs3.BBFDate,
                              ProdDate = rs3.ProdDate,
                              Quantity = (double)rs3.Quantity
                          };

                          ml.RawNav2List.Add(r);
                      }

                      foreach (var rs4 in sbListOutBlock)// Blocked внешние
                      {
                          Raws r = new Raws
                          {
                              Storage = rs4.Storage,
                              LotName = rs4.LotName,
                              LotDescr = rs4.LotDescription,
                              BBFDate = rs4.BBFDate,
                              ProdDate = rs4.ProdDate,
                              Quantity = (double)rs4.Quantity
                          };

                          ml.RawMesOut1List.Add(r);
                      }

                      foreach (var rs5 in sbListInBlock)// Blocked Собств
                      {
                          Raws r = new Raws
                          {
                              Storage = rs5.Storage,
                              LotName = rs5.LotName,
                              LotDescr = rs5.LotDescription,
                              BBFDate = rs5.BBFDate,
                              ProdDate = rs5.ProdDate,
                              Quantity = (double)rs5.Quantity
                          };

                          ml.RawMesOut2List.Add(r);
                      }

                      foreach (var rs6 in sListOld)// старые заявки
                      {
                          Raws r = new Raws
                          {
                              Storage = rs6.OrderNymber,
                              BBFDate = rs6.PlanOperDate,
                              ProdDate = rs6.PlanDate,
                              Quantity = rs6.PlanQuantity
                          };

                          ml.OldTaskNavList.Add(r);
                      }

                      foreach (var rs7 in sListNew)// новые заявки
                      {
                          Raws r = new Raws
                          {
                              Storage = rs7.OrderNymber,
                              BBFDate = rs7.PlanOperDate,
                              ProdDate = rs7.PlanDate,
                              Quantity = rs7.PlanQuantity,
                              AlterDate = new DateTime(rs7.PlanOperDate.Year, rs7.PlanOperDate.Month, 1),
                              AddDate = rs7.PlanOperDate.AddDays(1)
                          };

                          if (rs7.StatusMZP.Trim() == "")
                          {
                              ml.TaskNavList1.Add(r);
                          }
                          else
                          {
                              ml.TaskNavList2.Add(r);
                          }
                      }
                  }
              }));

              Invoke(new Action(() =>
              {
                  GC.Collect();
                  GC.WaitForPendingFinalizers();
                  GC.Collect();
              }));

                  dt1 = new DataTable();
                  dateTimePicker1.Value = DateTime.Today.Date;


              Invoke(new Action(() =>
              {
                  fm.label1.Text = "Расчет данных";
                  fm.progressBar1.Value = 77;
              }));

              Invoke(new Action(() =>
              {
                  RecalculateData();
              }));
              ClearDGV();*/
        }

        private void GetCurrentConsurmption()
        {
            /* DateTime CurrD = DateTime.Today.Date.AddHours(8);
             var ConsList = fm.edb.fn_select_FactConsurmption().ToList();     //fm.edb.fn_select_CurrentMaterials(CurrD).ToList();

             foreach (var c in CFMList)
             {
                 var CL = ConsList.Where(x => x.MaterialCode == c.MaterialCode).ToList();

                 if (CL.Count>0)
                 {
                     c.CurrentConsumption = (double)CL[CL.Count - 1].Quantity;
                 }
             }
             GC.Collect();*/
        }

        private void RecalculateData()
        {
            // UpdateConsumables();
            ShowData();
        }

        private void ShowLocalData()
        {
            string MCode;
            double Value;
            string SValue;
          //  string[] OldString;
            string[] NewString;
            char RowSplitter = '|';

            List<string[]> TempList = new List<string[]>();

            CFMListGold.Clear();
            CFMListRed.Clear();
            CFMListMinus.Clear();
            CFMListGreen.Clear();
            CFMListPlus.Clear();
            CFMListPink.Clear();
            CFMListLine.Clear();
            CFMListGoldens.Clear();
            CFMListLinePlus.Clear();
            //rows

            foreach (DataGridViewRow r in dataGridViewMain.Rows)
            {
                MCode = r.Cells[0].Value.ToString();
                int ccc = r.Cells.Count;
                Int32 Cols = 16 + WorkMonthList.Count;   //14!!!!

                var mmm = CFMList.Where(x => x.MaterialCode == MCode).FirstOrDefault();
                var cfm = CFMList.Where(x => x.MaterialCode == MCode).FirstOrDefault();
                SValue = "";

                foreach (string[] s in RowList)
                {
                    if (s[0]==MCode)
                    {
                      //  OldString = s;

                        for (int i=0; i<Cols; i++)
                        {
                            SValue = SValue + s[i] + RowSplitter;
                        }
                    }
                }

                if (mmm!=null)
                {
                    //9-13
                    if (cbMaterialDelay.Checked == false)
                    {
                        if ((Math.Abs(cfm.RawStorage) <= 1) && (Math.Abs(cfm.RawStorage) > 0))
                        {
                            SValue = SValue + cfm.RawStorage.ToString("N1") + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawStorage.ToString("N0") + RowSplitter;
                        }

                        if ((Math.Abs(cfm.RawEnterprise) <= 1) &&(Math.Abs(cfm.RawEnterprise) > 0))
                        {
                            SValue = SValue + cfm.RawEnterprise.ToString("N1") + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawEnterprise.ToString("N0") + RowSplitter;
                        }

                        if ((Math.Abs(cfm.RawNav1) <= 1) && (Math.Abs(cfm.RawNav1) > 0))
                        {
                            SValue = SValue + cfm.RawNav1.ToString("N1") + " (";
                        }
                        else
                        {
                            SValue = SValue + cfm.RawNav1.ToString("N0") + " (";
                        }

                        if ((Math.Abs(cfm.RawNav2) <= 1) && (Math.Abs(cfm.RawNav2) >0))
                        {
                            SValue = SValue + cfm.RawNav2.ToString("N1") + ") " + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawNav2.ToString("N0") + ") " + RowSplitter;
                        }

                        if ((Math.Abs(cfm.RawMesOut2) <= 1) && (Math.Abs(cfm.RawMesOut2) > 0))
                        {
                            SValue = SValue + cfm.RawMesOut2.ToString("N1") + " (";
                        }
                        else
                        {
                            SValue = SValue + cfm.RawMesOut2.ToString("N0") + " (";
                        }

                        if ((Math.Abs(cfm.RawMesOut1) <= 1) && (Math.Abs(cfm.RawMesOut1) > 0))
                        {
                            SValue = SValue + cfm.RawMesOut1.ToString("N1") + ") " + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawMesOut1.ToString("N0") + ") " + RowSplitter;
                        }

                        if ((Math.Abs(cfm.OldTaskNav) <= 1) && (Math.Abs(cfm.OldTaskNav) > 0))
                        {
                            SValue = SValue + cfm.OldTaskNav.ToString("N1") + RowSplitter;   //0-12
                        }
                        else
                        {
                            SValue = SValue + cfm.OldTaskNav.ToString("N0") + RowSplitter;   //0-12
                        }
                    }
                    else
                    {
                        if ((Math.Abs(cfm.RawStorageUD) <= 1) && (Math.Abs(cfm.RawStorageUD) > 0))
                        {
                            SValue = SValue + cfm.RawStorageUD.ToString("N1") + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawStorageUD.ToString("N0") + RowSplitter;
                        }

                        if ((Math.Abs(cfm.RawEnterpriseUD) <= 1) && (Math.Abs(cfm.RawEnterpriseUD) > 0))
                        {
                            SValue = SValue + cfm.RawEnterpriseUD.ToString("N1") + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawEnterpriseUD.ToString("N0") + RowSplitter;
                        }

                        if ((Math.Abs(cfm.RawNav1UD) <= 1) && (Math.Abs(cfm.RawNav1UD) > 0))
                        {
                            SValue = SValue + cfm.RawNav1UD.ToString("N1") + " (";
                        }
                        else
                        {
                            SValue = SValue + cfm.RawNav1UD.ToString("N0") + " (";
                        }

                        if ((Math.Abs(cfm.RawNav2UD) <= 1) && (Math.Abs(cfm.RawNav2UD) > 0))
                        {
                            SValue = SValue + cfm.RawNav2UD.ToString("N1") + ") " + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawNav2UD.ToString("N0") + ") " + RowSplitter;
                        }

                        if ((Math.Abs(cfm.RawMesOut2UD) <= 1) && (Math.Abs(cfm.RawMesOut2UD) > 0))
                        {
                            SValue = SValue + cfm.RawMesOut2UD.ToString("N1") + " (";
                        }
                        else
                        {
                            SValue = SValue + cfm.RawMesOut2UD.ToString("N0") + " (";
                        }

                        if ((Math.Abs(cfm.RawMesOut1UD) <= 1) && (Math.Abs(cfm.RawMesOut1UD) > 0))
                        {
                            SValue = SValue + cfm.RawMesOut1UD.ToString("N1") + ") " + RowSplitter;
                        }
                        else
                        {
                            SValue = SValue + cfm.RawMesOut1UD.ToString("N0") + ") " + RowSplitter;
                        }

                        if ((Math.Abs(cfm.OldTaskNavUD) <= 1) && (Math.Abs(cfm.OldTaskNavUD) > 0))
                        {
                            SValue = SValue + cfm.OldTaskNavUD.ToString("N1") + RowSplitter;   //0-12
                        }
                        else
                        {
                            SValue = SValue + cfm.OldTaskNavUD.ToString("N0") + RowSplitter;   //0-12
                        }                      
                    }

                    if (cbViews.Text == "Остаток")
                    {
                        Value = cfm.CurrentConsumption;

                        if (cbMaterialDelay.Checked == false)
                        {
                            foreach (var sl in cfm.RawStorageList)
                            {
                                Value = Value + sl.Quantity;
                            }

                            foreach (var el in cfm.RawEnterpriseList)
                            {
                                Value = Value + el.Quantity;
                            }
                        }
                        else
                        {
                            foreach (var sl in cfm.RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                            {
                                Value = Value + sl.Quantity;
                            }

                            foreach (var el in cfm.RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                            {
                                Value = Value + el.Quantity;
                            }
                        }
                     
                        foreach (DateTime dt in DTList)
                        {
                            //   if (CFMList[i].MaterialCode == "1031008522")
                            //   { }

                            //income
                            double T2 = 0;
                            double T1 = 0;

                            var FDTList = DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();

                            if (cbUseMaterialPlanning.Checked == true)
                            {
                                List<Raws> tData2 = new List<Raws>();

                                if (cbAddDay.Checked == false)
                                {
                                    tData2 = cfm.TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                                }
                                else
                                {
                                    tData2 = cfm.TaskNavList2.Where(x => x.AddDate == dt).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                                {
                                    tData2 = cfm.TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                                }

                                foreach (var td2 in tData2)
                                {
                                    T2 = T2 + td2.Quantity;
                                }

                                List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                                if (cbAddDay.Checked == false)
                                {
                                    Cdata = fm.ContList.Where(x => (x.MaterialCode == cfm.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                                }
                                else
                                {
                                    Cdata = fm.ContList.Where(x => (x.MaterialCode == cfm.MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                                {
                                    Cdata = fm.ContList.Where(x => (x.MaterialCode == cfm.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                                }

                                foreach (var cd in Cdata)
                                {
                                    T2 = T2 + cd.Quantity;
                                }
                                Value = Value + T2;

                                //ask in progress
                                List<Raws> tData1 = new List<Raws>();

                                if (cbAddDay.Checked == false)
                                {
                                    tData1 = cfm.TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                                }
                                else
                                {
                                    tData1 = cfm.TaskNavList1.Where(x => x.AddDate == dt).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                                {
                                    tData1 = cfm.TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                                }

                                foreach (var td1 in tData1)
                                {
                                    T1 = T1 + td1.Quantity;
                                }
                                Value = Value + T1;
                            }

                            //credit
                            if (cbUsingMaterial.Checked == true)
                            {
                                var dll = XLDataList.Where(x => (x.ProdCode == cfm.MaterialCode) && (x.DateWork.Date == dt)).ToList();
                                foreach (var dl in dll)
                                {
                                    Value = Value - dl.Quantity;
                                }
                            }

                            //future planning Asks
                            double Q = 0;

                            var adata = AppList;    // fm.tdb.tApplication.ToList();
                                                    //  if (adata.Count > 0)
                            {
                                adata = adata.Where(x => x.MaterialCode == cfm.MaterialCode).ToList();
                                //    if (adata.Count > 0)
                                {
                                    adata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                    if ((adata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                    {
                                        adata = AppList;
                                        adata = adata.Where(x => x.MaterialCode == cfm.MaterialCode).ToList();
                                        adata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                    }

                                    foreach (var ad in adata)
                                    {
                                        Q = Q + ad.Quantity;
                                    }

                                  /*  USCList.Add(new UserSelectedCell
                                    {
                                        RowIndex = i,
                                        ColIndex = Cols,
                                        Quantity = Q
                                    });*/
                                    //  Value = Value + Q;
                                }
                            }

                            if (Value < 0)
                            {
                              //  if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                {
                                    //  CFMBool[i] = false;
                                    if (!CFMListGold.Contains(cfm.MaterialCode))  //-
                                    {
                                        CFMListGold.Add(cfm.MaterialCode);
                                        CFMListRed.Add(cfm.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = CFMListMinus.Where(x => x.ProdCode == cfm.MaterialCode).ToList();
                                        try
                                        {
                                            Coll[0].ColNo.Add(Cols);
                                        }
                                        catch (Exception xx)
                                        {
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cfm.MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            CFMListMinus.Add(Col);
                                        }
                                    }
                                }
                            }

                            if ((Math.Abs(Value) <= 1) && (Math.Abs(Value) > 0))
                            {
                                SValue = SValue + Value.ToString("N1") + " "; // sVal = sVal + Value.ToString("N1") + " ";
                            }
                            else
                            {
                                SValue = SValue + Value.ToString("N0") + " ";    // sVal = sVal + Value.ToString("N0") + " ";
                            }

                            if (T2 != 0)
                            {
                                if ((Math.Abs(T2) <= 1) && (Math.Abs(T2) > 0))
                                {
                                    SValue = SValue + " (" + T2.ToString("N1") + ") ";
                                }
                                else
                                {
                                    SValue = SValue + " (" + T2.ToString("N0") + ") ";
                                }
                                if (!CFMListGreen.Contains(cfm.MaterialCode))  //+
                                {
                                    CFMListGreen.Add(cfm.MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListPlus.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListPlus.Where(x => x.ProdCode == cfm.MaterialCode).ToList();
                                    try
                                    {
                                        Coll[0].ColNo.Add(Cols); ;
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListPlus.Add(Col);
                                    }
                                }
                            }
                            if (T1 != 0)
                            {
                                if ((Math.Abs(T1) <= 1) && (Math.Abs(T1) > 0))
                                {
                                    SValue = SValue + " [" + T1.ToString("N1") + "] ";
                                }
                                else
                                {
                                    SValue = SValue + " [" + T1.ToString("N0") + "] ";
                                }

                                if (!CFMListPink.Contains(cfm.MaterialCode))  //[
                                {
                                    CFMListPink.Add(cfm.MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListLine.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListLine.Where(x => x.ProdCode == cfm.MaterialCode).ToList();
                                    try
                                    {
                                        Coll[0].ColNo.Add(Cols);
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListLine.Add(Col);
                                    }
                                }

                                if (CFMListGreen.Contains(cfm.MaterialCode))  //[+
                                {
                                    if (!CFMListGoldens.Contains(cfm.MaterialCode))
                                    {
                                        CFMListGoldens.Add(cfm.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListLinePlus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = CFMListLinePlus.Where(x => x.ProdCode == cfm.MaterialCode).ToList();
                                        try
                                        {
                                            Coll[0].ColNo.Add(Cols);
                                        }
                                        catch (Exception xx)
                                        {
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cfm.MaterialCode,
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
                                if ((Math.Abs(Q) <= 1) && (Math.Abs(Q) > 0))
                                {
                                    SValue = SValue + " {" + Q.ToString("N1") + "} ";
                                }
                                else
                                {
                                    SValue = SValue + " {" + Q.ToString("N0") + "} ";
                                }
                            }

                            SValue = SValue + "" + RowSplitter;
                           // r.Cells[Cols].Value = SValue;
                            Cols = Cols + 1;
                        }//foreach
                    }  //остаток
                    else if (cbViews.Text == "Потребление")
                    {
                        var xlData = XLDataList.Where(x => x.ProdCode == cfm.MaterialCode).ToList();

                        foreach (DateTime dt in DTList)
                        {
                            var xlD = xlData.Where(x => x.DateWork.Date == dt).ToList();

                            if (xlD.Count > 0)
                            {
                                if ((Math.Abs(xlD[0].Quantity) <= 1) && (Math.Abs(xlD[0].Quantity) > 0))
                                {
                                    SValue = SValue + xlD[0].Quantity.ToString("N1") + RowSplitter;     // sVal = sVal + xlD[0].Quantity.ToString("N1") + RowSplitter;
                                }
                                else
                                {
                                    SValue = SValue + xlD[0].Quantity.ToString("N0") + RowSplitter;
                                }
                            }
                            else
                            {
                                SValue = SValue + "" + RowSplitter;      //sVal = sVal + "" + RowSplitter;
                            }
                         //   r.Cells[Cols].Value = SValue;
                            Cols = Cols + 1;
                        }
                    }
                    else  //Дефицит
                    {
                        Value = 0;

                        if (cbMaterialDelay.Checked == false)
                        {
                            foreach (var sl in cfm.RawStorageList)
                            {
                                Value = Value + sl.Quantity;
                            }

                            foreach (var el in cfm.RawEnterpriseList)
                            {
                                Value = Value + el.Quantity;
                            }
                        }
                        else
                        {
                            foreach (var sl in cfm.RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                            {
                                Value = Value + sl.Quantity;
                            }

                            foreach (var el in cfm.RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                            {
                                Value = Value + el.Quantity;
                            }
                        }

                      //  Cols = 16 + WorkMonthList.Count;
                        foreach (DateTime dt in DTList)
                        {
                            var FDTList = DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();
                            //income
                            List<Raws> tData2 = new List<Raws>();
                            if (cbAddDay.Checked == false)
                            {
                                tData2 = cfm.TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                            }
                            else
                            {
                                tData2 = cfm.TaskNavList2.Where(x => x.AddDate == dt).ToList();
                            }

                            double T2 = 0;
                            if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                            {
                                tData2 = cfm.TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                            }

                            foreach (var td2 in tData2)
                            {
                                T2 = T2 + td2.Quantity;
                            }

                            List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                            if (cbAddDay.Checked == false)
                            {
                                Cdata = fm.ContList.Where(x => (x.MaterialCode == cfm.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                            }
                            else
                            {
                                Cdata = fm.ContList.Where(x => (x.MaterialCode == cfm.MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                            }

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                            {
                                Cdata = fm.ContList.Where(x => (x.MaterialCode == cfm.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                            }

                            foreach (var cd in Cdata)
                            {
                                T2 = T2 + cd.Quantity;
                            }
                            Value = Value + T2;

                            //ask in progress
                            List<Raws> tData1 = new List<Raws>();

                            if (cbAddDay.Checked == false)
                            {
                                tData1 = cfm.TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                            }
                            else
                            {
                                tData1 = cfm.TaskNavList1.Where(x => x.AddDate == dt).ToList();
                            }
                            double T1 = 0;

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                            {
                                tData1 = cfm.TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                            }

                            foreach (var td1 in tData1)
                            {
                                T1 = T1 + td1.Quantity;
                            }
                            Value = Value + T1;

                            //credit
                            var dll = XLDataList.Where(x => (x.ProdCode == cfm.MaterialCode) && (x.DateWork.Date == dt)).ToList();
                            foreach (var dl in dll)
                            {
                                Value = Value - dl.Quantity;
                            }

                            //future planning Asks
                            double Q = 0;
                            var adata = AppList;    // fm.tdb.tApplication.ToList();
                                                    //  if (adata.Count > 0)
                            {
                                adata = adata.Where(x => x.MaterialCode == cfm.MaterialCode).ToList();
                                //    if (adata.Count > 0)
                                {
                                    adata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                    if ((adata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                    {
                                        adata = AppList;
                                        adata = adata.Where(x => x.MaterialCode == cfm.MaterialCode).ToList();
                                        adata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                    }

                                    foreach (var ad in adata)
                                    {
                                        Q = Q + ad.Quantity;
                                    }

                                   /* USCList.Add(new UserSelectedCell
                                    {
                                        RowIndex = i,
                                        ColIndex = Cols,
                                        Quantity = Q
                                    });*/
                                    //  Value = Value + Q;
                                }
                            }

                            if (Value < 0)
                            {
                               // if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                {
                                    //  CFMBool[i] = false;
                                    if (!CFMListGold.Contains(cfm.MaterialCode))  //-
                                    {
                                        CFMListGold.Add(cfm.MaterialCode);
                                        CFMListRed.Add(cfm.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = CFMListMinus.Where(x => x.ProdCode == cfm.MaterialCode).ToList();
                                        try
                                        {
                                            Coll[0].ColNo.Add(Cols);
                                        }
                                        catch (Exception xx)
                                        {
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cfm.MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            CFMListMinus.Add(Col);
                                        }
                                    }
                                }

                                if ((Math.Abs(Value) <= 1) && (Math.Abs(Value) > 0))
                                {
                                    SValue = SValue + Value.ToString("N1");
                                }
                                else
                                {
                                    SValue= SValue + Value.ToString("N0") ;
                                }
                            }
                            SValue = SValue + "" + RowSplitter;
                        }  //foreach
                    }//deficite
                   
                    NewString = SValue.Split(RowSplitter);
                    //  r.Cells[Cols].Value = SValue;
                    Cols = Cols + 1;
                    TempList.Add(NewString);  // sVal.Split(RowSplitter);
                }
            }

            dataGridViewMain.DataSource = null;
            dataGridViewMain.Rows.Clear();
            dataGridViewMain.Columns.Clear();

            dataGridViewMain.Columns.Clear();
            dataGridViewMain.Columns.Add("MaterialCode", "Код материала"); //0
            dataGridViewMain.Columns.Add("MaterialName", "Наименование"); //1
            dataGridViewMain.Columns.Add("MaterialClass", "Класс материала"); //2 
            dataGridViewMain.Columns.Add("Cratn", "Кратность поставки");//3
            dataGridViewMain.Columns.Add("DayCount", "Мин срок пост. дней");//4
            dataGridViewMain.Columns.Add("SureStorage", "Страховой запас");//5
            dataGridViewMain.Columns.Add("User", "Исполнитель");//6
            dataGridViewMain.Columns.Add("Sender", "Поставщик");//7
            //Month need insert
            foreach (var ddd in WorkMonthList)
            {
                dataGridViewMain.Columns.Add(ddd.ToString("MM_yyyy"), "Потребность на " + ddd.ToString("MM.yyyy"));
            }

            dataGridViewMain.Columns.Add("Used", "Уже потреблено");//8
            dataGridViewMain.Columns.Add("Storage", "Остатки склады");//9
            dataGridViewMain.Columns.Add("Enterprise", "Остатки пр-во");//10
            dataGridViewMain.Columns.Add("OutStorage", "Внешние остатки (Собств по НАВ)");//11
            dataGridViewMain.Columns.Add("Blocked", "Заблокировано (Внешние МЕС)");//12
            dataGridViewMain.Columns.Add("OldTaskNav", "Старые заявки НАВ");//13

            if (cbViews.Text == "Остаток")
            {
                foreach (DateTime dt in DTList)
                {
                    dataGridViewMain.Columns.Add(dt.ToString("dd_MM_yyyy"), "Остаток на " + dt.ToString("dd.MM.yyyy"));
                }
            }
            else if (cbViews.Text == "Потребление")
            {
                foreach (DateTime dt in DTList)
                {
                    dataGridViewMain.Columns.Add(dt.ToString("dd_MM_yyyy"), "Потребление на " + dt.ToString("dd.MM.yyyy"));
                }
            }
            else if (cbViews.Text == "Дефицит")
            {
                foreach (DateTime dt in DTList)
                {
                    dataGridViewMain.Columns.Add(dt.ToString("dd_MM_yyyy"), "Дефицит на " + dt.ToString("dd.MM.yyyy"));
                }
            }

            dataGridViewMain.Columns.Add("Zerrrro", "Zerrrro");


            RowList.Clear();
            RowList = TempList;

            foreach (var sr in RowList)
            {
                dataGridViewMain.Rows.Add(sr);
            }

            dataGridViewMain.Columns[0].Width = 80;
            dataGridViewMain.Columns[1].Width = 300;
            dataGridViewMain.Columns[2].Width = 120;
            dataGridViewMain.Columns[6].Width = 100;
            dataGridViewMain.Columns[7].Width = 100;
            dataGridViewMain.Columns[dataGridViewMain.ColumnCount - 1].Visible = false;

            //   DGV_UpdateColumnZero();
            //   DGV_Get_CellColor();

            for (int i = 0; i < 2; i++)
            {
                dataGridViewMain.Columns[i].Frozen = true;
            }

            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;
            dataGridViewMain.DoubleBuffered = true;
            dataGridViewMain.Update();
            //cellcolor

            foreach (DataGridViewColumn c in dataGridViewMain.Columns)
            {
                //c.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void ShowData()
        {
            //очистить таблицы
            dgvPlanned.Rows.Clear();
            dgvRecipes.Rows.Clear();
            dgvUsing.Rows.Clear();
            dgvFactis.Rows.Clear();
            dgvStorages.Rows.Clear();
            dgvComment.DataSource = null;
            dgvComment.Rows.Clear();

            ClearButtons();

            dataGridViewMain.DataSource = null;
            dataGridViewMain.Rows.Clear();
            dataGridViewMain.Columns.Clear();
          //  dataGridViewMain.DoubleBuffered = false;
            //  WorkMonthList.Clear();
            MaterialData = new List<DailyResult>();
            FilteredMaterialDataIn = new List<DailyResult>();
            FilteredMaterialDataOut = new List<DailyResult>();
            List<string> AskList = new List<string>();

            CurrCode = "";

            string[] NewString;
            char RowSplitter = '|';
            string sVal = "";
            DateTime CurrDate;
            bool fl;
            Int32 J;
            double Value = 0;
            DateTime XlDataStart = fm.ExcelDataList.Min(x => x.DateWork);
            XlDataStart = new DateTime(XlDataStart.Year, XlDataStart.Month, 1);
            DateTime XlDataEnd = fm.ExcelDataList.Max(x => x.DateWork);
            XlDataEnd = new DateTime(XlDataEnd.Year, XlDataEnd.Month, 1);
            XlDataEnd = XlDataEnd.AddMonths(1).AddDays(-1);
            if (RowList == null)
            {
                RowList = new List<string[]>();
            }

            AppList = fm.tdb.tApplication.ToList();

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

            Cursor = Cursors.WaitCursor;
            Int32 Cols;

            dt1 = new DataTable();
         //   dt1.Columns.Clear();
         //   dt1.Rows.Clear();

            dt1.Columns.Add("Код материала");  //0
            dt1.Columns.Add("Наименование");  //1
            dt1.Columns.Add("Класс материала"); //2
            dt1.Columns.Add("Кратность поставки");  //3
            dt1.Columns.Add("Мин срок пост. дней");//4
            dt1.Columns.Add("Страховой запас");//5
            dt1.Columns.Add("Исполнитель");//6
            dt1.Columns.Add("Поставщик");//7
                                         //Month need insert
            foreach (var ddd in WorkMonthList)
            {
                dt1.Columns.Add("Потребность на " + ddd.ToString("MM.yyyy"));
            }
            dt1.Columns.Add("Уже потреблено");//8
            dt1.Columns.Add("Остатки склады");//9
            dt1.Columns.Add("Остатки пр-во");//10
            dt1.Columns.Add("Внешние остатки");//11   (Собств по НАВ)
            dt1.Columns.Add("Остатки по НАВ");//12
            dt1.Columns.Add("Заблокировано: Внешние");//13  (МЕС)
            dt1.Columns.Add("Заблокировано: МЕС");//14
            dt1.Columns.Add("Старые заявки НАВ");//15

            CurrDate = DateTime.Today.Date;

            DTList = (from d in DTList
                      where d >= CurrDate
                      select d).ToList();

            if (cbViews.Text == "Остаток")
            {
                foreach (DateTime dt in DTList)
                {
                    dt1.Columns.Add("Остаток на " + dt.ToString("dd.MM.yyyy"));  // dt.ToString("dd_MM_yyyy"),
                }
            }
            else if (cbViews.Text == "Потребление")
            {
                foreach (DateTime dt in DTList)
                {
                    dt1.Columns.Add("Потребление на " + dt.ToString("dd.MM.yyyy"));
                }
            }
            else if (cbViews.Text == "Дефицит")
            {
                foreach (DateTime dt in DTList)
                {
                    dt1.Columns.Add("Дефицит на " + dt.ToString("dd.MM.yyyy"));
                }
            }

            dt1.Columns.Add("Zerrrro");

            Int32 ccc = dataGridViewMain.Columns.Count;

            MaterialData = CFMList.Select(x => new DailyResult
            {
                ProductCode = x.MaterialCode,
                LineName = x.MaterialName,
                Quantity = 1,
                MaterialType = x.MaterialGroup.Trim()
            }).ToList();
            MaterialData = MaterialData.Distinct().ToList();

            if (DataLoaded == false)
            {
                for (int i = CFMList.Count - 1; i >= 0; i--)
                {
                    foreach (var t in TBP)
                    {
                        if (CFMList.Count > 0)
                        {
                            if (CFMList[i].MaterialCode == t.ProductCode)
                            {
                                CFMList.RemoveAt(i);
                            }
                        }
                    }
                }
            }

            for (int i = 0; i < CFMList.Count; i++)
            {   //0-8
                var cfm1 = CFMList[i];
        
                if (cbMaterialDelay.Checked == false)
                {
                    Value = cfm1.Value_;
                    sVal = cfm1.ShortString;
                }
                else
                {
                    Value = cfm1.ValueUD;
                    sVal = cfm1.ShortStringUD;
                }


                if ((cfm1.WaitingDaysString == "") || (cfm1.MaterialMultiplyString == "") /*||(cfm1.StorageQuantityString=="")*/)
                {
                    CFMListGray.Add(cfm1.MaterialCode);
                }

                if (cbViews.Text == "Остаток")
                {
                    Cols = 16 + WorkMonthList.Count;  //14

                    if (cfm1.DontShow == true)
                    {
                        if (cbUseMaterialPlanning.Checked == true)
                        {
                            if (cbMaterialDelay.Checked == true)
                            {
                                sVal = sVal + cfm1.ValueS;
                            }
                            else
                            {
                                sVal = sVal + cfm1.ValueUDS;
                            }

                            if (cfm1.Green == true)
                            {
                                CFMListGreen.Add(cfm1.MaterialCode);

                                foreach (var p in cfm1.ListPlus)
                                {
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm1.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(p);
                                    CFMListPlus.Add(Col);
                                }
                            }
                            if (cfm1.Pink == true)
                            {
                                CFMListPink.Add(cfm1.MaterialCode);

                                foreach (var p in cfm1.ListLine)
                                {
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm1.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(p);
                                    CFMListLine.Add(Col);
                                }
                            }
                        }
                        else
                        {
                            if (cbMaterialDelay.Checked == true)
                            {
                                sVal = sVal + cfm1.AlterValueS;
                            }
                            else
                            {
                                sVal = sVal + cfm1.AlterValueUDS;
                            }
                        }
                    }
                    else
                    {
                        double OldValue = 0;
                        string StringValue = "0";
                        var xlData = XLDataList.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                        
                        var adata = AppList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();    // AppList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();
                        var cListData = fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode)).ToList();

                        // foreach (DateTime dt in DTList)
                        for (int j = 0; j < DTList.Count; j++)
                        {
                            //income
                            double T2 = 0;
                            double T1 = 0;
                            DateTime dt = DTList[j];

                            var FDTList = DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();

                            if (cbUseMaterialPlanning.Checked == true)
                            {
                                List<Raws> tData2 = new List<Raws>();

                                if (cbAddDay.Checked == false)
                                {
                                    tData2 = cfm1.TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                                }
                                else
                                {
                                    tData2 = cfm1.TaskNavList2.Where(x => x.AddDate == dt).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                                {
                                    tData2 = cfm1.TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                                }

                                //  foreach (var td2 in tData2)
                                //  {
                                T2 = T2 + tData2.Sum(x => x.Quantity);     // td2.Quantity;
                                                                           //}

                                List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                                if (cbAddDay.Checked == false)
                                {
                                    Cdata = cListData.Where(x => x.ExpDT.Date == dt.Date).ToList();   // fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                                }
                                else
                                {
                                    Cdata = cListData.Where(x => x.ExpDT.Date.AddDays(1) == dt.Date).ToList();        //fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                                {
                                    Cdata = cListData.Where(x => ((DateTime)x.AlterDate).Date == dt.Date).ToList();        // fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                                }

                                // foreach (var cd in Cdata)
                                // {
                                //     T2 = T2 + cd.Quantity;
                                // }
                                T2 = T2 + Cdata.Sum(x => x.Quantity);

                                if (T2 > 0)
                                {
                                    if (Value < 0)
                                    {
                                        Value = 0;
                                    }

                                    if (!AskList.Contains(cfm1.MaterialCode))
                                    {
                                        AskList.Add(cfm1.MaterialCode);
                                    }
                                }

                                Value = Value + T2;

                                //ask in progress
                                List<Raws> tData1 = new List<Raws>();

                                if (cbAddDay.Checked == false)
                                {
                                    tData1 = cfm1.TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                                }
                                else
                                {
                                    tData1 = cfm1.TaskNavList1.Where(x => x.AddDate == dt).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                                {
                                    tData1 = cfm1.TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                                }

                                //   foreach (var td1 in tData1)
                                //   {
                                //       T1 = T1 + td1.Quantity;
                                //   }
                                T1 = T1 + tData1.Sum(x => x.Quantity);

                                if (T1 > 0)  //  уточнить у каменской
                                {
                                    if (!AskList.Contains(cfm1.MaterialCode))
                                    {
                                        AskList.Add(cfm1.MaterialCode);
                                    }

                                   /* if (Value < 0)
                                    {
                                        Value = 0;
                                    }*/
                                }

                                //Value = Value + T1;
                            }

                            //credit
                            if (cbUsingMaterial.Checked == true)
                            {
                                var dll = xlData.Where(x => x.DateWork.Date == dt).ToList();      //XLDataList.Where(x => (x.ProdCode == cfm1.MaterialCode) && (x.DateWork.Date == dt)).ToList();
                                //  foreach (var dl in dll)
                                //  {
                                //      Value = Value - dl.Quantity;
                                //  }
                                Value = Value - dll.Sum(x => x.Quantity);
                            }

                            //future planning Asks
                            double Q = 0;
                            var bdata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();
                            if ((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                            {
                                bdata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                            }
                            Q = Q + bdata.Sum(x => x.Quantity);

                            USCList.Add(new UserSelectedCell
                            {
                                RowIndex = i,
                                ColIndex = Cols,
                                Quantity = Q
                            });
                            //  Value = Value + Q;

                            if (Value < 0)
                            {
                                // if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                {
                                    if (AskList.Contains(cfm1.MaterialCode))
                                    {
                                        for (int k = AskList.Count - 1; k >= 0; k--)
                                        {
                                            if (AskList[k] == cfm1.MaterialCode)
                                            {
                                                AskList.RemoveAt(k);
                                                k = 0;
                                            }
                                        }
                                    }

                                    //  CFMBool[i] = false;
                                    if (!CFMListGold.Contains(cfm1.MaterialCode))  //-
                                    {
                                        CFMListGold.Add(cfm1.MaterialCode);
                                        CFMListRed.Add(cfm1.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm1.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = CFMListMinus.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                                        try
                                        {
                                            if (!Coll[0].ColNo.Contains(Cols))
                                            {
                                                Coll[0].ColNo.Add(Cols);
                                            }
                                        }
                                        catch (Exception xx)
                                        {
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cfm1.MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            CFMListMinus.Add(Col);
                                        }
                                    }
                                }
                            }

                            if (Value != OldValue)
                            {
                                OldValue = Value;

                                if ((Math.Abs(Value) <= 1) && (Value != 0))
                                {
                                    StringValue = Value.ToString("N1") + " ";
                                }
                                else
                                {
                                    StringValue = Value.ToString("N0") + " ";
                                }
                            }
                            sVal = sVal + StringValue;

                            /* if ((Math.Abs(Value) <= 1) && (Value != 0))
                             {
                                 sVal = sVal + Value.ToString("N1") + " ";
                             }
                             else
                             {
                                 sVal = sVal + Value.ToString("N0") + " ";
                             }*/

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
                                if (!CFMListGreen.Contains(cfm1.MaterialCode))  //+
                                {
                                    CFMListGreen.Add(cfm1.MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm1.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListPlus.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListPlus.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                                    try
                                    {
                                        if (!Coll[0].ColNo.Contains(Cols))
                                        {
                                            Coll[0].ColNo.Add(Cols);
                                        }
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm1.MaterialCode,
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

                                if (!CFMListPink.Contains(cfm1.MaterialCode))  //[
                                {
                                    CFMListPink.Add(cfm1.MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm1.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListLine.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListLine.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                                    try
                                    {
                                        if (!Coll[0].ColNo.Contains(Cols))
                                        {
                                            Coll[0].ColNo.Add(Cols);
                                        }
                                    }
                                    catch (Exception xx)
                                    {
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm1.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListLine.Add(Col);
                                    }
                                }

                                if (CFMListGreen.Contains(cfm1.MaterialCode))  //[+
                                {
                                    if (!CFMListGoldens.Contains(cfm1.MaterialCode))
                                    {
                                        CFMListGoldens.Add(cfm1.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm1.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListLinePlus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = CFMListLinePlus.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                                        try
                                        {
                                            if (!Coll[0].ColNo.Contains(Cols))
                                            {
                                                Coll[0].ColNo.Add(Cols);
                                            }
                                        }
                                        catch (Exception xx)
                                        {
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cfm1.MaterialCode,
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
                        }  //foreach
                    }//dontShow==false
                }  //остаток
                else if (cbViews.Text == "Потребление")
                {
                    var xlData = XLDataList.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();                    

                    /*  foreach (DateTime dt in DTList)
                      {
                          var xlD = xlData.Where(x => x.ProdDate == dt).ToList();

                          if (xlD.Count > 0)
                          {
                              double summ = xlD.Sum(x => x.Quantity);

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
                      }*/

                    foreach (DateTime dt in DTList)
                    {
                        var xlD = xlData.Where(x => x.DateWork.Date == dt).ToList();

                        if (xlD.Count > 0)
                        {
                            double summ = xlD.Sum(x => x.Quantity);

                            if ((Math.Abs(summ) <= 1) && (summ != 0))
                            {
                                sVal = sVal + summ.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + summ.ToString("N0") + RowSplitter;
                            }

                            /* if ((Math.Abs(xlD[0].Quantity) <= 1) && (xlD[0].Quantity != 0))
                             {
                                 sVal = sVal + xlD[0].Quantity.ToString("N1") + RowSplitter;
                             }
                             else
                             {
                                 sVal = sVal + xlD[0].Quantity.ToString("N0") + RowSplitter;
                             }*/

                        }
                        else
                        {
                            sVal = sVal + "" + RowSplitter;
                        }
                    }
                }
                else  //Дефицит
                {
                 //   Value = Value - cfm1.CurrentConsumption - cfm1.RawNav1;
                    var adata = fm.tdb.tApplication.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();     // AppList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();

                    var xlData = XLDataList.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                    var cListData = fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode)).ToList();
                    Cols = 16 + WorkMonthList.Count;  //14
                                                      //   foreach (DateTime dt in DTList)
                    for (int j = 0; j < DTList.Count; j++)
                    {
                        if (cfm1.DontShow == true)
                        {

                        }
                        else
                        {
                            DateTime dt = DTList[j];

                            var FDTList = DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();
                            //income
                            List<Raws> tData2 = new List<Raws>();
                            if (cbAddDay.Checked == false)
                            {
                                tData2 = cfm1.TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                            }
                            else
                            {
                                tData2 = cfm1.TaskNavList2.Where(x => x.AddDate == dt).ToList();
                            }

                            double T2 = 0;
                            if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                            {
                                tData2 = cfm1.TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                            }

                            //  foreach (var td2 in tData2)
                            //  {
                            //      T2 = T2 + td2.Quantity;
                            //  }
                            T2 = T2 + tData2.Sum(x => x.Quantity);

                            List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                            if (cbAddDay.Checked == false)
                            {
                                Cdata = cListData.Where(x => x.ExpDT.Date == dt.Date).ToList();    //fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                            }
                            else
                            {
                                Cdata = cListData.Where(x => x.ExpDT.Date.AddDays(1) == dt.Date).ToList();    //fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                            }

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                            {
                                Cdata = cListData.Where(x => ((DateTime)x.AlterDate).Date == dt.Date).ToList();     // fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                            }

                            //  foreach (var cd in Cdata)
                            //  {
                            //      T2 = T2 + cd.Quantity;
                            //  }
                            T2 = T2 + Cdata.Sum(x => x.Quantity);

                            if (T2 > 0)
                            {
                                if (Value < 0)
                                {
                                    Value = 0;
                                }

                                if (!AskList.Contains(cfm1.MaterialCode))
                                {
                                    AskList.Add(cfm1.MaterialCode);
                                }
                            }

                            Value = Value + T2;

                            //ask in progress
                            List<Raws> tData1 = new List<Raws>();

                            if (cbAddDay.Checked == false)
                            {
                                tData1 = cfm1.TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                            }
                            else
                            {
                                tData1 = cfm1.TaskNavList1.Where(x => x.AddDate == dt).ToList();
                            }
                            double T1 = 0;

                            if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                            {
                                tData1 = cfm1.TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                            }

                            //  foreach (var td1 in tData1)
                            //  {
                            //      T1 = T1 + td1.Quantity;
                            //  }
                            T1 = T1 + tData1.Sum(x => x.Quantity);

                            if (T1 > 0)
                            {
                                if (!AskList.Contains(cfm1.MaterialCode))
                                {
                                    AskList.Add(cfm1.MaterialCode);
                                }

                                /*if (Value < 0)
                                {
                                    Value = 0;
                                }*/
                            }

                           // Value = Value + T1;

                            //credit
                            var dll = xlData.Where(x => x.DateWork.Date == dt).ToList();
                            /*foreach (var dl in dll)
                            {
                                Value = Value - dl.Quantity;
                            }*/

                            Value = Value - dll.Sum(x => x.Quantity);

                            //future planning Asks
                            double Q = 0;
                            //   var adata = AppList;    // fm.tdb.tApplication.ToList();
                            //  if (adata.Count > 0)
                            //    adata = adata.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();
                            //    if (adata.Count > 0)
                            {
                                var bdata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                if ((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                {
                                    //  adata = AppList;
                                    //  adata = adata.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();
                                    bdata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                }

                                //    foreach (var ad in adata)
                                //    {
                                //        Q = Q + ad.Quantity;
                                //    }
                                Q = Q + bdata.Sum(x => x.Quantity);

                                USCList.Add(new UserSelectedCell
                                {
                                    RowIndex = i,
                                    ColIndex = Cols,
                                    Quantity = Q
                                });
                                //  Value = Value + Q;
                            }

                            if (Value < 0)
                            {
                                //  if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                {
                                    if (AskList.Contains(cfm1.MaterialCode))
                                    {
                                        for (int k = AskList.Count - 1; k >= 0; k--)
                                        {
                                            if (AskList[k] == cfm1.MaterialCode)
                                            {
                                                AskList.RemoveAt(k);
                                                k = 0;
                                            }
                                        }
                                    }

                                    //  CFMBool[i] = false;
                                    if (!CFMListGold.Contains(cfm1.MaterialCode))  //-
                                    {
                                        CFMListGold.Add(cfm1.MaterialCode);
                                        CFMListRed.Add(cfm1.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm1.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = CFMListMinus.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                                        try
                                        {
                                            if (!Coll[0].ColNo.Contains(Cols))
                                            {
                                                Coll[0].ColNo.Add(Cols);
                                            }
                                        }
                                        catch (Exception xx)
                                        {
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cfm1.MaterialCode,
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
                        }
                        sVal = sVal + "" + RowSplitter;
                        Cols = Cols + 1;
                    }
                }

                NewString = sVal.Split(RowSplitter);
                dt1.Rows.Add(NewString);  //    dataGridViewMain
                                          //  RowList.Add(NewString);
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
         //   dataGridViewMain.VirtualMode = true;

            for (int i = 8; i < 15; i++)
            {
                dataGridViewMain.Columns[i + WorkMonthList.Count].DefaultCellStyle.BackColor = Color.Gainsboro;
                dataGridViewMain.Columns[i + WorkMonthList.Count].HeaderCell.Style.Font = new Font(dataGridViewMain.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f, FontStyle.Bold);
            }

            dataGridViewMain.Columns[15 + WorkMonthList.Count].HeaderCell.Style.Font = new Font(dataGridViewMain.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f, FontStyle.Bold);

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

          /*  dataGridViewMain.Columns[0].SortMode = DataGridViewColumnSortMode.Programmatic;
            dataGridViewMain.Columns[1].SortMode = DataGridViewColumnSortMode.Programmatic;
            dataGridViewMain.Columns[2].SortMode = DataGridViewColumnSortMode.Programmatic;*/

            for (int i = 3; i < dataGridViewMain.Columns.Count; i++)
            {
                // c.SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridViewMain.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            DataLoaded = true;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Cursor = Cursors.Default;
            /*
            DTList = XLDataList.OrderBy(x => x.DateWork).Select(x => x.DateWork).ToList();
            DTList = DTList.Distinct().ToList();
            */
        }

        private void DGV_Get_CellColor()
        {
           /* Cursor = Cursors.WaitCursor;
            Int32 Offset = WorkMonthList.Count;
            //cells values
            Font f1 = new Font(dataGridViewMain.DefaultCellStyle.Font, FontStyle.Bold);
            //1-2 cells
            for (int j = 0; j < dataGridViewMain.Rows.Count; j++)  //     CFMList
            {
                // if ((j >= LowPageIndex) && (j <= HighPageIndex))
                // {
                string PC = dataGridViewMain.Rows[j].Cells[0].Value.ToString();   //dataGridViewMain.Rows[j].Cells[0].Value.ToString();   CFMList[j].MaterialCode

                if (URSCodeList.Contains(PC))                   //(dataGridViewMain.Rows[j].Cells["Код материала"].Value.ToString().ToLower() == URSList[i].MaterialCode.ToLower())
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = SColor;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = SColor;
                        }
                        catch (Exception xx)
                        { }
                    }
                    else if (URMCodeList.Contains(PC))
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = MColor;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = MColor;
                        }
                        catch (Exception xx)
                        { }
                    }
                    else if (fm.OPRList.Contains(PC))
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightCyan; //  .SteelBlue;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightCyan;
                        }
                        catch (Exception xx)
                        { }
                    }
                    else if (CFMListGold.Contains(PC))
                    {
                        if (CFMList[j].MaterialGroup != "ГП") //(CFMBool[j] == false) 
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightYellow;   //.Gold;
                                dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightYellow;   //.Gold;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                    else if (CFMListGray.Contains(PC))
                    {
                        if (CFMList[j].MaterialGroup != "ГП")
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGray;  //DarkGray
                                dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGray;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }

                    if (CFMListRed.Contains(PC)) //-
                    {
                        var CM = CFMListMinus.Where(x => x.ProdCode == PC).ToList();
                        if (CM.Count > 0)
                        {
                            foreach (var c in CM[0].ColNo)
                            {
                                try
                                {
                                    dataGridViewMain.Rows[j].Cells[c].Style.ForeColor = Color.Red; // Color.DarkRed;
                                 //   dataGridViewMain.Rows[j].Cells[c].Style.Font = f1;
                                }
                                catch (Exception xx)
                                { }
                            }
                        }
                    }

                    if (CFMListGreen.Contains(PC))  //+
                    {
                        var CP = CFMListPlus.Where(x => x.ProdCode == PC).ToList();
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

                    if (CFMListPink.Contains(PC))
                    {
                        var CL = CFMListLine.Where(x => x.ProdCode == PC).ToList();
                        if (CL.Count > 0)
                        {
                            foreach (var c in CL[0].ColNo)
                            {
                                try
                                {
                                    dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightPink;    //.LightPink;
                                }
                                catch (Exception xx)
                                { }
                            }
                        }
                    }

                    if (CFMListGoldens.Contains(PC))
                    {
                        var CLP = CFMListLinePlus.Where(x => x.ProdCode == PC).ToList();
                        if (CLP.Count > 0)
                        {
                            foreach (var c in CLP[0].ColNo)
                            {
                                try
                                {
                                    dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightYellow;   //.Gold;
                                }
                                catch (Exception xx)
                                { }
                            }
                        }
                    }

                    if (TaskList.Contains(PC))
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGreen;   //.Lime;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGreen;
                        }
                        catch (Exception xx)
                        { }
                    }
                    else if (MesList.Contains(PC))
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.MistyRose;   //.Rose;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.MistyRose;
                        }
                        catch (Exception xx)
                        { }
                    }


                    //9/10 color
                    try
                    {
                        var cfm = CFMList.Where(x => x.MaterialCode == PC).First();
                        if (cbMaterialDelay.Checked == false)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1);
                                dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2);
                            }
                            catch (Exception xx)
                            { }
                        }
                        else
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1UD);
                                dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2UD);
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                    catch (Exception xxx)
                    { }
               // } // j between LowPage and HighPage
            }
            //  DGV_UpdateColumnZero();
            Cursor = Cursors.Default;*/
        }

        private void DGV_UpdateColumnZero()
        {
            /* for (int j = 0; j < dataGridViewMain.RowCount; j++)
             {
                 string PC = dataGridViewMain.Rows[j].Cells[0].Value.ToString();


             }*/
        }

        private void LoadData()  //get the lists for filter data from MainForm
        {
            XLDataList = fm.ExcelDataList; //.Where(x => x.DateWork >= DateTime.Today.Date).ToList();
            XLDataList = XLDataList.Where(x => x.Quantity > 0).ToList();

            List<lData> DataListTempOld = new List<lData>();
            List<lData> DataListTempNew = new List<lData>();
            WorkMonthList = new List<DateTime>();
            lData LD;
            bool fl;
            List<string> OprComp = fm.OPRTableList.Select(x => x.MatCode).ToList();

            button1_a.BackColor = Color.LightGray; //    Color.DarkGray; #FFFF00 
            button1_b.BackColor = Color.LightPink;    // Color.LightPink;
            button1_c.BackColor = Color.LightYellow;      //Color.Gold;
            button1_d.BackColor = Color.LightCyan;     // Color.SteelBlue;
            button1_e.BackColor = Color.LightGreen;    //Color.Lime;
            button1_f.BackColor = Color.RosyBrown;
            button1_g.BackColor = Color.MistyRose;  

            comboBoxFiltering.Items.Clear();

            TBP = fm.tdb.tBlockedMaterial.Where(x => x.Using == false)
                    .Select(x => new ProductRecipes
                    {
                        ProductCode = x.ProductCode,
                        ProductName = x.ProductName
                    })
                    .ToList();

            foreach (var XLD in fm.ExcelDataList)  //.OrderBy(x => x.DateWork).ThenBy(x => x.Line).ToList()
            {
                comboBoxFiltering.Items.Add(XLD.DateWork.ToString("dd.MM.yyyy") + " " + XLD.Line + " " + XLD.ProdCodeStr + " " + XLD.AlterProdName + " " + XLD.Quantity.ToString("N0") + " кг");
            }

            Int32 Iteration = 0;
            MaterialNamesList = new List<string>();
            /*   MaterialData = new List<DailyResult>();
               FilteredMaterialData = new List<DailyResult>();*/
            DTList = new List<DateTime>();

            XLDataList = XLDataList.Where(x => x.Quantity > 0).ToList();
            ProductNamesList = XLDataList.Select(x => x.ProdCode).ToList();
            ProductNamesList = ProductNamesList.Distinct().ToList();
            DTList = XLDataList.OrderBy(x => x.DateWork).Select(x => x.DateWork).ToList();
            DTList = DTList.Distinct().ToList();

            try
            {
                StartDate = DTList.Min();
                EndDate = DTList.Max();
            }
            catch (Exception xx)
            {
                StartDate = DateTime.Today.Date;
                EndDate = StartDate.AddDays(1);
            }

            //Reciepts
            PRList = new List<ProductRecipes>();
            PRList = fm.ProdRecipeList.Where(x => ProductNamesList.Contains(x.ProductCode)).ToList();

         //   var PPR = PRList.Where(x => x.ProductCode == "1010023051").ToList();

            foreach (DateTime dt in DTList)
            {
                var DataList = XLDataList.Where(x => x.DateWork == dt).ToList();

                foreach (var DL in DataList)
                {
                    /*  if (DL.ProdCode == "1010024662")
                       {
                        if (DL.DateWork == Convert.ToDateTime("2022-10-02"))
                        { }
                    }*/

                    var Recipe = PRList.Where(x => x.ProductCode == DL.ProdCode).ToList();
                    if (Recipe.Count > 0)
                    {
                        var RecLine = Recipe[0].RecLineList.Where(x => x.LineNumber == DL.Line).ToList();
                        if (RecLine.Count > 0)
                        {
                            var RecLineDate = RecLine[0].RecList.Where(x => (x.WorkDateStart <= dt) && (x.WorkDateEnd >= dt)).ToList();
                            if (RecLineDate.Count > 0)
                            {
                                var RList = RecLineDate[0].rList;

                                foreach (var RL in RList)
                                {
                                    LD = new lData();
                                    LD.ProdCodeStr = "_____";
                                    if (DL.RePack == true)
                                    {
                                        if (RL.IPG.ToLower() == "тара")
                                        {
                                            LD.ProdCode = RL.MaterialCode;
                                            LD.ProdCodeStr = RL.MaterialCode;
                                            LD.DateWork = dt;
                                            LD.ProdName = RL.MaterialName;
                                            LD.Quantity = DL.Quantity * RL.MaterialQuantity;
                                            LD.Line = DL.Line;
                                            LD.IPG = "тара";
                                            LD.Spices = false;
                                        }
                                    }
                                    else
                                    {
                                        LD.ProdCode = RL.MaterialCode;
                                        LD.ProdCodeStr = RL.MaterialCode;
                                        LD.DateWork = dt;
                                        LD.ProdName = RL.MaterialName;
                                        LD.Quantity = DL.Quantity * RL.MaterialQuantity;
                                        LD.Line = DL.Line;
                                        LD.IPG = RL.IPG;
                                        LD.Spices = DL.Spices;   //control: split 1-st level only!!!!!!!!!
                                    }

                                    if (LD.ProdCodeStr != "_____")
                                    {
                                        fl = false;

                                        for (int i = 0; i < DataListTempNew.Count; i++)
                                        {
                                            if (DataListTempNew[i].ProdCode == LD.ProdCode)
                                            {
                                                if (DataListTempNew[i].Line == LD.Line)
                                                {
                                                    if (DataListTempNew[i].DateWork == LD.DateWork)
                                                    {
                                                        if (DataListTempNew[i].Spices == LD.Spices)
                                                        {
                                                            DataListTempNew[i].Quantity = DataListTempNew[i].Quantity + LD.Quantity;
                                                            fl = true;
                                                            i = DataListTempNew.Count;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (fl == false)
                                        {
                                            DataListTempNew.Add(LD);
                                        }
                                    }
                                }
                            }
                        }
                    }
                }
            }
            //end 1st step
            var d = DataListTempNew.Where(x => x.ProdCode == "3000106479").ToList();

            //10 и больше итераций полуфабрикатов - предложение выхода
            {
                Iteration = 0;
                XLDataList.Clear();

                while (DataListTempNew.Count > 0)
                {
                    Iteration = Iteration + 1;
                    for (int i = DataListTempNew.Count; i > 0; i--)
                    {
                        //   if (DataListTempNew[i-1].ProdCode == "100001456")
                        //   { }  

                       /* if (OprComp.Contains(DataListTempNew[i - 1].ProdCode))
                        {
                            var OPRRecList = fm.ProdRecipeOPRList.Where(x => x.ProductCode == DataListTempNew[i - 1].ProdCode).ToList();
                            if (OPRRecList.Count > 0)
                            {
                                var OPRRecLineList = OPRRecList[0].RecLineList;
                                if (OPRRecLineList.Count > 0)
                                {
                                    var OPRRecLineDateList = OPRRecLineList[0].RecList.Where(x => (x.WorkDateStart <= DataListTempNew[i - 1].DateWork) && (x.WorkDateEnd >= DataListTempNew[i - 1].DateWork)).ToList();
                                    if (OPRRecLineDateList.Count > 0)
                                    {
                                        var RList = OPRRecLineDateList[0].rList;
                                        foreach (var RR in RList)
                                        {
                                            LD = new lData
                                            {
                                                ProdCode = RR.MaterialCode,
                                                ProdCodeStr = RR.MaterialCode,
                                                DateWork = DataListTempNew[i - 1].DateWork,
                                                ProdName = RR.MaterialName,
                                                Quantity = RR.MaterialQuantity * DataListTempNew[i - 1].Quantity,
                                                Line = DataListTempNew[i - 1].Line,
                                                IPG = RR.IPG,
                                                Spices = false      //DataListTempNew[i - 1].Spices
                                            };

                                            fl = false;
                                            for (int j = 0; j < XLDataList.Count; j++)
                                            {
                                                if (XLDataList[j].ProdCode == LD.ProdCode)
                                                {
                                                    if (XLDataList[j].Line == LD.Line)
                                                    {
                                                        if (XLDataList[j].DateWork == LD.DateWork)
                                                        {
                                                            if (XLDataList[j].Spices == LD.Spices)
                                                            {
                                                                XLDataList[j].Quantity = XLDataList[j].Quantity + LD.Quantity;
                                                                fl = true;
                                                                j = XLDataList.Count;
                                                            }
                                                        }
                                                    }
                                                }
                                            }

                                            if (fl == false)
                                            {
                                                //  if (LD.Spices == false)
                                                //  {
                                                if ((LD.IPG != "ГП") && (LD.Quantity > 0))
                                                {
                                                    XLDataList.Add(LD);
                                                }
                                                //  }
                                                //  else
                                                //  {
                                                //      if (LD.IPG.ToLower() != "пф")
                                                //      {
                                                //          XLDataList.Add(LD);
                                                //      }
                                                //  }


                                                //      if ((LD.ProdCode == "1031008522") || (LD.ProdCodeStr == "1031008522"))
                                                //      { }
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        else*/   //all component add into List
                        {
                            fl = false;
                            for (int j = 0; j < XLDataList.Count; j++)
                            {
                                if (XLDataList[j].ProdCode == DataListTempNew[i - 1].ProdCode)
                                {
                                    if (XLDataList[j].DateWork == DataListTempNew[i - 1].DateWork)
                                    {
                                        if (XLDataList[j].Line == DataListTempNew[i - 1].Line)
                                        {
                                            if (XLDataList[j].Spices == DataListTempNew[i - 1].Spices)
                                            {
                                                fl = true;
                                                XLDataList[j].Quantity = XLDataList[j].Quantity + DataListTempNew[i - 1].Quantity;
                                                j = XLDataList.Count;
                                            }
                                        }
                                    }
                                }
                            }

                            if (fl == false)
                            {
                                LD = new lData
                                {
                                    ProdCode = DataListTempNew[i - 1].ProdCode,
                                    ProdName = DataListTempNew[i - 1].ProdName,
                                    DateWork = DataListTempNew[i - 1].DateWork,
                                    Quantity = DataListTempNew[i - 1].Quantity,
                                    IPG = DataListTempNew[i - 1].IPG,
                                    Line = DataListTempNew[i - 1].Line,
                                    Spices = DataListTempNew[i - 1].Spices
                                };

                                //   if ((LD.ProdCode == "1031008522") || (LD.ProdCodeStr == "1031008522"))
                                //   { }

                                /* if (LD.Spices == false)
                                 {*/
                                if (LD.Quantity > 0)         // ((LD.IPG != "ГП") && (LD.Quantity > 0))
                                {
                                    XLDataList.Add(LD);
                                }
                                /* }
                                 else
                                 {
                                     if (LD.IPG.ToLower() != "пф")
                                     {
                                         XLDataList.Add(LD);
                                     }
                                 }*/
                            }
                        }

                        /*  if (DataListTempNew[i - 1].IPG.ToLower() != "пф")
                          {*/
                        DataListTempNew.RemoveAt(i - 1);
                        /* }  
                         else
                         {
                             if (DataListTempNew[i - 1].Spices==false)
                             {
                                 DataListTempNew.RemoveAt(i - 1);
                             }
                         }*/
                    }

                    DataListTempOld.Clear();
                    DataListTempOld.AddRange(DataListTempNew);
                    DataListTempNew.Clear();


                    for (int i = 0; i < DataListTempOld.Count; i++)
                    {
                        var RecList = fm.ProdRecipeList.Where(x => x.ProductCode == DataListTempOld[i].ProdCode).ToList();

                        if (RecList.Count > 0)
                        {
                            var RecLineList = RecList[0].RecLineList.Where(x => x.LineNumber == DataListTempOld[i].Line).ToList();

                            if (RecLineList.Count > 0)
                            {
                                var RecLineDateList = RecLineList[0].RecList.Where(x => (x.WorkDateStart <= DataListTempOld[i].DateWork) && (x.WorkDateEnd >= DataListTempOld[i].DateWork)).ToList();

                                if (RecLineDateList.Count > 0)
                                {
                                    var RList = RecLineDateList[0].rList;
                                    foreach (var RR in RList)
                                    {
                                        LD = new lData
                                        {
                                            ProdCode = RR.MaterialCode,
                                            ProdCodeStr = RR.MaterialCode,
                                            DateWork = DataListTempOld[i].DateWork,
                                            ProdName = RR.MaterialName,
                                            Quantity = RR.MaterialQuantity * DataListTempOld[i].Quantity,
                                            Line = DataListTempOld[i].Line,
                                            IPG = RR.IPG
                                        };

                                        fl = false;
                                        for (int j = 0; j < DataListTempNew.Count; j++)
                                        {
                                            if (DataListTempNew[j].ProdCode == LD.ProdCode)
                                            {
                                                if (DataListTempNew[j].Line == LD.Line)
                                                {
                                                    if (DataListTempNew[j].DateWork == LD.DateWork)
                                                    {
                                                        DataListTempNew[j].Quantity = DataListTempNew[j].Quantity + LD.Quantity;
                                                        fl = true;
                                                        j = DataListTempNew.Count;
                                                    }
                                                }
                                            }
                                        }

                                        if (fl == false)
                                        {
                                            DataListTempNew.Add(LD);
                                            //  if ((LD.ProdCode == "1031008522") || (LD.ProdCodeStr == "1031008522"))
                                            //  { }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    if (Iteration > 10)
                    {
                        DialogResult dr = MessageBox.Show("Проведено " + Iteration.ToString() + " итераций. Полуфабрикаты не разложены на компоненты!" + System.Environment.NewLine + "Желаете прекратить обработку?", "Сообщение системы", MessageBoxButtons.OKCancel);
                        if (dr == DialogResult.OK)
                        {
                            DataListTempNew.Clear();
                        }
                    }
                }//while

            }
            //end iteration
            DataListTempNew.Clear();
            DataListTempOld.Clear();
            //    XLDataList = XLDataList.Where(x => x.IPG.Trim().ToLower() != "гп").ToList();
            MaterialNamesList = XLDataList.Select(x => x.ProdCode).ToList();
            MaterialNamesList = MaterialNamesList.Distinct().ToList();

            //   var XXData = XLDataList.Where(x => x.ProdCode == "1031007266").OrderBy(x=>x.Line).ToList();


            /*   MaterialData = XLDataList.Select(x => new DailyResult
               {
                   ProductCode = x.ProdCode,
                   LineName = x.ProdName,
                   Quantity = 1,
                   MaterialType = x.IPG.Trim()
               }).ToList();
               MaterialData = MaterialData.Distinct().ToList();*/

            //   List<string> MTypeList = MaterialData.Select(x => x.MaterialType).ToList();
            //   MTypeList = MTypeList.Distinct().ToList();

            //    FilteredMaterialData = MaterialData.Where(x => x.MaterialType.Trim() == "ГП").ToList();

            //Colors
            URSList = fm.tdb.tUserRecordSelected.ToList();
            URSCodeList = new List<string>();
            URSCodeList = URSList.Select(x => x.MaterialCode).Distinct().ToList();

            URMList = fm.tdb.tMercuryRecordSelected.ToList();
            URMCodeList = new List<string>();
            URMCodeList = URMList.Select(x => x.MaterialCode).Distinct().ToList();


            CSList = fm.tdb.tColorSelection.ToList();
            var SColorList = CSList.Where(x => x.Destination.ToLower() == "selecteduser").ToList();
            try
            {
                SColor = Color.FromName(SColorList[0].SelectedColor);
            }
            catch (Exception xx)
            { }

            SColorList = CSList.Where(x => x.Destination.ToLower() == "mercuryuser").ToList();
            try
            {
                MColor = Color.FromName(SColorList[0].SelectedColor);
            }
            catch (Exception xx)
            { }

            Int64 RCount = fm.tdb.tApplication.ToList().Count;

            if (RCount > 0)
            {
                button2.Enabled = true;
                button3.Enabled = true;
            }
            else
            {
                button2.Enabled = false;
                button3.Enabled = false;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void GetStorage()  //filtered data and calculate storage
        {
          /*  SBList = new List<StockBalances>();
            SBMList = new List<StockBalancesMes>();
            SBPList = new List<StockBalancesPlant>();
            SBRList = new List<StockBalancesResult>();

            DateTime LocalStartDate = DateTime.Today.Date.AddDays(-1);
            DateTime LocalEndDate = DateTime.Today.Date;
            DateTime CurrDay = LocalEndDate;

            List<StockBalancesResult> sbrln = new List<StockBalancesResult>();
            List<StockBalancesResult> sbrld = new List<StockBalancesResult>();
            List<StockBalancesResult> sbrlt = new List<StockBalancesResult>();
            List<StockBalancesMes> sbmln = new List<StockBalancesMes>();
            List<StockBalancesPlant> sbpln = new List<StockBalancesPlant>();*/
            //nav
        //    var sbDataList = fm.StockBaklanceList.Where(x => /*(x.TestQualityGr != "3Блок")  &&  (x.StorageType == "Собств") &&*/  (x.WorkDate.Date == LocalStartDate.Date) && (x.WorkDate.Hour >= 18) && (x.WorkDate.Hour <= 22)).ToList();
            //  sbDataList = sbDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
         //   sbDataList = sbDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();      //Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
           /* foreach (SelectStockBalances1_Result ss in sbDataList)
            {
                SBList.Add(fm.NewStockBalance(ss));
            }*/

           // sbDataList = fm.StockBaklanceList.Where(x => /*(x.TestQualityGr != "3Блок") && (x.StorageType == "Собств") && */ (x.WorkDate.Date == LocalEndDate.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            //    sbDataList = sbDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
           // sbDataList = sbDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();        //.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
           /* foreach (SelectStockBalances1_Result ss in sbDataList)
            {
                SBList.Add(fm.NewStockBalance(ss));
            }*/

         //   sbDataList.Clear();
         //   sbDataList = null;
            //mes
         //   var sbmDataList = fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == LocalStartDate.Date) && (x.WorkDate.Hour >= 18) && (x.WorkDate.Hour <= 22)).ToList();
         //   sbmDataList = sbmDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
        //    sbmDataList = sbmDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();       //Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
           /* foreach (SelectStockBalancesMes1_Result ssm in sbmDataList)
            {
                SBMList.Add(fm.NewStockBalancesMes(ssm));
            }*/

        //    sbmDataList = fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == LocalEndDate.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
         //   sbmDataList = sbmDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
          //  sbmDataList = sbmDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();        //Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
          /*  foreach (SelectStockBalancesMes1_Result ssm in sbmDataList)
            {
                SBMList.Add(fm.NewStockBalancesMes(ssm));
            }

            sbmDataList.Clear();
            sbmDataList = null;*/

          /*  var sbpDataList = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == LocalEndDate.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            sbpDataList = sbpDataList.Where(x => fm.StorageNamesList.Contains(x.LineName)).ToList();

            sbpDataList = sbpDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();         //.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
            foreach (var z in sbpDataList)
            {
                SBPList.Add(fm.NewStockBalancesPlant(z));
            }

            sbpDataList.Clear();
            sbpDataList = null;*/

            //суммируем остатки по лотам
            /*     var ssbList = SBMList.Where(d => (d.WorkDate.Date == LocalEndDate.Date) && (d.WorkDate.Hour >= 6) && (d.WorkDate.Hour <= 10)).ToList();
                 sbrld.Clear();
                 sbrld = ssbList.GroupBy(x => new { x.MaterialCode }).  // x.Storage, x.MaterialCode, x.MaterialLot
                     Select(g => new StockBalancesResult
                     {
                         WorkDate = LocalEndDate.Date,
                         //   Storage=g.Key.Storage.Trim(),
                         MaterialCode = g.Key.MaterialCode.Trim(),
                         //   LotName=g.Key.MaterialLot.Trim(),
                         QuantityMesDay = g.Sum(x => x.Quantity)
                     }).ToList();

                 ssbList = SBMList.Where(d => (d.WorkDate.Date == LocalStartDate.Date) && (d.WorkDate.Hour >= 18) && (d.WorkDate.Hour <= 22)).ToList();
                 sbrln.Clear();
                 sbrln = ssbList.GroupBy(x => new { x.MaterialCode}).  // x.Storage,  x.MaterialLot 
                     Select(g => new StockBalancesResult
                     {
                         WorkDate = LocalStartDate.Date,
                         //  Storage = g.Key.Storage.Trim(),
                         MaterialCode = g.Key.MaterialCode.Trim(),
                         //  LotName = g.Key.MaterialLot.Trim(),
                         QuantityMesNight = g.Sum(x => x.Quantity)
                     }).ToList();

                 foreach (var d in sbrld)
                 {
                     var m = new StockBalancesResult
                     {
                         WorkDate = LocalEndDate.Date,
                         //  Storage=d.Storage,
                         MaterialCode = d.MaterialCode,
                         QuantityMesDay = d.QuantityMesDay
                     };

                     SBRList.Add(m);
                 }

                 foreach (var d in sbrln)
                 {
                     var m = new StockBalancesResult
                     {
                         WorkDate = LocalEndDate.Date,
                         //  Storage = d.Storage,
                         MaterialCode = d.MaterialCode,
                         QuantityMesNight = d.QuantityMesNight
                     };

                     var s = SBRList.Where(x =>(x.MaterialCode == m.MaterialCode)).ToList();  // (x.Storage == m.Storage) && 

                     if (s.Count > 0)
                     {
                         foreach (var sl in s)
                         {
                             sl.QuantityMesNight = m.QuantityMesNight;
                             sl.QuantityMesDelta = sl.QuantityMesDay - sl.QuantityMesNight;
                         }
                     }
                     else
                     {
                         m.QuantityMesDelta = (-1) * m.QuantityMesNight;
                         SBRList.Add(m);
                     }
                 }

                 var sbDataListN = fm.StockBaklanceList.Where(x => (x.WorkDate.Date == LocalStartDate.Date) && (x.WorkDate.Hour >= 18) && (x.WorkDate.Hour <= 22)).ToList();  /// (x.TestQualityGr != "3Блок") && (x.StorageType == "Собств") &&
                 //   sbDataListN = sbDataListN.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
                 sbDataListN = sbDataListN.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
                 sbrlt.Clear();
                 sbrlt = sbDataListN.GroupBy(x => new {  x.MaterialCode}).  //x.Storage, , x.LotName
                     Select(g => new StockBalancesResult
                     {
                         WorkDate = LocalEndDate.Date,
                         //    Storage = g.Key.Storage.Trim(),
                         MaterialCode = g.Key.MaterialCode.Trim(),
                         QuantityStorageNight = (double)g.Sum(x => x.Quantity)
                     }).ToList();

                 foreach (var st in sbrlt)
                 {
                     var m = new StockBalancesResult
                     {
                         WorkDate = LocalEndDate.Date,
                         //   Storage=st.Storage,
                         MaterialCode = st.MaterialCode,
                         QuantityStorageNight = st.QuantityStorageNight
                     };

                     var s = SBRList.Where(x => (x.MaterialCode == m.MaterialCode)).ToList();  // (x.Storage == m.Storage) &&

                     if (s.Count > 0)
                     {
                         foreach (var sl in s)
                         {
                             sl.QuantityStorageNight = m.QuantityStorageNight;
                             sl.QuantityStorageDay = m.QuantityStorageNight - sl.QuantityMesDelta;
                         }
                     }
                     else
                     {
                         m.QuantityStorageDay = m.QuantityStorageNight;
                         SBRList.Add(m);
                     }
                 }

                 sbpln.Clear();
                 sbpln = SBPList.GroupBy(x => new { x.MaterialCode }).
                     Select(g => new StockBalancesPlant
                     {
                         WorkDate = LocalEndDate.Date,
                         MaterialCode = g.Key.MaterialCode.Trim(),
                         Quantity = g.Sum(x => x.Quantity)
                     }).ToList();


                 foreach (var sb in sbpln)
                 {
                     var m = new StockBalancesResult
                     {
                         WorkDate = LocalEndDate.Date,
                         //   Storage=st.Storage,
                         MaterialCode = sb.MaterialCode,
                         QuantityPlantDay = sb.Quantity
                     };

                     var s = SBRList.Where(x => (x.MaterialCode == m.MaterialCode)).ToList();// (x.Storage == m.Storage) &&
                     if (s.Count > 0)
                     {
                         foreach (var sl in s)
                         {
                             sl.QuantityPlantDay = m.QuantityPlantDay;
                             sl.QuantityDayItog = sl.QuantityStorageDay + sl.QuantityPlantDay;
                         }
                     }
                     else
                     {
                         m.QuantityDayItog = m.QuantityPlantDay;
                         SBRList.Add(m);
                     }
                 }*/

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        public void UpdateApplication(Int32 row, Int32 col, double Quant)
        {
            AppList = fm.tdb.tApplication.ToList();
            var ul = USCList.Where(x => (x.RowIndex == row) && (x.ColIndex == col)).ToList();
            double Q = Quant;

            if (ul.Count > 0)  //exist!
            {
                ul[0].Quantity = ul[0].Quantity + Quant;
                string s = dataGridViewMain.Rows[row].Cells[col].Value.ToString();
                if (s.IndexOf("{") > 0)
                {
                    s = s.Substring(0, s.IndexOf("{") - 1);
                }
                s = s + " {" + ul[0].Quantity.ToString("N0") + "}";
                dataGridViewMain.Rows[row].Cells[col].Value = s;
                Q = ul[0].Quantity;
            }
            else  //non exists
            {
                USCList.Add(new UserSelectedCell
                {
                    RowIndex = row,
                    ColIndex = col,
                    Quantity = Quant
                });

                dataGridViewMain.Rows[row].Cells[col].Value = dataGridViewMain.Rows[row].Cells[col].Value.ToString() + " {" + Quant.ToString("N0") + "}";
            }
            //dontShow==true

            string Code = dataGridViewMain.Rows[row].Cells[0].Value.ToString();
            var Rec = CFMList.Where(x => x.MaterialCode == Code).FirstOrDefault();

            if (Rec.DontShow == true)
            {
                Int32 C = col - 16 - WorkMonthList.Count;
                if (C >= 0)
                {
                    string V;
                    string vUD;
                    Rec.ValueS = "";
                    Rec.ValueUDS = "";
                    char RowSplitter = '|';

                    if ((Math.Abs(Rec.Value_) <= 1) && (Rec.Value_ != 0))
                    {
                        V = Rec.Value_.ToString("N1");
                    }
                    else
                    {
                        V = Rec.Value_.ToString("N0");
                    }

                    if ((Math.Abs(Rec.ValueUD) <= 1) && (Rec.ValueUD != 0))
                    {
                        vUD = Rec.ValueUD.ToString("N1");
                    }
                    else
                    {
                        vUD = Rec.ValueUD.ToString("N0");
                    }

                    //foreach (DateTime dt in DTList)
                    for (int i = 0; i < DTList.Count; i++)
                    {
                        Rec.ValueS = Rec.ValueS + vUD + ""; // + RowSplitter;
                        Rec.ValueUDS = Rec.ValueUDS + V + ""; // + RowSplitter;

                        if (i == C)
                        {
                            Rec.ValueS = Rec.ValueS + " {" + Q.ToString("N0") + "} ";
                            Rec.ValueUDS = Rec.ValueUDS + " {" + Q.ToString("N0") + "} ";
                        }

                        Rec.ValueS = Rec.ValueS + RowSplitter;
                        Rec.ValueUDS = Rec.ValueUDS + RowSplitter;
                    }
                }
            }
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

        private void button3_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormTaskClear"] == null)
            {
                Cursor = Cursors.WaitCursor;
                ftc = new FormTaskClear(this,fm);
                ftc.Visible = true;
                //   fs.MdiParent = this;
                ftc.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                ftc.BringToFront();
                ftc.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void ConsurmptionForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }

        private void ConsurmptionForm_Activated(object sender, EventArgs e)
        {
            fm.Enabled = false;
        }

        private void cbViews_SelectedIndexChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            panel4.Enabled = false;
            ClearDGV();
             RecalculateData();
            //ShowLocalData();
            //  DGV_UpdateColumnZero();

            /*LowPageIndex = memoryCache.GetLowIndex();
            HighPageIndex = memoryCache.GetHihgIndex();
            OldLowPageIndex = LowPageIndex;
            OldHighPageIndex = HighPageIndex;*/

            DGV_Get_CellColor();



            panel4.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<DailyResult> FilteredMaterialData1 = new List<DailyResult>();
            bool fl = false;
            string Code = "";
            CurrCode = "";

            if (tbMaterialCode.Text.Trim() != "")
            {
                FilteredMaterialData1 = MaterialData.Where(x => x.ProductCode.ToLower().Contains(tbMaterialCode.Text.Trim().ToLower())).ToList();
                //   FilteredMaterialData1 = FilteredMaterialData1.Where(x => x.ProductCode.ToLower().IndexOf(tbMaterialCode.Text.Trim().ToLower()) > 1).ToList();

                foreach (var f in FilteredMaterialData1)
                {
                    f.Location = -1 * f.ProductCode.ToLower().IndexOf(tbMaterialCode.Text.Trim().ToLower());
                }
                fl = true;
            }

            if (tbMaterialName.Text.Trim() != "")
            {
                if (fl == true)
                {
                    FilteredMaterialData1 = FilteredMaterialData1.Where(x => x.LineName.ToLower().Contains(tbMaterialName.Text.Trim().ToLower())).ToList();
                }
                else
                {
                    FilteredMaterialData1 = MaterialData.Where(x => x.LineName.ToLower().Contains(tbMaterialName.Text.Trim().ToLower())).ToList();
                }

                foreach (var f in FilteredMaterialData1)
                {
                    f.Location = f.LineName.ToLower().IndexOf(tbMaterialName.Text.Trim().ToLower());
                }
                fl = true;
            }

            if (comboBox1.Text != "")
            {
                if (fl == true)
                {
                    FilteredMaterialData1 = FilteredMaterialData1.Where(x => comboBox1.Text.ToLower().Contains(x.MaterialType.ToLower())).ToList();      //(x => x.MaterialType.ToLower()==comboBox1.Text.ToLower()).ToList();
                }
                else
                {
                    FilteredMaterialData1 = MaterialData.Where(x => comboBox1.Text.ToLower().Contains(x.MaterialType.ToLower())).ToList();
                }

            }

            /*  if (fl == true)
              {
                  FilteredMaterialData1 = FilteredMaterialData1.OrderBy(x => x.Location).ToList();
              }
              else
              {
                  FilteredMaterialData1 = FilteredMaterialData1.OrderBy(x => x.ProductCode).ToList();
              }*/

            FilteredMaterialDataIn.Clear();

            foreach (var fm1 in FilteredMaterialData1)
            {
                fl = true;

                foreach (var fm in FilteredMaterialDataIn)
                {
                    if (fm.ProductCode == fm1.ProductCode)
                    {
                        fl = false;
                    }
                }

                if (fl == true)
                {
                    DailyResult dr = new DailyResult
                    {
                        ProductCode = fm1.ProductCode,
                        LineName = fm1.LineName
                    };
                    FilteredMaterialDataIn.Add(dr);
                }
            }

            fl = false;
            //    List<string> Code= FilteredMaterialData
            //  FilteredMaterialData = FilteredMaterialData.Distinct().ToList();

            if (FilteredMaterialDataIn.Count > 0)
            {
                /* if (Application.OpenForms["FormFilterResult"] == null)
                 {
                     Cursor = Cursors.WaitCursor;
                     ffr = new FormFilterResult(this, fm);
                     ffr.Visible = true;
                     //   fs.MdiParent = this;
                     ffr.WindowState = FormWindowState.Normal;
                     Cursor = Cursors.Default;
                 }
                 else
                 {
                     ffr.BringToFront();
                     ffr.Visible = true;
                     //   cf.UpdateData();
                 }
                 this.Enabled = false;*/

                dataGridViewMain.ClearSelection();
                dataGridViewMain.CurrentCell = null;

                label5.Text = "Запись 1";
                label5.Visible = true;

                label6.Text = "из " + FilteredMaterialDataIn.Count.ToString();
                label6.Visible = true;
                //moveFirst
                Code = FilteredMaterialDataIn[0].ProductCode;
                CurrIndex = 0;
                Int32 RInd = 0;
                /*   for (int i = 0; i < CFMList.Count; i++)
                   {
                       if (CFMList[i].MaterialCode == Code)
                       {
                           RInd = i;
                           i = CFMList.Count + 1;
                       }
                   }*/

                for (int i = 0; i < dataGridViewMain.RowCount; i++)
                {
                    if (dataGridViewMain.Rows[i].Cells[0].Value.ToString() == Code)
                    {
                        RInd = i;
                        i = dataGridViewMain.RowCount + 1;
                    }
                }
                SelRowIndes = RInd;
                dataGridViewMain.FirstDisplayedScrollingRowIndex = RInd;
                dataGridViewMain.Rows[SelRowIndes].Selected = true;

                if (FilteredMaterialDataIn.Count > 1)
                {
                    button5.Visible = true;
                    button6.Visible = true;
                }
                else
                {
                    button5.Visible = false;
                    button6.Visible = false;
                }
                button9.Visible = true;

            //    ShowInfo(Code, SelRowIndes);
            }
            else
            {
                label5.Visible = false;
                label6.Visible = false;
                button5.Visible = false;
                button6.Visible = false;
                button9.Visible = false;

                dataGridViewMain.ClearSelection();
                dataGridViewMain.CurrentCell = null;

                MessageBox.Show("К сожалению, ничего не найдено", "Сообщение системы");
            }
            FilteredMaterialData1.Clear();
            Cursor = Cursors.Default;
        }

        private void dataGridViewMain_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor = Cursors.WaitCursor;         

             RowInd = e.RowIndex;
             ColInd = e.ColumnIndex;

            if ((RowInd != -1) /*&& (ColInd != -1)*/)
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
               // List<Raws> sListA = new List<Raws>();
              //  List<Raws> sListB = new List<Raws>();

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
                    dgvUsing.Rows.Clear();
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
                            fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;  //Cells["Класс материала"]
                            var E1 = fm.ExcelDataList.Where(x => DDList.Contains(x.DateWork) && x.Quantity>0).ToList();

                            for (int i = 0; i < DDList.Count; i++)     //  foreach (DateTime dt1 in DDList)   //DList
                            {
                                //planned
                                //  if (dt1 >= DateTime.Today.Date)
                                {
                                    var ELDListDate = E1.Where(x => (x.DateWork == DDList[i])).ToList();   //fm.ExcelDataList.Where(x => (x.DateWork == DDList[i]) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //dt1

                                     for (int j=0; j< ELDListDate.Count; j++)   // foreach (var ELD in ELDListDate)
                                     {
                                       /*if (dt1==Convert.ToDateTime("2022-10-02"))
                                        {
                                            if (ELD.ProdCode=="1010024662")
                                            { }
                                        }*/

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
                                //var LR = fm.lRec1.Where(x => x.MaterialCode == ProductCode).ToList();
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
                                        double Q = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[16 + WorkMonthList.Count].Value.ToString().Trim());

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

                        /* if (cbMaterialDelay.Checked == false)
                         {
                             sListA = cfm1[0].ListA.OrderBy(x => x.BBFDate).ToList();   //fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();
                             sListB =cfm1[0].ListD.OrderBy(x => x.BBFDate).ToList();          //fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim() != "")).OrderBy(x => x.PlanOperDate).ToList();
                         }
                         else
                         {
                             sListA = cfm1[0].ListA.Where(x => x.BBFDate >= DateTime.Today.Date).OrderBy(x => x.BBFDate).ToList();    // fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.PlanOperDate.Date >= DateTime.Today.Date) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();  //cfm1[0].TaskNavList1.Where(x => x.BBFDate >= DateTime.Today.Date).OrderBy(x => x.BBFDate).ToList();  
                             sListB = cfm1[0].ListD.Where(x => x.BBFDate >= DateTime.Today.Date).OrderBy(x => x.BBFDate).ToList();     ///fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.PlanOperDate.Date >= DateTime.Today.Date) && (x.StatusMZP.Trim() != "")).OrderBy(x => x.PlanOperDate).ToList();   //cfm1[0].TaskNavList2.Where(x => x.BBFDate >= DateTime.Today.Date).OrderBy(x => x.BBFDate).ToList();   
                         }*/

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

                        if (sListB.Count > 0)
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

                //  if (dataGridViewMain.Rows[e.RowIndex].Cells[2].Value.ToString().ToLower() == "гп")
                {
                    ToolStripMenuItem m1;
                    Int32 Offset = WorkMonthList.Count;
                    UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                    // contextMenuStripData.Items.Clear();
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

                        var ExData = XLDataList.Where(x => (x.ProdCode == UserSelectedCode) && (x.DateWork == CurrDate)).ToList(); //fm.ExcelList.Where(x => (x.ProductCode == UserSelectedCode) && (x.DateWork == dt)).ToList();

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
                       // var CM = CFMList.Where(x => x.MaterialCode == UserSelectedCode).ToList();

                        if (cfm1.Count > 0)
                        {
                            foreach (var cl in cfm1[0].ConsurmptionList)
                            {
                                m1 = new ToolStripMenuItem(cl.Storage + " => " + cl.Quantity.ToString("N1"));
                                contextMenuStripData.Items.Add(m1);
                            }
                        }
                    }
                    else if (e.ColumnIndex == (9 + Offset))   //остатки склады
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
                    else if (e.ColumnIndex == (10 + Offset))   //пр-во
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

        private void dataGridViewMain_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
        {
            Rectangle rect = dataGridViewMain.GetCellDisplayRectangle(e.ColumnIndex, e.RowIndex, true);
            AxX = rect.X +/* fm.Left +*/ this.Left;
            AxY = rect.Y +/* fm.Top + */this.Top + this.panel1.Height + 60;

            ColInd = e.ColumnIndex;
            RowInd = e.RowIndex;

            string LName = "";
            DateTime dt;

            if (e.Button == MouseButtons.Right)
            {
                if (e.RowIndex != -1)
                {
                    if ((e.ColumnIndex == 0) || (e.ColumnIndex == 1))//color selection
                    {
                        UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                        //Int32 X = Cursor.Position.X; 
                        //Int32 Y =  Cursor.Position.Y;                        
                        if ((dataGridViewMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == SColor) || (dataGridViewMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Style.BackColor == MColor))
                        {
                            contextMenuStripDrop.Show(AxX, AxY);
                        }
                        else
                        {
                            contextMenuStripSel.Show(AxX, AxY);
                        }
                    }
                    else //add request
                    {
                        //  if (dataGridViewMain.Rows[e.RowIndex].Cells["Класс материала"].Value.ToString().Trim().ToLower()=="гп")
                        {
                            if ((dataGridViewMain.Columns[e.ColumnIndex].HeaderText.Contains("Остаток на")) || (dataGridViewMain.Columns[e.ColumnIndex].HeaderText.Contains("Потребление на")) || ((dataGridViewMain.Columns[e.ColumnIndex].HeaderText.Contains("Дефицит на"))))
                            {
                                LName = dataGridViewMain.Columns[e.ColumnIndex].HeaderText;
                                LName = LName.Substring(LName.Length - 11, 11).Trim();
                                dt = fm.ConvertDataToDate(LName);

                                if (dt.Date >= DateTime.Today.Date)
                                {
                                    USC = new UserSelectedCell
                                    {
                                        UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString(),
                                        UserSelectedProdName = dataGridViewMain.Rows[e.RowIndex].Cells[1].Value.ToString(),
                                        UserSelectedDateTime = dt,
                                        RowIndex = e.RowIndex,
                                        ColIndex = e.ColumnIndex,
                                    };

                                    try
                                    {
                                        USC.Quantity = fm.ConvertStringToDouble(dataGridViewMain.Rows[e.RowIndex].Cells[e.ColumnIndex].Value.ToString().Trim());
                                    }
                                    catch (Exception xx)
                                    {
                                        USC.Quantity = 0;
                                    }

                                    //   UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                                    //   UserSelectedProdName = dataGridViewMain.Rows[e.RowIndex].Cells[1].Value.ToString();
                                    //   UserSelectedDateTime = dt;                         

                                    contextMenuStripRequest.Show(AxX, AxY);

                                    AxX = rect.X;  // AxX- fm.Left - this.Left;
                                    AxY = rect.Y; //  AxY - fm.Top - this.Top - this.panel1.Height - 80;
                                }
                            }
                        }
                    }
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
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

        private void dataGridViewMain_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
            //  DGV_Get_CellColor();
            //  DGV_UpdateColumnZero();
        }

        private void добавитьЗаявкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormRequest"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fr = new FormRequest(fm, this, USC);
                fr.Visible = true;
                //  fr.MdiParent = fm;
                fr.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fr.BringToFront();
                //   cf.UpdateData();
            }

            if (AxX > fr.Width / 2)
            {
                fr.Left = AxX - 25;
            }
            else
            {
                fr.Left = AxX;
            }
            if (AxY > fr.Height / 2)
            {
                fr.Top = AxY - 25;
            }
            else
            {
                fr.Top = AxY;
            }

            this.Enabled = false;
        }

        private void ToolStripMenuItemDrop_Click(object sender, EventArgs e)
        {
            if (UserSelectedCode != "")
            {
                var Delete = fm.tdb.Pr_DeleteUserRecord(UserSelectedCode);

                var Delete1 = fm.tdb.Pr_DeleteMercuryRecord(UserSelectedCode);

                for (int i = 0; i < CFMList.Count; i++)
                {
                    if (CFMList[i].MaterialCode == UserSelectedCode)
                    {
                        dataGridViewMain.Rows[i].Cells[0].Style.BackColor = Color.White;
                        dataGridViewMain.Rows[i].Cells[1].Style.BackColor = Color.White;
                    }
                }
            }

            UserSelectedCode = "";

            URSList = fm.tdb.tUserRecordSelected.ToList();
            URSCodeList = new List<string>();
            URSCodeList = URSList.Select(x => x.MaterialCode).Distinct().ToList();
            CSList = fm.tdb.tColorSelection.ToList();
            var SColorList = CSList.Where(x => x.Destination.ToLower() == "selecteduser").ToList();
            try
            {
                SColor = Color.FromName(SColorList[0].SelectedColor);
            }
            catch (Exception xx)
            { }

            URMList = fm.tdb.tMercuryRecordSelected.ToList();
            URMCodeList = new List<string>();
            URMCodeList = URMList.Select(x => x.MaterialCode).Distinct().ToList();
            SColorList = CSList.Where(x => x.Destination.ToLower() == "mercuryuser").ToList();
            try
            {
                MColor = Color.FromName(SColorList[0].SelectedColor);
            }
            catch (Exception xx)
            { }
            //   LoadData();
            //   ShowData();
            //   DGV_Get_CellColor();
            // DGV_UpdateColumnZero();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void ToolStripMenuItemSelect_Click(object sender, EventArgs e)
        {
            if (UserSelectedCode != "")
            {
                var Insert = fm.tdb.Pr_InsertUserRecord(UserSelectedCode, System.Environment.UserName);
            }
            UserSelectedCode = "";

            URSList = fm.tdb.tUserRecordSelected.ToList();
            URSCodeList = new List<string>();
            URSCodeList = URSList.Select(x => x.MaterialCode).Distinct().ToList();
            CSList = fm.tdb.tColorSelection.ToList();
            var SColorList = CSList.Where(x => x.Destination.ToLower() == "selecteduser").ToList();
            try
            {
                SColor = Color.FromName(SColorList[0].SelectedColor);
            }
            catch (Exception xx)
            { }
            //   LoadData();
            //   ShowData();
            //  DGV_Get_CellColor();
            //  DGV_UpdateColumnZero();

            for (int i = 0; i < CFMList.Count; i++)
            {
                if (URSCodeList.Contains(CFMList[i].MaterialCode))
                {
                    dataGridViewMain.Rows[i].Cells[0].Style.BackColor = SColor;
                    dataGridViewMain.Rows[i].Cells[1].Style.BackColor = SColor;
                }

            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            string s = "1000 {100}";
            s = s.Substring(0, s.IndexOf("{") - 1);
            MessageBox.Show(s);
        }

        private void ConsurmptionForm_Shown(object sender, EventArgs e)
        {
            ClearDGV();
            ShowData();

            DGV_Get_CellColor();
        }

        private void button5_Click(object sender, EventArgs e)
        {
            CurrIndex = CurrIndex - 1;
            if (CurrIndex < 0)
            {
                CurrIndex = FilteredMaterialDataIn.Count - 1;
            }
            CurrCode = "";

            dataGridViewMain.ClearSelection();
            dataGridViewMain.CurrentCell = null;        //.Rows[SelRowIndes].Selected = false;

            label5.Text = "Запись " + (CurrIndex + 1).ToString();
            string Code = FilteredMaterialDataIn[CurrIndex].ProductCode;
            Int32 RInd = 0;
         /*   for (int i = 0; i < CFMList.Count; i++)
            {
                if (CFMList[i].MaterialCode == Code)
                {
                    RInd = i;
                    i = CFMList.Count + 1;
                }
            }*/

            for (int i=0; i<dataGridViewMain.RowCount; i++)
            {
                if (dataGridViewMain.Rows[i].Cells[0].Value.ToString()==Code)
                {
                    RInd = i;
                    i = dataGridViewMain.RowCount + 1;
                }
            }

            ClearDGV();
            DGV_Get_CellColor();

            dataGridViewMain.FirstDisplayedScrollingRowIndex = RInd;
            SelRowIndes = RInd;
            dataGridViewMain.Rows[RInd].Selected = true;
         //  ShowInfo(Code, SelRowIndes);           
        }

        private void ShowInfo(string ProductCode, Int32 RowInd)
        {
            Cursor = Cursors.WaitCursor;

            List<StockBalancesResult> sbrMesList = new List<StockBalancesResult>();
            List<StockBalancesResult> sbrNavList = new List<StockBalancesResult>();
            List<StockBalancesResult> sbrItogList = new List<StockBalancesResult>();
            List<StockBalancesPlant> sbpList = new List<StockBalancesPlant>();
            List<NavPurchaseTable> sListA = new List<NavPurchaseTable>();
            List<NavPurchaseTable> sListB = new List<NavPurchaseTable>();

            dgvPlanned.Rows.Clear();
            dgvRecipes.Rows.Clear();
            dgvFactis.Rows.Clear();
            dgvStorages.Rows.Clear();
            dgvComment.Rows.Clear();

            string[] dataString;
            DateTime PlanDate;
            string State;
            bool fl;
            DateTime CurrDate = DateTime.Today.Date;
            List<DateTime> DDList = new List<DateTime>();
            double Quantity = 0;
            var cfm1 = CFMList.Where(x => x.MaterialCode == ProductCode).ToList();
      
            {
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
                                "Остатки производство:",
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

                //tasks

                if (cbMaterialDelay.Checked == false)
                {
                    sListA = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();
                    sListB = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim().ToLower() != "")).OrderBy(x => x.PlanOperDate).ToList();
                }
                else
                {
                    sListA = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.PlanOperDate.Date >= DateTime.Today.Date) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();
                    sListB = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.PlanOperDate.Date >= DateTime.Today.Date) && (x.StatusMZP.Trim().ToLower() != "")).OrderBy(x => x.PlanOperDate).ToList();
                }

                var sListBB = sListB.Where(x => /*(x.PlanOperDate.Date >= DateTime.Today.Date) &&*/ (x.StatusMZP.Trim().ToLower() == "заказано")).ToList();

                //containers
                var cList = fm.ContList.Where(x => (x.MaterialCode == ProductCode)).OrderBy(x => x.ExpDT).ToList();
             //   cList = cList.Where(x => x.ExpDT >= DateTime.Today.Date).ToList();

                foreach (var c in cList)
                {
                    NavPurchaseTable npt = new NavPurchaseTable
                    {
                        PlanOperDate = c.ExpDT,
                        PlanDate = c.ExpDT,
                        PlanQuantity = c.Quantity,
                        FactQuantity = c.Quantity,
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

                        dataString = new string[]
                         {
                                    s.OrderNymber,
                                    s.PlanOperDate.ToString("dd.MM.yyyy")+" "+s.Initiator+" "+s.StatusMZP,
                                    s.PlanQuantity.ToString("N1"),   //FactQuantity
                                    PlanDate.ToString("dd.MM.yyyy"),
                                    PlanDate.ToString("dd.MM.yyyy"),
                                    State
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

                        dataString = new string[]
                         {
                                    s.MZP.Trim()==""? s.OrderNymber:  s.OrderNymber+"/"+s.MZP,
                                    s.PlanOperDate.ToString("dd.MM.yyyy")+" "+s.Initiator+" "+s.StatusMZP,
                                    s.PlanQuantity.ToString("N1"),  //FactQuantity
                                    PlanDate.ToString("dd.MM.yyyy"),
                                    PlanDate.ToString("dd.MM.yyyy"),
                                    State
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

            //     //reciepts, planned, fact etc
            //   if (fl == true)
            {
                dataString = new string[] { "", "", "", "Данный компонент используется в:" };
                dgvPlanned.Rows.Add(dataString);
                dataString = new string[] { "", "", "", "", "", "Данный компонент используется в:" };
                dgvFactis.Rows.Add(dataString);

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

                var CurProduct = fm.lRec.Where(x => x.MaterialCode == ProductCode).ToList();    // fm.tdb.fn_select_CurrentProductView(ProductCode).ToList(); // fm.tdb.fn_select_CurrentProductS(ProductCode).ToList();      //fm.tdb.fn_select_CurrentProduct1(dt1, ProductCode).ToList();  
                                                                                                /*foreach (var cp in CurProduct)
                                                                                                {
                                                                                                    cp.LineNumber = fm.ReplaceLineName(cp.LineNumber);
                                                                                                }*/

                if (ColIndex < 8)
                {
                    DDList = DList;
                }
                else
                {
                    DDList = DList.Where(x => x.Date >= DateTime.Today.Date).OrderBy(x => x.Date).ToList();
                }

                foreach (DateTime dt1 in DDList)   //DList
                {
                    //planned
                    // if (dt1 >= DateTime.Today.Date)
                    {
                        fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;
                        var ELDListDate = fm.ExcelDataList.Where(x => (x.DateWork == dt1) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //  CurrDate

                        foreach (var ELD in ELDListDate)
                        {
                            /* var CP = CurProduct.Where(x => (x.ProductCode == ELD.ProdCode) && (x.LineNumber == ELD.Line)).ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
                             CP = CP.Where(x => x.WorkDate < dt1).OrderByDescending(x => x.WorkDate).ToList();  //Starting_Date

                             if (CP.Count > 0)*/

                            var CP = CurProduct.Where(x => (x.ProductCode == ELD.ProdCode) && (x.LineNumber == ELD.Line) && (x.WorkDate <= dt1)); //.ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
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
                                var PRL = fm.ProdRecipeList.Where(x => x.ProductCode == ELD.ProdCodeStr).ToList();
                                if (PRL.Count > 0)
                                {
                                    var LL = PRL[0].RecLineList.Where(x => x.LineNumber == ELD.Line).ToList();
                                    if (LL.Count > 0)
                                    {
                                        var DL = LL[0].RecList.Where(x => x.WorkDateStart <= dt1).OrderByDescending(x => x.WorkDateStart).ToList();  // CurrDate
                                        if (DL.Count > 0)
                                        {
                                            if (DL[0].WorkDateStart == SRR.WorkDate)
                                            {
                                                if (ELD.RePack == true)
                                                {
                                                    if (fl == true)
                                                    {
                                                        if ((double)SRR.Quantity > 0)
                                                        {
                                                            dataString = new string[] {
                                                                    dt1.ToString("dd.MM.yyyy"),  //  CurrDate
                                                                    ELD.Line,
                                                                    ELD.ProdCodeStr,
                                                                    ELD.AlterProdName,
                                                                    ELD.Quantity.ToString("N0"),
                                                                    ((double) SRR.Quantity).ToString("N5"),   //CP[0].Quantity
                                                                    (((double) SRR.Quantity) * ELD.Quantity).ToString("N1")
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
                                                                dt1.ToString("dd.MM.yyyy"),
                                                                ELD.Line,
                                                                ELD.ProdCodeStr,
                                                                ELD.AlterProdName,
                                                                ELD.Quantity.ToString("N0"),
                                                                ((double) SRR.Quantity).ToString("N5"),
                                                                (((double) SRR.Quantity) * ELD.Quantity).ToString("N1")
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

                        ELDListDate = null;
                    }  //dt1
                } //foreach
                //fact
                {
                    fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;
                    //  var MuaSection = fm.MUAList.Where(x => (x.DateTime >= DateStart) && (x.DateTime < DateEnd)).ToList();    

                    var FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, DateTime.Today.Date.AddYears(-2)).OrderBy(x => x.LotForErp).ToList();                      //fm.edb.fn_select_MaterialUsingByProduct(StartDate1, EndDate1, ProductCode).OrderBy(x => x.DateStart).ToList();

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
                                    f.MaterialCode,
                                    f.SystemName+" "+ f.MaterialName,
                                    ((double)f.Quantity).ToString("N2")
                            };

                            dgvFactis.Rows.Add(dataString);
                        }
                    }

                    FactData = null;
                }

                //recipes
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
                }
            }  //fl==true
            Cursor = Cursors.Default;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            CurrIndex = CurrIndex + 1;
            if (CurrIndex >= FilteredMaterialDataIn.Count)
            {
                CurrIndex = 0;
            }
            CurrCode = "";

            dataGridViewMain.ClearSelection();
            dataGridViewMain.CurrentCell = null;
            //  dataGridViewMain.Rows[SelRowIndes].Selected = false;

            label5.Text = "Запись " + (CurrIndex + 1).ToString();
            string Code = FilteredMaterialDataIn[CurrIndex].ProductCode;
            Int32 RInd = 0;
            /*   for (int i = 0; i < CFMList.Count; i++)
               {
                   if (CFMList[i].MaterialCode == Code)
                   {
                       RInd = i;
                       i = CFMList.Count + 1;
                   }
               }*/

            for (int i = 0; i < dataGridViewMain.RowCount; i++)
            {
                if (dataGridViewMain.Rows[i].Cells[0].Value.ToString() == Code)
                {
                    RInd = i;
                    i = dataGridViewMain.RowCount + 1;
                }
            }
            ClearDGV();
            DGV_Get_CellColor();

            dataGridViewMain.FirstDisplayedScrollingRowIndex = RInd;
            SelRowIndes = RInd;
            dataGridViewMain.Rows[RInd].Selected = true;

          //  ShowInfo(Code, SelRowIndes);   
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
                Int32 Cols = WorkMonthList.Count;
                Int32 CurCol = 16 + WorkMonthList.Count; //14
                DateTime WorkDate;
                List<DateTime> WorkDateList = new List<DateTime>();
                Int32 Interval;

                AppList = fm.tdb.tApplication.ToList();

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Остатки с МЗП";
                Excel1.Range range1;

                DateTime DateStart = DateTime.Today.Date;
                DateTime DateEnd = DateTime.Today.Date;

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

                    foreach (var el in CFMList[i].RawEnterpriseList)
                    {
                        Value = Value + el.Quantity;
                    }

                    foreach (var r1 in CFMList[i].RawNav1List)
                    {
                        Value = Value + r1.Quantity;
                    }

                    /*  foreach (var r2 in CFMList[i].RawNav2List)
                      {
                          Value = Value + r2.Quantity;
                      }*/

                    for (int j = 0; j < DTList.Count; j++)
                    {
                        T2 = 0;
                        var FDTList = DTList.Where(x => (x.Month == DTList[j].Month) && (x.Year == DTList[j].Year)).ToList();
                        //income
                        var tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == DTList[j].Date).ToList();
                        if ((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.AlterDate == DTList[j]).ToList();
                        }

                        foreach (var td2 in tData2)
                        {
                            T2 = T2 + td2.Quantity;
                        }

                        var Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == DTList[j].Date)).ToList();
                        if ((DTList[j].Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == DTList[j].Date)).ToList();
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
                        var tData1 = CFMList[i].TaskNavList1.Where(x => x.BBFDate == DTList[j].Date).ToList();
                        if ((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate == DTList[j]).ToList();
                        }

                        T1 = 0;
                        foreach (var td1 in tData1)
                        {
                            T1 = T1 + td1.Quantity;
                        }
                        //   Value = Value + T1;
                        //credit
                        var dll = XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value - dl.Quantity;
                        }

                        Q = 0;

                        var adata = AppList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();     // AppList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();
                        if ((adata.Count == 0) && (FDTList.Count == 1) && (DTList[j].Day == 1))
                        {
                            //   adata = AppList;
                            adata = AppList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();          //adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                            adata = adata.Where(x => (DateTime)x.AlterDate == DTList[j].Date).ToList();
                        }

                        foreach (var ad in adata)
                        {
                            Q = Q + ad.Quantity;
                        }

                        /*  if (Value < 0)
                          {
                              wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(Value, 0);
                              range1 = wSheet.Cells[i + 2, CurCol + j + 1] as Excel1.Range;
                              range1.Font.Bold = true;
                              range1.Font.Color = Color.Red;  //Color.DarkRed
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
                                foreach (var bd in adata)
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
                        // }
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

                    foreach (var el in CFMList[i].RawEnterpriseList)
                    {
                        Value = Value + el.Quantity;
                    }

                     foreach (var r1 in CFMList[i].RawNav1List)
                     {
                         Value = Value + r1.Quantity;
                     }

                   /*  foreach (var r2 in CFMList[i].RawNav2List)
                     {
                         Value = Value + r2.Quantity;
                     }*/

                    for (int j = 0; j < DTList.Count; j++)
                    {
                        //credit
                        var dll = XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();
                        foreach (var dl in dll)
                        {
                            Value = Value - dl.Quantity;
                        }

                        wSheet.Cells[i + 2, CurCol + j + 1] = Math.Round(Value, 0);

                        if (Value < 0)
                        {
                            range1 = wSheet.Cells[i + 2, CurCol + j + 1] as Excel1.Range;
                            range1.Font.Bold = true;
                            range1.Font.Color = Color.Red;   //Color.DarkRed;
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

                    foreach (var el in CFMList[i].RawEnterpriseList)
                    {
                        Value = Value + el.Quantity;
                    }

                    foreach (var r1 in CFMList[i].RawNav1List)
                    {
                        Value = Value + r1.Quantity;
                    }

                    /*  foreach (var r2 in CFMList[i].RawNav2List)
                      {
                          Value = Value + r2.Quantity;
                      }*/

                    for (int j = 0; j < DTList.Count; j++)
                    {
                        T2 = 0;
                        var FDTList = DTList.Where(x => (x.Month == DTList[j].Month) && (x.Year == DTList[j].Year)).ToList();
                        //income
                        var tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == DTList[j].Date).ToList();
                        if ((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = CFMList[i].TaskNavList2.Where(x => x.AlterDate == DTList[j]).ToList();
                        }

                        foreach (var td2 in tData2)
                        {
                            T2 = T2 + td2.Quantity;
                        }

                        var Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == DTList[j].Date)).ToList();
                        if ((DTList[j].Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == DTList[j].Date)).ToList();
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
                        var tData1 = CFMList[i].TaskNavList1.Where(x => x.BBFDate == DTList[j].Date).ToList();
                        if ((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate == DTList[j]).ToList();
                        }

                        T1 = 0;
                        foreach (var td1 in tData1)
                        {
                            T1 = T1 + td1.Quantity;
                        }
                      //  Value = Value + T1;

                        //credit
                        var dll = XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();
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

                    for (int j = 0; j < DTList.Count; j++)
                    {
                        Value = 0;
                        //credit
                        var dll = XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();
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

                //Среднемесячное потребление для ОПР. 1 месяц
                DateStart = WorkDateList[0];
                DateEnd = WorkDateList[1];

                var dataZ = fm.tdb.pr_GetOPZ(DateStart, DateEnd).ToList();
                for (int i = 0; i < Clist.Count; i++)
                {
                    var DZ = dataZ.Where(x => x.MaterialCode == Clist[i].MaterialCode).ToList();

                    if (DZ.Count > 0)
                    {
                        try
                        {
                            Clist[i].Month1 = -1 * (double)DZ[0].Quantity;
                        }
                        catch (Exception xx)
                        {
                            Clist[i].Month1 = 0;
                        }
                    }
                }

                //2 month
                DateStart = WorkDateList[1];
                DateEnd = WorkDateList[2];
                dataZ = fm.tdb.pr_GetOPZ(DateStart, DateEnd).ToList();

                for (int i = 0; i < Clist.Count; i++)
                {
                    var DZ = dataZ.Where(x => x.MaterialCode == Clist[i].MaterialCode).ToList();
                    if (DZ.Count > 0)
                    {
                        try
                        {
                            Clist[i].Month2 = -1 * (double)DZ[0].Quantity;
                        }
                        catch (Exception xx)
                        {
                            Clist[i].Month2 = 0;
                        }
                    }
                }

                //3 month
                DateStart = WorkDateList[2];
                DateEnd = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
                dataZ = fm.tdb.pr_GetOPZ(DateStart, DateEnd).ToList();

                for (int i = 0; i < Clist.Count; i++)
                {
                    var DZ = dataZ.Where(x => x.MaterialCode == Clist[i].MaterialCode).ToList();

                    if (DZ.Count > 0)
                    {
                        try
                        {
                            Clist[i].Month3 = -1 * (double)DZ[0].Quantity;
                        }
                        catch (Exception xx)
                        {
                            Clist[i].Month3 = 0;
                        }
                    }
                }

                for (int i = 0; i < Clist.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = Clist[i].MaterialCode;
                    wSheet.Cells[i + 2, 2] = Clist[i].MaterialName;
                    wSheet.Cells[i + 2, 3] = Clist[i].WaitingDays;
                    wSheet.Cells[i + 2, 4] = Clist[i].StorageQuantity;

                    Value = Clist[i].Month1 + Clist[i].Month2 + Clist[i].Month3;

                    wSheet.Cells[i + 2, 5] = Clist[i].Month1;
                    wSheet.Cells[i + 2, 6] = Clist[i].Month2;
                    wSheet.Cells[i + 2, 7] = Clist[i].Month3;

                    /* var OList = Clist[i].OutList.Where(x => x.ProdDate == WorkDateList[0]).ToList();
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
                     }*/

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

                    if (T2 > 0)
                    {
                        Value = (T2 * 30 / Value) - Interval;
                        wSheet.Cells[i + 2, 13] = Math.Round(Value, 0);
                    }
                    else
                    { }
                }//OPR

                /*   var OPRListNames = fm.OPRList.Except(Clist.Select(x => x.MaterialCode).ToList()).ToList();
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
                       Value = 0;


                       var dData = fm.StockBaklanceList.Where(x => x.WorkDate.Date == DateTime.Today.Date && x.MaterialCode == OprList[i].MatCode).ToList();
                       Value = dData.Sum(x => (double)x.Quantity);  //storage
                       wSheet.Cells[i + Offer, 9] = Value;

                       var dData1 = fm.StockBalancesPlantList.Where(x => x.WorkDate.Date == DateTime.Today.Date && x.MaterialCode == OprList[i].MatCode).ToList();
                       Value = Value + dData1.Sum(x => x.Quantity);
                       wSheet.Cells[i + Offer, 10] = dData1.Sum(x => x.Quantity);  //Plant

                       var dData2 = fm.StockBalancesMesList.Where(x => x.WorkDate.Date == DateTime.Today.Date && x.MaterialCode == OprList[i].MatCode).ToList();
                       Value = Value + dData2.Sum(x => x.Quantity);
                       wSheet.Cells[i + Offer, 11] = dData2.Sum(x => x.Quantity);  //mes

                       wSheet.Cells[i + Offer, 12] = Value;
                   }*/

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

        private void cbUseMaterialPlanning_CheckedChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            panel4.Enabled = false;
            ClearDGV();
             RecalculateData();
            //ShowLocalData();
            //  DGV_UpdateColumnZero();
            /*LowPageIndex = memoryCache.GetLowIndex();
            HighPageIndex = memoryCache.GetHihgIndex();
            OldLowPageIndex = LowPageIndex;
            OldHighPageIndex = HighPageIndex;*/

            DGV_Get_CellColor();


            panel4.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void button8_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Int32 Offset = WorkMonthList.Count;

            var ConsList = fm.ConsDataList;    // fm.edb.fn_select_FactConsurmption().ToList();
            for (int i = 0; i < CFMList.Count; i++)
            {
                var CL = ConsList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                CFMList[i].ConsurmptionList.Clear();
                CFMList[i].CurrentConsumption = 0;

                if (CL.Count > 0)
                {
                    CFMList[i].CurrentConsumption = CL.Sum(x => (double)x.Quantity);
                }

                foreach (var cl in CL)
                {
                    Raws r = new Raws
                    {
                        Storage = cl.JobName,
                        Quantity = (double)cl.Quantity
                    };

                    CFMList[i].ConsurmptionList.Add(r);
                }

                /*  dataGridViewMain.Rows[i].Cells[8 + Offset].Value = CFMList[i].CurrentConsumption.ToString("N0");*/
            }

            ClearDGV();
            ShowData();
         //   DGV_Get_CellColor();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Cursor = Cursors.Default;
        }

        private void comboBoxFiltering_SelectedIndexChanged(object sender, EventArgs e)
        {
            if (cbPlanFilter.Checked == true)
            {
              //  ClearDGV();
                List<string> OprComp = fm.OPRTableList.Select(x => x.MatCode).ToList();

                Cursor = Cursors.WaitCursor;
                if (comboBoxFiltering.Text.Trim() != "")
                {
                    var XLD = fm.ExcelDataList[comboBoxFiltering.SelectedIndex];
                    //contol spices!!!!!
                    var tdata = fm.tdb.tComponentException.Where(x => x.IsActive == true).ToList();
                    List<string> ExceptNames = tdata.Select(x => x.ProductCode).ToList();
                    ExceptNames = ExceptNames.Distinct().ToList();

                    foreach (string en in ExceptNames)
                    {
                        if (XLD.ProdCode == en)
                        {
                            XLD.Spices = true;
                        }
                    }

                    List<DailyResult> FilteredMaterialData1 = new List<DailyResult>();
                    string Code = XLD.ProdCode;
                    DateTime DateWork = XLD.DateWork;
                    bool RePack = XLD.RePack;
                    bool fl;
                    string Line = XLD.Line;
                    lData LD;
                    List<lData> DataListTempOld = new List<lData>();
                    List<lData> DataListTempNew = new List<lData>();
                    List<lData> LocalList = new List<lData>();
                    Int32 Iteration;

                    //get Reciept by Date, Code, Line
                    var Recipe = fm.ProdRecipeList.Where(x => x.ProductCode == XLD.ProdCode).ToList();
                    if (Recipe.Count > 0)
                    {
                        var RecLine = Recipe[0].RecLineList.Where(x => x.LineNumber == XLD.Line).ToList();
                        if (RecLine.Count > 0)
                        {
                            var RecLineDate = RecLine[0].RecList.Where(x => (x.WorkDateStart <= DateWork) && (x.WorkDateEnd >= DateWork)).ToList();
                            if (RecLineDate.Count > 0)
                            {
                                var RList = RecLineDate[0].rList;

                                foreach (var RL in RList)
                                {
                                    LD = new lData();
                                    LD.ProdCodeStr = "_____";
                                    if (XLD.RePack == true)
                                    {
                                        if (RL.IPG.ToLower() == "тара")
                                        {
                                            LD.ProdCode = RL.MaterialCode;
                                            LD.ProdCodeStr = RL.MaterialCode;
                                            LD.DateWork = XLD.DateWork;
                                            LD.ProdName = RL.MaterialName;
                                            LD.Quantity = XLD.Quantity * RL.MaterialQuantity;
                                            LD.Line = XLD.Line;
                                            LD.IPG = "тара";
                                            LD.Spices = false;
                                        }
                                    }
                                    else
                                    {
                                        LD.ProdCode = RL.MaterialCode;
                                        LD.ProdCodeStr = RL.MaterialCode;
                                        LD.DateWork = XLD.DateWork;
                                        LD.ProdName = RL.MaterialName;
                                        LD.Quantity = XLD.Quantity * RL.MaterialQuantity;
                                        LD.Line = XLD.Line;
                                        LD.IPG = RL.IPG;
                                        LD.Spices = XLD.Spices;
                                    }

                                    if (LD.ProdCodeStr != "_____")
                                    {
                                        fl = false;

                                        for (int i = 0; i < DataListTempNew.Count; i++)
                                        {
                                            if (DataListTempNew[i].ProdCode == LD.ProdCode)
                                            {
                                                if (DataListTempNew[i].Line == LD.Line)
                                                {
                                                    if (DataListTempNew[i].DateWork == LD.DateWork)
                                                    {
                                                        if (DataListTempNew[i].Spices == LD.Spices)
                                                        {
                                                            DataListTempNew[i].Quantity = DataListTempNew[i].Quantity + LD.Quantity;
                                                            fl = true;
                                                            i = DataListTempNew.Count;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                        if (fl == false)
                                        {
                                            DataListTempNew.Add(LD);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //end 1st step

                    //10 и больше итераций полуфабрикатов - предложение выхода
                    {
                        Iteration = 0;
                        LocalList.Clear();

                        while (DataListTempNew.Count > 0)
                        {
                            Iteration = Iteration + 1;
                            for (int i = DataListTempNew.Count; i > 0; i--)
                            {
                                /* if (DataListTempNew[i-1].ProdCode == "3000105460")
                                 { }*/

                              /*  if (OprComp.Contains(DataListTempNew[i - 1].ProdCode))
                                {
                                    var OPRRecList = fm.ProdRecipeOPRList.Where(x => x.ProductCode == DataListTempNew[i - 1].ProdCode).ToList();
                                    if (OPRRecList.Count > 0)
                                    {
                                        var OPRRecLineList = OPRRecList[0].RecLineList;
                                        if (OPRRecLineList.Count > 0)
                                        {
                                            var OPRRecLineDateList = OPRRecLineList[0].RecList.Where(x => (x.WorkDateStart <= DataListTempNew[i - 1].DateWork) && (x.WorkDateEnd >= DataListTempNew[i - 1].DateWork)).ToList();
                                            if (OPRRecLineDateList.Count > 0)
                                            {
                                                var RList = OPRRecLineDateList[0].rList;
                                                foreach (var RR in RList)
                                                {
                                                    LD = new lData
                                                    {
                                                        ProdCode = RR.MaterialCode,
                                                        ProdCodeStr = RR.MaterialCode,
                                                        DateWork = DataListTempNew[i - 1].DateWork,
                                                        ProdName = RR.MaterialName,
                                                        Quantity = RR.MaterialQuantity * DataListTempNew[i - 1].Quantity,
                                                        Line = DataListTempNew[i - 1].Line,
                                                        IPG = RR.IPG,
                                                        Spices = false  // DataListTempNew[i - 1].Spices
                                                    };

                                                    fl = false;
                                                    for (int j = 0; j < LocalList.Count; j++)
                                                    {
                                                        if (LocalList[j].ProdCode == LD.ProdCode)
                                                        {
                                                            if (LocalList[j].Line == LD.Line)
                                                            {
                                                                if (LocalList[j].DateWork == LD.DateWork)
                                                                {
                                                                    if (LocalList[j].Spices == LD.Spices)
                                                                    {
                                                                        LocalList[j].Quantity = LocalList[j].Quantity + LD.Quantity;
                                                                        fl = true;
                                                                        j = LocalList.Count;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                    }

                                                    if (fl == false)
                                                    {
                                                        // if (LD.Spices == false)
                                                        // {
                                                        LocalList.Add(LD);
                                                        //  }
                                                        //  else
                                                        //  {
                                                        //      if (LD.IPG.ToLower() != "пф")
                                                        //      {
                                                        //          LocalList.Add(LD);
                                                        //      }
                                                        //  }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else*/
                                {
                                    fl = false;
                                    for (int j = 0; j < LocalList.Count; j++)
                                    {
                                        if (LocalList[j].ProdCode == DataListTempNew[i - 1].ProdCode)
                                        {
                                            if (LocalList[j].DateWork == DataListTempNew[i - 1].DateWork)
                                            {
                                                if (LocalList[j].Line == DataListTempNew[i - 1].Line)
                                                {
                                                    if (LocalList[j].Spices == DataListTempNew[i - 1].Spices)
                                                    {
                                                        fl = true;
                                                        LocalList[j].Quantity = LocalList[j].Quantity + DataListTempNew[i - 1].Quantity;
                                                        j = LocalList.Count;
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (fl == false)
                                    {
                                        LD = new lData
                                        {
                                            ProdCode = DataListTempNew[i - 1].ProdCode,
                                            ProdName = DataListTempNew[i - 1].ProdName,
                                            DateWork = DataListTempNew[i - 1].DateWork,
                                            Quantity = DataListTempNew[i - 1].Quantity,
                                            IPG = DataListTempNew[i - 1].IPG,
                                            Line = DataListTempNew[i - 1].Line,
                                            Spices = DataListTempNew[i - 1].Spices
                                        };

                                        //  if (LD.Spices == false)
                                        //  {
                                        //   if (LD.Quantity > 0)   // ((LD.IPG != "ГП") && (LD.Quantity > 0))
                                        {
                                            LocalList.Add(LD);
                                        }
                                        /*  }
                                          else
                                          {
                                              if (LD.IPG.ToLower() != "пф")
                                              {
                                                  LocalList.Add(LD);
                                              }
                                          }*/
                                    }
                                }

                                // if (DataListTempNew[i - 1].IPG.ToLower() != "пф")
                                // {
                                DataListTempNew.RemoveAt(i - 1);
                                /* }
                                 else
                                 {
                                     if (DataListTempNew[i - 1].Spices == false)
                                     {
                                         DataListTempNew.RemoveAt(i - 1);
                                     }
                                 }*/
                            }

                            DataListTempOld.Clear();
                            DataListTempOld.AddRange(DataListTempNew);
                            DataListTempNew.Clear();

                            for (int i = 0; i < DataListTempOld.Count; i++)
                            {
                                var RecList = fm.ProdRecipeList.Where(x => x.ProductCode == DataListTempOld[i].ProdCode).ToList();

                                if (RecList.Count > 0)
                                {
                                    var RecLineList = RecList[0].RecLineList.Where(x => x.LineNumber == DataListTempOld[i].Line).ToList();

                                    if (RecLineList.Count > 0)
                                    {
                                        var RecLineDateList = RecLineList[0].RecList.Where(x => (x.WorkDateStart <= DataListTempOld[i].DateWork) && (x.WorkDateEnd >= DataListTempOld[i].DateWork)).ToList();

                                        if (RecLineDateList.Count > 0)
                                        {
                                            var RList = RecLineDateList[0].rList;
                                            foreach (var RR in RList)
                                            {
                                                LD = new lData
                                                {
                                                    ProdCode = RR.MaterialCode,
                                                    ProdCodeStr = RR.MaterialCode,
                                                    DateWork = DataListTempOld[i].DateWork,
                                                    ProdName = RR.MaterialName,
                                                    Quantity = RR.MaterialQuantity * DataListTempOld[i].Quantity,
                                                    Line = DataListTempOld[i].Line,
                                                    IPG = RR.IPG
                                                };

                                                fl = false;
                                                for (int j = 0; j < DataListTempNew.Count; j++)
                                                {
                                                    if (DataListTempNew[j].ProdCode == LD.ProdCode)
                                                    {
                                                        if (DataListTempNew[j].Line == LD.Line)
                                                        {
                                                            if (DataListTempNew[j].DateWork == LD.DateWork)
                                                            {
                                                                DataListTempNew[j].Quantity = DataListTempNew[j].Quantity + LD.Quantity;
                                                                fl = true;
                                                                j = DataListTempNew.Count;
                                                            }
                                                        }
                                                    }
                                                }

                                                if (fl == false)
                                                {
                                                    DataListTempNew.Add(LD);
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (Iteration > 10)
                            {
                                DialogResult dr = MessageBox.Show("Проведено " + Iteration.ToString() + " итераций. Полуфабрикаты не разложены на компоненты!" + System.Environment.NewLine + "Желаете прекратить обработку?", "Сообщение системы", MessageBoxButtons.OKCancel);
                                if (dr == DialogResult.OK)
                                {
                                    DataListTempNew.Clear();
                                }
                            }
                        }//while
                    }
                    //end iteration
                    DataListTempNew.Clear();
                    DataListTempOld.Clear();

                    LocalList = LocalList.OrderBy(x => x.ProdName).ToList();

                    /* TBP = fm.tdb.tBlockedMaterial.Where(x => x.Using == false)
                         .Select(x => new ProductRecipes
                         {
                             ProductCode = x.ProductCode,
                             ProductName = x.ProductName
                         })
                         .ToList();*/

                    for (int i = LocalList.Count - 1; i >= 0; i--)
                    {
                        foreach (var t in TBP)
                        {
                            if (LocalList.Count > 0)
                            {
                                if (LocalList[i].ProdCode == t.ProductCode)
                                {
                                    LocalList.RemoveAt(i);
                                }
                            }
                        }
                    }

                    //  FilteredMaterialDataOut.Clear();

                    foreach (var ll in LocalList)
                    {
                        DailyResult dl = new DailyResult
                        {
                            LineName = ll.ProdName,
                            ProductCode = ll.ProdCode,
                            Location = 0
                        };

                        FilteredMaterialData1.Add(dl);
                    }

                    FilteredMaterialDataOut.Clear();

                    foreach (var fm1 in FilteredMaterialData1)
                    {
                        fl = true;

                        foreach (var fm in FilteredMaterialDataOut)
                        {
                            if (fm.ProductCode == fm1.ProductCode)
                            {
                                fl = false;
                            }
                        }

                        if (fl == true)
                        {
                            DailyResult dr = new DailyResult
                            {
                                ProductCode = fm1.ProductCode,
                                LineName = fm1.LineName
                            };
                            FilteredMaterialDataOut.Add(dr);
                        }
                    }

                    fl = false;


                    if (FilteredMaterialDataOut.Count > 0)
                    {
                        if (Application.OpenForms["FormFilterResult"] == null)
                        {
                            Cursor = Cursors.WaitCursor;
                            ffr = new FormFilterResult(this, fm, XLD.DateWork);
                            ffr.Visible = true;
                            //   fs.MdiParent = this;
                            ffr.WindowState = FormWindowState.Normal;
                            Cursor = Cursors.Default;
                        }
                        else
                        {
                            ffr.BringToFront();
                            ffr.Visible = true;
                            //   cf.UpdateData();
                        }
                        this.Enabled = false;

                        /* dataGridViewMain.ClearSelection();
                         dataGridViewMain.CurrentCell = null;

                         label5.Text = "Запись 1";
                         label5.Visible = true;

                         label6.Text = "из " + FilteredMaterialData.Count.ToString();
                         label6.Visible = true;
                         //moveFirst
                         Code = FilteredMaterialData[0].ProductCode;
                         CurrIndex = 0;
                         Int32 RInd = 0;
                         for (int i = 0; i < CFMList.Count; i++)
                         {
                             if (CFMList[i].MaterialCode == Code)
                             {
                                 RInd = i;
                                 i = CFMList.Count + 1;
                             }
                         }
                         SelRowIndes = RInd;
                         dataGridViewMain.FirstDisplayedScrollingRowIndex = RInd;
                         dataGridViewMain.Rows[SelRowIndes].Selected = true;

                         button5.Visible = true;
                         button6.Visible = true;

                         ShowInfo(Code, SelRowIndes);*/
                    }
                    else
                    {
                        label5.Visible = false;
                        label6.Visible = false;
                        button5.Visible = false;
                        button6.Visible = false;
                    }
                    FilteredMaterialData1.Clear();
                }
                Cursor = Cursors.Default;
            }
        }

        private void cbMaterialDelay_CheckedChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            panel4.Enabled = false;
            ClearDGV();
             RecalculateData();
          //  ShowLocalData();
            //  DGV_UpdateColumnZero();

            /*LowPageIndex = memoryCache.GetLowIndex();
            HighPageIndex = memoryCache.GetHihgIndex();
            OldLowPageIndex = LowPageIndex;
            OldHighPageIndex = HighPageIndex;*/

            DGV_Get_CellColor();

            panel4.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void ClearDGV()
        {
            dataGridViewMain.ClearSelection();
            dataGridViewMain.CurrentCell = null;            

            dgvFactis.Rows.Clear();
            dgvPlanned.Rows.Clear();
            dgvRecipes.Rows.Clear();
            dgvUsing.Rows.Clear();
            dgvStorages.Rows.Clear();
            dgvComment.Rows.Clear();
        }

        private void cbPlanFilter_CheckedChanged(object sender, EventArgs e)
        {

        }

        private void cbAddDay_CheckedChanged(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            panel4.Enabled = false;
            ClearDGV();
             RecalculateData();
            //ShowLocalData();
            //  DGV_UpdateColumnZero();
            DGV_Get_CellColor();

            LowPageIndex = memoryCache.GetLowIndex();
            HighPageIndex = memoryCache.GetHihgIndex();
            OldLowPageIndex = LowPageIndex;
            OldHighPageIndex = HighPageIndex;

            panel4.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void cbDontShowAfter_CheckedChanged(object sender, EventArgs e)
        {
            /*  Cursor = Cursors.WaitCursor;
              ClearDGV();

              RecalculateData();

              //  DGV_UpdateColumnZero();
              DGV_Get_CellColor();
              Cursor = Cursors.Default;*/
        }

        private void ClearButtons()
        {
            dataGridViewMain.ClearSelection();
            dataGridViewMain.CurrentCell = null;

            label5.Visible = false;
            label6.Visible = false;
            button5.Visible = false;
            button6.Visible = false;

            try
            {
                FilteredMaterialDataIn.Clear();
            }
            catch (Exception xx)
            { }

            tbMaterialCode.Text = "";
            tbMaterialName.Text = "";
            comboBox1.Text = "";
            button9.Visible = false;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            ClearButtons();
        }

        private void выделитьМеркурийToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (UserSelectedCode != "")
            {
                var Insert = fm.tdb.Pr_InsertMercuryRecord(UserSelectedCode, System.Environment.UserName);
            }
            UserSelectedCode = "";

            URMList = fm.tdb.tMercuryRecordSelected.ToList();
            URMCodeList = new List<string>();
            URMCodeList = URMList.Select(x => x.MaterialCode).Distinct().ToList();
            CSList = fm.tdb.tColorSelection.ToList();
            var SColorList = CSList.Where(x => x.Destination.ToLower() == "mercuryuser").ToList();
            try
            {
                MColor = Color.FromName(SColorList[0].SelectedColor);
            }
            catch (Exception xx)
            { }
            //   LoadData();
            //   ShowData();
            //  DGV_Get_CellColor();
            //  DGV_UpdateColumnZero();

            for (int i = 0; i < CFMList.Count; i++)
            {
                if (URMCodeList.Contains(CFMList[i].MaterialCode))
                {
                    dataGridViewMain.Rows[i].Cells[0].Style.BackColor = MColor;
                    dataGridViewMain.Rows[i].Cells[1].Style.BackColor = MColor;
                }

            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void dataGridViewMain_Sorted(object sender, EventArgs e)
        {
        
        }

        private void dataGridViewMain_Sorted_1(object sender, EventArgs e)
        {
            DGV_Get_CellColor();
        }

        private void dataGridViewMain_DataBindingComplete(object sender, DataGridViewBindingCompleteEventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            Int32 Offset = WorkMonthList.Count;
            //Font f1 = new Font(dataGridViewMain.DefaultCellStyle.Font, FontStyle.Bold);

            for (int j = 0; j < dataGridViewMain.Rows.Count; j++)  //     CFMList
            {
                // if ((j >= LowPageIndex) && (j <= HighPageIndex))
                // {
                string PC = dataGridViewMain.Rows[j].Cells[0].Value.ToString();   //dataGridViewMain.Rows[j].Cells[0].Value.ToString();   CFMList[j].MaterialCode

                if (URSCodeList.Contains(PC))                   //(dataGridViewMain.Rows[j].Cells["Код материала"].Value.ToString().ToLower() == URSList[i].MaterialCode.ToLower())
                {
                    try
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = SColor;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = SColor;
                    }
                    catch (Exception xx)
                    { }
                }
                else if (URMCodeList.Contains(PC))
                {
                    try
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = MColor;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = MColor;
                    }
                    catch (Exception xx)
                    { }
                }
                else if (fm.OPRList.Contains(PC))
                {
                    try
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightCyan; //  .SteelBlue;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightCyan;
                    }
                    catch (Exception xx)
                    { }
                }
                else if (CFMListGold.Contains(PC))
                {
                    if (CFMList[j].MaterialGroup != "ГП") //(CFMBool[j] == false) 
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightYellow;   //.Gold;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightYellow;   //.Gold;
                        }
                        catch (Exception xx)
                        { }
                    }
                }
                else if (CFMListGray.Contains(PC))
                {
                    if (CFMList[j].MaterialGroup != "ГП")
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGray;  //DarkGray
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGray;
                        }
                        catch (Exception xx)
                        { }
                    }
                }

                if (CFMListRed.Contains(PC)) //-
                {
                    var CM = CFMListMinus.Where(x => x.ProdCode == PC).ToList();
                    if (CM.Count > 0)
                    {
                        foreach (var c in CM[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.ForeColor = Color.Red; // Color.DarkRed;
                                                                                               //   dataGridViewMain.Rows[j].Cells[c].Style.Font = f1;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                if (CFMListGreen.Contains(PC))  //+
                {
                    var CP = CFMListPlus.Where(x => x.ProdCode == PC).ToList();
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

                if (CFMListPink.Contains(PC))
                {
                    var CL = CFMListLine.Where(x => x.ProdCode == PC).ToList();
                    if (CL.Count > 0)
                    {
                        foreach (var c in CL[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightPink;    //.LightPink;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                if (CFMListGoldens.Contains(PC))
                {
                    var CLP = CFMListLinePlus.Where(x => x.ProdCode == PC).ToList();
                    if (CLP.Count > 0)
                    {
                        foreach (var c in CLP[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightYellow;   //.Gold;
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                if (TaskList.Contains(PC))
                {
                    try
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGreen;   //.Lime;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGreen;
                    }
                    catch (Exception xx)
                    { }
                }
                else if (MesList.Contains(PC))
                {
                    try
                    {
                        dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.MistyRose;   //.Rose;
                        dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.MistyRose;
                    }
                    catch (Exception xx)
                    { }
                }


                //9/10 color
                try
                {
                    var cfm = CFMList.Where(x => x.MaterialCode == PC).First();
                    if (cbMaterialDelay.Checked == false)
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1);
                            dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2);
                        }
                        catch (Exception xx)
                        { }
                    }
                    else
                    {
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1UD);
                            dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2UD);
                        }
                        catch (Exception xx)
                        { }
                    }
                }
                catch (Exception xxx)
                { }
                // } // j between LowPage and HighPage
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
                  string  LName = dataGridViewMain.Columns[ColInd].HeaderText;
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
                        var E1 = fm.ExcelDataList.Where(x => DDList.Contains(x.DateWork) && x.Quantity > 0).ToList();

                        for (int i=0; i<DDList.Count; i++)     // foreach (DateTime dt1 in DDList)   //DList
                        {
                            //planned
                            //  if (dt1 >= DateTime.Today.Date)
                            {
                                fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;  //Cells["Класс материала"]
                                var ELDListDate = E1.Where(x => (x.DateWork == DDList[i])).ToList();    // fm.ExcelDataList.Where(x => (x.DateWork == DDList[i]) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //dt1

                                 for (int j=0; j< ELDListDate.Count; j++)  //  foreach (var ELD in ELDListDate)
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
                    if (dgvUsing.RowCount<=1)
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

                            var LR2 = CurProduct.Where(x => x.LineNumber == s && x.WorkDate>=StartDate1).ToList();
                            LResult.AddRange(LR2);
                        }

                        LResult = LResult.OrderBy(x => x.WorkDate).ThenBy(x => x.LineNumber).ToList();
                        dgvUsing.Rows.Clear();

                        foreach (var lr in LResult)
                        {
                            if (lr.Quantity>0)
                            {
                                double Q = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[16 + WorkMonthList.Count].Value.ToString().Trim());

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
            GC.Collect();
            Cursor = Cursors.Default;
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormCommdet"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fc = new FormComment( fm, this);
                fc.Visible = true;
                //   fs.MdiParent = this;
                fc.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fc.BringToFront();
                fc.Visible = true;
                //   cf.UpdateData();
            }
        }
    }
}
