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
    //    FormLineList fll;
        FormRequest fr;
        FormFilterResult ffr;
        FormTaskClear ftc;
        FormComment fc;
        FormTransfer ft;
        FormReport frt;
        UpdateWindow UW;
        FormTask ftt;

        public DateTime StartDate;
        public DateTime EndDate;
        List<DateTime> DList;

        public BindingSource bs;
        public DataTable dt1;

        public List<tStorageTask> StorageTaskList;
        public List<tApplication> AppList;
        public List<tUserRecordSelected> URSList;
        public List<tMercuryRecordSelected> URMList;
        public List<string> URSCodeList;
        public List<string> URMCodeList;
        public List<tColorSelection> CSList;
        public List<string> ProductNamesList;
        public List<string> MaterialNamesList;
        public List<lData> XLDataList;
        public List<lData> XLDataListFilter;
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
        public List<string> CFMListTan;
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

        public List<UserSelectedCell> USCTList;
        public List<Colorres> CFMListPlum;
        public List<string> CFMListPlumes;
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
        bool LoadFlag;

        public ConsurmptionForm(FormMain FM)   //, FormLineList FLL
        {
            InitializeComponent();
            fm = FM;

        //    fll = FLL;
            label6.Visible = false;
            label5.Visible = false;
            button6.Visible = false;
            button5.Visible = false;
            button9.Visible = false;   

            CurrIndex = 0;
            SelRowIndes = 0;
            USCList = new List<UserSelectedCell>();
            DayNight = true;         //FLL.checkBox1.Checked;

            LoadData();  //workMonthList!
            GetStorage();

            dtpFact1.Value = DateTime.Today.Date.AddMonths(-1);
            dtpFact2.Value = DateTime.Today.Date;
            dtpPlan1.Value = DateTime.Today.Date.AddDays(-100);
          

            //   fll.Close();
            fm.Enabled = false;

            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;
            cbViews.Text = cbViews.Items[0].ToString();
            dgvStorages.Columns[2].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;
            dgvFactis.Columns[0].DefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleRight;

            comboBox1.Items.Clear();

            var TH = fm.lRec.Select(x => x.IPG).ToList();            //fm.tdb.tHeadersNames.OrderBy(x=>x.HName).ToList();
            TH.AddRange(fm.lRec1.Select(x => x.IPG).ToList());
            TH = TH.Distinct().ToList();
            TH.Add("Н/Д");
            TH.Sort();

            for (int i = 0; i < TH.Count; i++)
            {
                comboBox1.Items.Add(TH[i]);
            }

            PrepareData();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
 
            dateTimePicker1.Value = DateTime.Today.Date;

        //    RecalculateData();
            ClearDGV();

            //  Thread thread = new Thread(GetData);
            //  thread.Start();
        }

        private void PrepareData()
        {
            string sVal = "";
            string sValUD = "";
            char RowSplitter = '|';

            CFMListGray = new List<string>();
            USCTList = new List<UserSelectedCell>();

            DateTime XlDataStart = fm.ExcelDataList.Min(x => x.DateWork);
            XlDataStart = new DateTime(XlDataStart.Year, XlDataStart.Month, 1);
            DateTime XlDataEnd = fm.ExcelDataList.Max(x => x.DateWork);
            XlDataEnd = new DateTime(XlDataEnd.Year, XlDataEnd.Month, 1);
            XlDataEnd = XlDataEnd.AddMonths(1).AddDays(-1);

            DList = new List<DateTime>();
            DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();
            DList = DList.Distinct().OrderBy(x => x.Date).ToList();

            dtpPlan2.Value = XlDataEnd;
            WorkMonthList = new List<DateTime>();

            //карта
            DateTime DT1;
            var ConsList = fm.ConsDataList;  //   fm.edb.fn_select_FactConsurmption().ToList();
          //  CFMList = fm.mList;  //.Where(x => MaterialNamesList.Contains(x.MaterialCode)).OrderBy(x => x.MaterialCode).ToList();
                                 //   CFMList = CFMList.Where(x => x.MaterialCode == "8СИ00127").ToList();

            CFMList = new List<Material>();

            //New material
            var NList = fm.NewRecList.GroupBy(x => new { x.MaterialCode, x.MaterialName, x.IPG, x.Bom })
                .Select(g => new MaterialBaseData
                {
                    MaterialCode = g.Key.MaterialCode,
                    MaterialName = g.Key.MaterialName,
                    IPG = g.Key.IPG,
                    BOM = g.Key.Bom
                }).ToList();

            TaskList = new List<string>();
            TaskList = NList.Select(x => x.MaterialCode).ToList();
            TaskList = TaskList.Distinct().ToList();

            foreach (string s in TaskList)
            {
                var NL = NList.Where(x => x.MaterialCode == s).ToList();

                if (NL.Count >0)
                {
                    Material c = new Material();

                    c.MaterialCode = NL[0].MaterialCode;
                    c.MaterialName = NL[0].MaterialName;
                    c.BOM = NL[0].BOM;
                    c.MaterialGroup = NL[0].IPG;
                    c.MaterialStatus = -1; 

                    CFMList.Add(c);
                }
            }

            //Вывод
            var DDList = XLDataList.GroupBy(x => new { x.ProdCode, x.ProdName, x.IPG, x.BOM })
                .Select(g => new MaterialBaseData
                {
                    MaterialCode = g.Key.ProdCode.Trim(),
                    MaterialName = g.Key.ProdName.Trim(),
                    IPG = g.Key.IPG,
                    BOM = g.Key.BOM
                }).ToList();

            DDList = DDList.Distinct().ToList();

            TaskList = new List<string>();
            TaskList = DDList.Select(x => x.MaterialCode).ToList();
            TaskList = TaskList.Distinct().ToList();

            for (int i = TaskList.Count - 1; i >= 0; i--)
            {
                var cfm = CFMList.Where(x => x.MaterialCode == TaskList[i]).ToList();

                if (cfm.Count > 0)
                {
                    TaskList.RemoveAt(i);
                }
            }

         /*   for (int i =  TaskList.Count - 1; i>=0; i-- )
            {
                var cfm = CFMList.Where(x => x.MaterialCode == TaskList[i]).ToList();

                if (cfm.Count > 0)
                {
                    TaskList.RemoveAt(i);
                }
            }*/

            foreach (string s in TaskList)
            {
                var DD = DDList.Where(x => x.MaterialCode == s).ToList();

                if (DD.Count >0)
                {
                    Material c = new Material();

                    c.MaterialCode = DD[0].MaterialCode;
                    c.MaterialName = DD[0].MaterialName;
                    c.BOM = DD[0].BOM;
                    c.MaterialGroup = DD[0].IPG;

                    if (c.BOM.ToLower().Contains("вывод"))
                    {
                        c.MaterialStatus = 0;
                    }
                    else
                    {
                        c.MaterialStatus = 9;
                    }

                    CFMList.Add(c);
                }
            }

            /* for (int i = 0; i<DDList.Count; i++)
             {
                 Material c = new Material();

                 c.MaterialCode = DDList[i].MaterialCode;
                 c.MaterialName = DDList[i].MaterialName;
                 c.BOM = DDList[i].BOM;
                 c.MaterialGroup = DDList[i].IPG;

                 if (c.BOM.ToLower().Contains("вывод"))
                 {
                     c.MaterialStatus = 0;
                 }
                 else
                 {
                     c.MaterialStatus = 9;
                 }

                 CFMList.Add(c);
             }*/

            //заявки
            TaskList = new List<string>();
            var SL = fm.StorageList.ToList();     //.Where(x => (x.DateWork.Date == DateTime.Today.Date) && (x.DateWork.Hour >= 6) && (x.DateWork.Hour <= 10)).ToList();
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

                    var MList =   CFMList.Where(x => x.MaterialCode == s).ToList();         //fm.mList.Where(x => x.MaterialCode == s).ToList();

                    /*  if (s == "1031006114")
                      { }*/

                    if (MList.Count >0)
                    {
                        if (MList[0].MaterialStatus > 6)
                        {
                            MList[0].MaterialStatus = 6;
                        }
                    }
                    else
                    {
                        Material M = new Material();

                        M.MaterialCode = s;
                        M.MaterialName = sList[0].MaterialName;

                        var SM = fm.SMList.Where(x => x.MaterialCode == s).ToList();

                        if (SM.Count > 0)
                        {
                            if (SM[0].BOM.ToLower().Contains("вывод"))   
                            {
                                M.MaterialStatus = 0;
                            }
                        }

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
                        M.DontShow = false;

                        if (M.MaterialStatus > 6)
                        {
                            M.MaterialStatus = 6;
                        }


                        M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                        M.OldTaskNavUD = 0;

                        CFMList.Add(M);
                    }                 

                   /* if (MList.Count > 0)
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
                        M.DontShow = false;

                        if (M.MaterialStatus > 6)
                        {
                            M.MaterialStatus = 6;
                        }
                    }
                    else
                    {

                    }*/
                }
            }

            //остатки Mes
            var mml = fm.StockBalancesMesList.ToList();           //.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
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
                /*if (s == "1030100658")
                { }*/

                var mesList = mml.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (mesList.Count > 0)
                {
                    var MList = CFMList.Where(x => x.MaterialCode == s).ToList();           //fm.mList.Where(x => x.MaterialCode == s).ToList();

                    if (MList.Count >0)
                    {
                        if (MList[0].MaterialStatus > 7)
                        {
                            MList[0].MaterialStatus = 7;
                        }
                    }
                    else
                    {
                        Material M = new Material();

                        M.MaterialCode = s;
                        M.MaterialName = mesList[0].MaterialName;

                        var SM = fm.SMList.Where(x => x.MaterialCode == s).ToList();

                        if (SM.Count > 0)
                        {
                            if (SM[0].BOM.ToLower().Contains("вывод"))
                            {
                                M.MaterialStatus = 0;
                            }
                        }

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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }

                        M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                        M.OldTaskNavUD = 0;

                        CFMList.Add(M);
                    }

                 /*   if (MList.Count > 0)
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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }
                    }
                    else
                    {

                    }*/
                }
            }

            //остатки НАВ
            var sbl = fm.StockBaklanceList.ToList();          //.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.Quantity > 0)).ToList();

            MesList = sbl.Select(x => x.MaterialCode).Distinct().ToList();

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
                /*if (s == "1030100658")
                { }*/

                var mesList = sbl.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (mesList.Count > 0)
                {
                    var MList = CFMList.Where(x => x.MaterialCode == s).ToList();         //fm.mList.Where(x => x.MaterialCode == s).ToList();

                    if (MList.Count >0)
                    {
                        if (MList[0].MaterialStatus > 7)
                        {
                            MList[0].MaterialStatus = 7;
                        }
                    }
                    else
                    {
                        Material M = new Material();

                        M.MaterialCode = s;
                        M.MaterialName = mesList[0].MaterialName;

                        var SM = fm.SMList.Where(x => x.MaterialCode == s).ToList();

                        if (SM.Count > 0)
                        {
                            if (SM[0].BOM.ToLower().Contains("вывод"))
                            {
                                M.MaterialStatus = 0;
                            }
                        }

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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }

                        M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                        M.OldTaskNavUD = 0;

                        CFMList.Add(M);
                    }

                   /* if (MList.Count > 0)
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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }
                    }
                    else
                    {

                    }*/
                }
            }

            var sbp = fm.StockBalancesPlantList.ToList();         //.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.Quantity > 0)).ToList();  //остатки пр-во

            // plant  4, 6, 8
            var sbpNonZero = sbp.ToList();  /*Where(x => ((x.MaterialGroupID == 4) || (x.MaterialGroupID == 5) || (x.MaterialGroupID == 8)) ).*/

            List<string> PlanDataList = sbpNonZero.Select(x => x.MaterialCode).Distinct().ToList();

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
                /*if (s == "1030100658")
                { }*/

                var pdList = sbpNonZero.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (pdList.Count > 0)
                {
                    var MList = CFMList.Where(x => x.MaterialCode == s).ToList();       // fm.mList.Where(x => x.MaterialCode == s).ToList();

                    if (MList.Count >0)
                    {
                        if (MList[0].MaterialStatus > 7)
                        {
                            MList[0].MaterialStatus = 7;
                        }
                    }
                    else
                    {
                        Material M = new Material();

                        M.MaterialCode = s;
                        M.MaterialName = pdList[0].MaterialName;

                        var SM = fm.SMList.Where(x => x.MaterialCode == s).ToList();

                        if (SM.Count > 0)
                        {
                            if (SM[0].BOM.ToLower().Contains("вывод"))
                            {
                                M.MaterialStatus = 0;
                            }
                        }

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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }

                        M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                        M.OldTaskNavUD = 0;

                        CFMList.Add(M);
                        MesList.Add(s);
                    }

                  /*  if (MList.Count > 0)
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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }
                    }
                    else
                    {

                    }*/
                }
            }

            //plANT 7 (335, 336, 338, 346, 349, 375, 400)
            sbpNonZero.Clear();  // = sbp.Where(x => (x.MaterialGroupID == 7) && 
                                 //((x.MaterialTypeID == 335) || (x.MaterialTypeID == 336) || (x.MaterialTypeID == 338) || (x.MaterialTypeID == 346) || (x.MaterialTypeID == 359) || (x.MaterialTypeID == 375) || (x.MaterialTypeID == 400))).ToList();

            //   PlanDataList = sbpNonZero.Select(x => x.MaterialCode).Distinct().ToList();

            /*   for (int i = PlanDataList.Count - 1; i >= 0; i--)
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
               }*/

            //9 plant ( 175, 389)  
            //   sbpNonZero = sbp.Where(x => (x.MaterialGroupID == 9) && 
            //      ((x.MaterialTypeID == 175) || (x.MaterialTypeID == 389) )).ToList();
            //not like N'Отход%' and m.Name not like N'Репроцесс%'

            var NonZero = fm.tdb.pr_GetAlarmWaste().Where(x => x.WorkDate.Date == DateTime.Today.Date).ToList();  // && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)

            DT1 = NonZero.Max(x => x.WorkDate);
            DT1 = DT1.AddMinutes(-20);

            NonZero = NonZero.Where(x => x.WorkDate >= DT1).ToList();

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
                /*if (s == "1030100658")
                { }*/

                var pdList = NonZero.Where(x => x.MaterialCode == s).ToList();         //fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == s)).ToList();

                if (pdList.Count > 0)
                {
                    var MList = CFMList.Where(x => x.MaterialCode == s).ToList();      //fm.mList.Where(x => x.MaterialCode == s).ToList();

                    if (MList.Count >0)
                    {
                        if (MList[0].MaterialStatus > 7)
                        {
                            MList[0].MaterialStatus = 7;
                        }
                    }
                    else
                    {
                        Material M = new Material();

                        M.MaterialCode = s;
                        M.MaterialName = pdList[0].MaterialName;

                        var SM = fm.SMList.Where(x => x.MaterialCode == s).ToList();

                        if (SM.Count > 0)
                        {
                            if (SM[0].BOM.ToLower().Contains("вывод"))
                            {
                                M.MaterialStatus = 0;
                            }
                        }

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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }

                        M.OldTaskNav = 0; // sListOld.Sum(x => x.PlanQuantity - x.FactQuantity);   //Fact
                        M.OldTaskNavUD = 0;

                        CFMList.Add(M);

                        MesList.Add(s);
                    }

                    /*if (MList.Count > 0)
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
                        M.DontShow = false;

                        if (M.MaterialStatus > 7)
                        {
                            M.MaterialStatus = 7;
                        }
                    }
                    else
                    {

                    }*/
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

            CFMListTan = new List<string>();

            //итоги
            foreach (var ml in CFMList)
            {
                /*  if (ml.MaterialCode == "1031006114")
                  { }*/

                //1
                if (URSCodeList.Contains(ml.MaterialCode))
                {
                    if (ml.MaterialStatus > 1)
                    {
                        ml.MaterialStatus = 1;
                    }
                }

                //2
                if (URMCodeList.Contains(ml.MaterialCode))
                {
                    if (ml.MaterialStatus > 2)
                    {
                        ml.MaterialStatus = 2;
                    }
                }

                //3
                if (fm.OPRList.Contains(ml.MaterialCode))
                {
                    if (ml.MaterialStatus > 3)
                    {
                        ml.MaterialStatus = 3;
                    }
                }

                //5
                var LRX = fm.lRec.Where(x => x.MaterialCode == ml.MaterialCode).ToList();
                if (LRX.Count == 0)
                {
                    var LRX1 = fm.lRec.Where(x => x.ProductCode == ml.MaterialCode).ToList();

                    if (LRX1.Count == 0)
                    {
                        CFMListTan.Add(ml.MaterialCode);

                        if (ml.MaterialStatus > 5)
                        {
                            ml.MaterialStatus = 5;
                        }
                    }
                }

                //8
                if ((ml.WaitingDaysString == "") || (ml.MaterialMultiplyString == ""))  //||(cfm1.StorageQuantityString=="")
                {
                    CFMListGray.Add(ml.MaterialCode);

                    if (ml.MaterialStatus > 8)
                    {
                        ml.MaterialStatus = 8;
                    }
                }

                var sbList8 = sbl.Where(x => x.MaterialCode == ml.MaterialCode).ToList();    // fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList();//остатки НАВ
                //   sbList8 = sbList8.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();

                var sbListMes8 = mml.Where(x => x.MaterialCode == ml.MaterialCode).ToList();    // fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList(); //остатки склады

                var sbListMes8Free = sbListMes8.Where(x => x.TestQualityGr != "ЗБлок").ToList();                    //Where(x => x.TestQuality.ToLower() == "свободно" || x.TestQuality.ToLower() == "допущено окк").ToList();                                                                                //  sbListMes8=sbListMes8.Where(x => x.TestQuality != "3Блок").ToList();
                var sbListMes8Block = sbListMes8.Where(x => x.TestQualityGr == "ЗБлок").ToList();            //(x => x.TestQuality.ToLower() != "свободно" && x.TestQuality.ToLower() != "допущено окк").ToList();

                var sbpList8 = sbp.Where(x => x.MaterialCode == ml.MaterialCode).ToList();   // fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == ml.MaterialCode)).ToList();  //остатки пр-во
                                                                                             //     sbpList8 = sbpList8.Where(x => fm.StorageNamesList.Contains(x.LineName)).ToList();

                var sbListIn = sbList8.Where(x => x.StorageType == "Собств").ToList();  //Собств по НАВ
                var sbListOut = sbList8.Where(x => x.StorageType != "Собств").ToList(); //Внеш

                /* if (sbListIn.Count>0)
                 {
                     MessageBox.Show("Zero");
                 }*/

                var sbListInBlock = sbListIn.Where(x => x.TestQualityGr == "3Блок").ToList(); //забл собств   sbListMes8Block!!!!! TesyQualityGr!!!!!!!
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

                ml.RawStorageListAlter = new List<RawsAlter>();
                ml.RawEnterpriseListAlter = new List<RawsAlter>();

                ml.ListLine = new List<int>();
                ml.ListPlus = new List<int>();

                ml.RawStorage = sbListMes8Free.Sum(x => (double)x.Quantity * -1) - sbListInBlock.Sum(x => (double)x.Quantity);  /*- sbListOutBlock.Sum(x => x.Quantity)*/   // sbList8.Sum(x => x.Quantity) - ...
                ml.RawEnterprise = (double)sbpList8.Sum(x => x.Quantity);
                ml.RawNav1 = (double)sbListOut.Sum(x => x.Quantity) - (double)sbListOutBlock.Sum(x => x.Quantity);  //с учетом заблокированных!
                ml.RawNav2 = (double)sbListIn.Sum(x => x.Quantity) - (double)sbListInBlock.Sum(x => x.Quantity);
                ml.RawMesOut1 = (double)sbListOutBlock.Sum(x => x.Quantity);
                ml.RawMesOut2 = (double)sbListInBlock.Sum(x => x.Quantity) - (double)sbListMes8Block.Sum(x => x.Quantity);
                ml.OldTaskNav = sListOld.Sum(x => (x.PlanQuantity - x.FactQuantity));

                ml.RawStorageUD = sbListMes8Free.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => (double)x.Quantity * -1) - sbListInBlock.Sum(x => (double)x.Quantity) /*- sbListOutBlock.Sum(x => x.Quantity)*/;   //sbList8.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity)
                ml.RawEnterpriseUD = (double)sbpList8.Where(x => x.ValidTo >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawNav1UD = (double)sbListOut.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity) - (double)sbListOutBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawNav2UD = (double)sbListIn.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity) - (double)sbListInBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawMesOut1UD = (double)sbListOutBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
                ml.RawMesOut2UD = (double)sbListInBlock.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity) - (double)sbListMes8Block.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);
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

                //raw StorageAlter
                ml.RawStorageListAlter = sbListMes8Free.GroupBy(x => new { x.StorageID, x.Storage, x.LotName, x.LotDescription, x.MaterialCode, x.MaterialName, x.ProdDate, x.BBFDate, x.TestQuality }).
                    Select(g => new RawsAlter
                    {
                        StorageID = g.Key.StorageID,
                        Storage = g.Key.Storage,
                        MaterialName = g.Key.MaterialName,
                        MaterialCode = g.Key.MaterialCode,
                        LotName = g.Key.LotName,
                        LotDescr = g.Key.LotDescription,
                        ProdDate = g.Key.ProdDate,
                        BBFDate = g.Key.BBFDate,
                        TestQuality = g.Key.TestQuality,
                        Quantity = g.Sum(x => (double)x.Quantity * (-1))
                    }).ToList();

                //raw Storage
                foreach (var rs in sbListMes8Free)  //остатки склад     sbListMes8
                {
                    bool falg = false;

                    foreach (var sb in sbListInBlock)
                    {

                        if (sb.LotName == rs.LotName)
                        {
                            falg = true;
                        }
                    }

                    if (falg == false)
                    {
                        Raws r = new Raws
                        {
                            StorageID = rs.StorageID,
                            Storage = rs.Storage,
                            LotName = rs.LotName,
                            MaterialCode = rs.MaterialCode,
                            MaterialName = rs.MaterialName,
                            LotDescr = rs.LotName,
                            BBFDate = rs.BBFDate,
                            ProdDate = rs.ProdDate,
                            Quantity = (double)rs.Quantity * (-1)
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

                //raw EnterProseAlter
                ml.RawEnterpriseListAlter = sbpList8.GroupBy(x => new { x.LineID, x.LineName, x.MaterialLotName, x.MaterialCode, x.MaterialName, x.AlterLotName, x.ValidFrom, x.ValidTo, x.TestQuality }).
                    Select(g => new RawsAlter
                    {
                        StorageID = g.Key.LineID,
                        Storage = g.Key.LineName,
                        MaterialName = g.Key.MaterialName,
                        MaterialCode = g.Key.MaterialCode,
                        LotName = g.Key.MaterialLotName,
                        LotDescr = g.Key.AlterLotName,
                        ProdDate = g.Key.ValidFrom,
                        BBFDate = g.Key.ValidTo,
                        TestQuality = g.Key.TestQuality,
                        Quantity = g.Sum(x => (double)x.Quantity)
                    }).ToList();

                //raw EnterPrise
                foreach (var rs1 in sbpList8)  //остатки пр-во
                {
                    Raws r = new Raws
                    {
                        StorageID = rs1.LineID,
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
                        Quantity = (double)rs5.Quantity,
                        StatusMZP = rs5.TestQuality
                    };

                    ml.RawMesOut2List.Add(r);
                }

                foreach (var sb8B in sbListMes8Block)
                {
                    Raws r = new Raws();

                    r.Storage = sb8B.Storage;
                    r.LotName = sb8B.LotName;
                    r.LotDescr = sb8B.LotDescription;
                    r.BBFDate = sb8B.BBFDate;
                    r.ProdDate = sb8B.ProdDate;
                    r.Quantity = -1 * (double)sb8B.Quantity;
                    r.StatusMZP = sb8B.TestQuality;

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
                            MZP = rs6.MZP
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

                        if (rs6.StatusMZP.Trim() == "")
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
                            Quantity = rs7.PlanQuantity - rs7.FactQuantity,
                            AlterDate = rs7.PlanOperDate,         // new DateTime(rs7.PlanOperDate.Year, rs7.PlanOperDate.Month, 1),
                            AddDate = rs7.PlanOperDate.AddDays(1),
                            Initiator = rs7.Initiator,
                            StatusMZP = rs7.StatusMZP,
                            MZP = rs7.MZP
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
                        Storage = cl.LotNo,
                        StatusMZP = cl.LotDescription,
                        Quantity = cl.Quantity,
                        Initiator = "",
                        MZP = "",
                        BBFDate = cl.ExpDT,
                        ProdDate = cl.ExpDT
                    };

                    ml.ListC.Add(r);
                    ml.ListD.Add(r);
                }
            }

            while (XlDataStart < XlDataEnd)
            {
                DateTime CurrDate = XlDataStart.AddMonths(1).AddDays(-1);

                foreach (var ml in CFMList)
                {
                    var xData = XLDataList.Where(x => (x.DateWork >= XlDataStart) && (x.DateWork <= CurrDate) && (x.ProdCode == ml.MaterialCode)).ToList();

                    /* if (ml.MaterialCode=="3000106749")
                     {
                         if (XlDataStart == Convert.ToDateTime("2022-10-01"))
                         { }
                     }*/

                    double Value = xData.Sum(x => x.Quantity);
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
               
                    ml.Value_ = ml.CurrentConsumption + ml.RawNav1 + ml.RawStorage;  //+ ml.RawEnterprise;
                    ml.ValueUD = ml.CurrentConsumption + ml.RawNav1 + ml.RawStorageUD; // + ml.RawEnterpriseUD;
               
                    ml.ValueEnterp = ml.CurrentConsumption + ml.RawNav1 + ml.RawStorage + ml.RawEnterprise;
                    ml.ValueEnterpUD = ml.CurrentConsumption + ml.RawNav1 + ml.RawStorageUD + ml.RawEnterpriseUD;            

                /* ml.Value_ = ml.CurrentConsumption + ml.RawNav1 + ml.RawStorageList.Sum(x => x.Quantity) + ml.RawEnterpriseList.Sum(x => x.Quantity);
                 ml.ValueUD = ml.CurrentConsumption + ml.RawNav1
                          + ml.RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity)
                          + ml.RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).Sum(x => x.Quantity);*/  //исключить просрочку

                sVal = ml.MaterialCode + RowSplitter + ml.MaterialName + RowSplitter +
                      ml.MaterialGroup + RowSplitter + ml.MaterialMultiplyString + RowSplitter +
                      ml.WaitingDaysString + RowSplitter + ml.StorageQuantityString + RowSplitter +
                      ml.Responsible + RowSplitter + ml.Sender + RowSplitter;

                for (int j = 0; j < WorkMonthList.Count; j++)
                {
                    if ((Math.Abs(ml.OutList[j].Quantity)  >= 100))
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N0") + RowSplitter;
                    }
                    else if ((Math.Abs(ml.OutList[j].Quantity) < 100) && (Math.Abs(ml.OutList[j].Quantity) >= 10))
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N1") + RowSplitter;
                    }
                    else if ((Math.Abs(ml.OutList[j].Quantity) < 10) && (Math.Abs(ml.OutList[j].Quantity) >= 1))
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N2") + RowSplitter;
                    }
                    else if (Math.Abs(ml.OutList[j].Quantity) > 0)
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N3") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + "" + RowSplitter;
                    }

                    /*if ((Math.Abs(ml.OutList[j].Quantity) <= 1) && (Math.Abs(ml.OutList[j].Quantity) > 0))
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N1") + RowSplitter;
                    }
                    else
                    {
                        sVal = sVal + ml.OutList[j].Quantity.ToString("N0") + RowSplitter;
                    }*/
                }


                if ((Math.Abs(ml.CurrentConsumption) >= 100))  //8
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.CurrentConsumption) < 100) && (Math.Abs(ml.CurrentConsumption) >= 10))
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.CurrentConsumption) < 10) && (Math.Abs(ml.CurrentConsumption) >= 1))
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.CurrentConsumption) > 0)
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.CurrentConsumption) <= 1) && (Math.Abs(ml.CurrentConsumption) > 0)) //8
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + ml.CurrentConsumption.ToString("N0") + RowSplitter;
                }*/

                sValUD = sVal;

                if ((Math.Abs(ml.RawStorage) >= 100))  //9
                {
                    sVal = sVal + ml.RawStorage.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawStorage) < 100) && (Math.Abs(ml.RawStorage) >= 10))
                {
                    sVal = sVal + ml.RawStorage.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawStorage) < 10) && (Math.Abs(ml.RawStorage) >= 1))
                {
                    sVal = sVal + ml.RawStorage.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawStorage) > 0)
                {
                    sVal = sVal + ml.RawStorage.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawStorage) <= 1) && (ml.RawStorage != 0)) //9
                {
                    sVal = sVal + ml.RawStorage.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + ml.RawStorage.ToString("N0") + RowSplitter;
                }*/

                if ((Math.Abs(ml.RawEnterprise) >= 100))  //10
                {
                    sVal = sVal + ml.RawEnterprise.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawEnterprise) < 100) && (Math.Abs(ml.RawEnterprise) >= 10))
                {
                    sVal = sVal + ml.RawEnterprise.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawEnterprise) < 10) && (Math.Abs(ml.RawEnterprise) >= 1))
                {
                    sVal = sVal + ml.RawEnterprise.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawEnterprise) > 0)
                {
                    sVal = sVal + ml.RawEnterprise.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }

                /*   if ((Math.Abs(ml.RawEnterprise) <= 1) && (ml.RawEnterprise != 0))  //10
               {
                   sVal = sVal + ml.RawEnterprise.ToString("N1") + RowSplitter;
               }
               else
               {
                   sVal = sVal + ml.RawEnterprise.ToString("N0") + RowSplitter;
               }*/

                if ((Math.Abs(ml.RawNav1) >= 100))  //11
                {
                    sVal = sVal + ml.RawNav1.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav1) < 100) && (Math.Abs(ml.RawNav1) >= 10))
                {
                    sVal = sVal + ml.RawNav1.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav1) < 10) && (Math.Abs(ml.RawNav1) >= 1))
                {
                    sVal = sVal + ml.RawNav1.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawNav1) > 0)
                {
                    sVal = sVal + ml.RawNav1.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }


                /*if ((Math.Abs(ml.RawNav1) <= 1) && (ml.RawNav1 != 0)) //внешние 11
                {
                    sVal = sVal + ml.RawNav1.ToString("N1") + RowSplitter;   // " (";
                }
                else
                {
                    sVal = sVal + ml.RawNav1.ToString("N0") + RowSplitter;  //" (";
                }*/

                if ((Math.Abs(ml.RawNav2) >= 100))  //12
                {
                    sVal = sVal + ml.RawNav2.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav2) < 100) && (Math.Abs(ml.RawNav2) >= 10))
                {
                    sVal = sVal + ml.RawNav2.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav2) < 10) && (Math.Abs(ml.RawNav2) >= 1))
                {
                    sVal = sVal + ml.RawNav2.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawNav2) > 0)
                {
                    sVal = sVal + ml.RawNav2.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }


                /*if ((Math.Abs(ml.RawNav2) <= 1) && (ml.RawNav2 != 0))  //собств по НАВ   12
                {
                    sVal = sVal + ml.RawNav2.ToString("N1") + RowSplitter;  //+ ") "
                }
                else
                {
                    sVal = sVal + ml.RawNav2.ToString("N0") + RowSplitter;  //+ ") "
                }*/

                if ((Math.Abs(ml.RawMesOut1) >= 100))  //13
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut1) < 100) && (Math.Abs(ml.RawMesOut1) >= 10))
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut1) < 10) && (Math.Abs(ml.RawMesOut1) >= 1))
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawMesOut1) > 0)
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawMesOut1) <= 1) && (ml.RawMesOut1 != 0)) //забл внешние 13
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N1") + RowSplitter; // " (";
                }
                else
                {
                    sVal = sVal + ml.RawMesOut1.ToString("N0") + RowSplitter; // " (";
                }*/

                if ((Math.Abs(ml.RawMesOut2) >= 100))  //14
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut2) < 100) && (Math.Abs(ml.RawMesOut2) >= 10))
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut2) < 10) && (Math.Abs(ml.RawMesOut2) >= 1))
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawMesOut2) > 0)
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawMesOut2) <= 1) && (ml.RawMesOut2 != 0))   // забл  МЕС  14
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N1") + RowSplitter;  // + ")" 
                }
                else
                {
                    sVal = sVal + ml.RawMesOut2.ToString("N0") + RowSplitter;  //+ ")"
                }*/

                if ((Math.Abs(ml.OldTaskNav) >= 100))  //15
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.OldTaskNav) < 100) && (Math.Abs(ml.OldTaskNav) >= 10))
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.OldTaskNav) < 10) && (Math.Abs(ml.RawMesOut2) >= 1))
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.OldTaskNav) > 0)
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N3") + RowSplitter;
                }
                else
                {
                    sVal = sVal + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.OldTaskNav) <= 1) && (ml.OldTaskNav != 0))  //15 старые заявки НАВ
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N1") + RowSplitter;   //0-15
                }
                else
                {
                    sVal = sVal + ml.OldTaskNav.ToString("N0") + RowSplitter;   //0-15
                }*/

                //UD
                if ((Math.Abs(ml.RawStorageUD) >= 100))  //
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawStorageUD) < 100) && (Math.Abs(ml.RawStorageUD) >= 10))
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawStorageUD) < 10) && (Math.Abs(ml.RawStorageUD) >= 1))
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawStorageUD) > 0)
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawStorageUD) <= 1) && (ml.RawStorageUD != 0))
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N1") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + ml.RawStorageUD.ToString("N0") + RowSplitter;
                }*/

                if ((Math.Abs(ml.RawEnterpriseUD) >= 100))  //
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawEnterpriseUD) < 100) && (Math.Abs(ml.RawEnterpriseUD) >= 10))
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawEnterpriseUD) < 10) && (Math.Abs(ml.RawEnterpriseUD) >= 1))
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawEnterpriseUD) > 0)
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawEnterpriseUD) <= 1) && (ml.RawEnterpriseUD != 0))
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N1") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + ml.RawEnterpriseUD.ToString("N0") + RowSplitter;
                }*/

                if ((Math.Abs(ml.RawNav1UD) >= 100))  //
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav1UD) < 100) && (Math.Abs(ml.RawNav1UD) >= 10))
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav1UD) < 10) && (Math.Abs(ml.RawNav1UD) >= 1))
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawNav1UD) > 0)
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawNav1UD) <= 1) && (ml.RawNav1UD != 0))
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N1") + RowSplitter;  //+ " ("
                }
                else
                {
                    sValUD = sValUD + ml.RawNav1UD.ToString("N0") + RowSplitter; // +" ("
                }*/

                if ((Math.Abs(ml.RawNav2UD) >= 100))  //
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav2UD) < 100) && (Math.Abs(ml.RawNav2UD) >= 10))
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawNav2UD) < 10) && (Math.Abs(ml.RawNav2UD) >= 1))
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawNav2UD) > 0)
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawNav2UD) <= 1) && (ml.RawNav2UD != 0))
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N1") + RowSplitter;  // + ") "
                }
                else
                {
                    sValUD = sValUD + ml.RawNav2UD.ToString("N0") + RowSplitter;  // +") "
                }*/

                if ((Math.Abs(ml.RawMesOut1UD) >= 100))  //
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut1UD) < 100) && (Math.Abs(ml.RawMesOut1UD) >= 10))
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut1UD) < 10) && (Math.Abs(ml.RawMesOut1UD) >= 1))
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawMesOut1UD) > 0)
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawMesOut1UD) <= 1) && (ml.RawMesOut1UD != 0))
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N1") + RowSplitter;  //+ " ("
                }
                else
                {
                    sValUD = sValUD + ml.RawMesOut1UD.ToString("N0") + RowSplitter;  //+ " ("
                }*/

                if ((Math.Abs(ml.RawMesOut2UD) >= 100))  //
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut2UD) < 100) && (Math.Abs(ml.RawMesOut2UD) >= 10))
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.RawMesOut2UD) < 10) && (Math.Abs(ml.RawMesOut2UD) >= 1))
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.RawMesOut2UD) > 0)
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.RawMesOut2UD) <= 1) && (ml.RawMesOut2UD != 0))
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N1") + RowSplitter;  //") " +
                }
                else
                {
                    sValUD = sValUD + ml.RawMesOut2UD.ToString("N0") + RowSplitter;  //+ ") "
                }*/

                if ((Math.Abs(ml.OldTaskNavUD) >= 100))  //
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N0") + RowSplitter;
                }
                else if ((Math.Abs(ml.OldTaskNavUD) < 100) && (Math.Abs(ml.OldTaskNavUD) >= 10))
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N1") + RowSplitter;
                }
                else if ((Math.Abs(ml.OldTaskNavUD) < 10) && (Math.Abs(ml.OldTaskNavUD) >= 1))
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N2") + RowSplitter;
                }
                else if (Math.Abs(ml.OldTaskNavUD) > 0)
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N3") + RowSplitter;
                }
                else
                {
                    sValUD = sValUD + "" + RowSplitter;
                }

                /*if ((Math.Abs(ml.OldTaskNavUD) <= 1) && (ml.OldTaskNavUD != 0))
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N1") + RowSplitter;   //0-15
                }
                else
                {
                    sValUD = sValUD + ml.OldTaskNavUD.ToString("N0") + RowSplitter;   //0-15
                }*/

                ml.ShortString = sVal;
                ml.ShortStringUD = sValUD;

                if (ml.DontShow == true)
                {
                    string V;
                    string vUD;
                    string V_RE;
                    string VUD_RE;

                    ml.ValueS = "";
                    ml.ValueUDS = "";
                    ml.AlterValueS = "";
                    ml.AlterValueUDS = "";

                    ml.AlterValueS_RE = "";
                    ml.ValueUDS_RE = "";
                    ml.ValueS_RE = "";
                    ml.AlterValueUDS_RE = "";

                    double Q = ml.Value_;
                    double QUD = ml.ValueUD;

                    double QRE = ml.ValueEnterp;
                    double QUDRE = ml.ValueEnterpUD;

                    if ((Math.Abs(ml.Value_) >= 100))  //
                    {
                        V =  ml.Value_.ToString("N0") ;
                    }
                    else if ((Math.Abs(ml.Value_) < 100) && (Math.Abs(ml.Value_) >= 10))
                    {
                        V = ml.Value_.ToString("N1")  ;
                    }
                    else if ((Math.Abs(ml.Value_) < 10) && (Math.Abs(ml.Value_) >= 1))
                    {
                        V = ml.Value_.ToString("N2");
                    }
                    else if (Math.Abs(ml.Value_) > 0)
                    {
                        V =  ml.Value_.ToString("N3") ;
                    }
                    else
                    {
                        V =  "";
                    }


                    /*if ((Math.Abs(ml.Value_) <= 1) && (ml.Value_ != 0))
                    {
                        V = ml.Value_.ToString("N1");
                    }
                    else
                    {
                        V = ml.Value_.ToString("N0");
                    }*/

                    if ((Math.Abs(ml.ValueUD) >= 100))  //
                    {
                        vUD = ml.ValueUD.ToString("N0");
                    }
                    else if ((Math.Abs(ml.ValueUD) < 100) && (Math.Abs(ml.ValueUD) >= 10))
                    {
                        vUD = ml.ValueUD.ToString("N1");
                    }
                    else if ((Math.Abs(ml.ValueUD) < 10) && (Math.Abs(ml.ValueUD) >= 1))
                    {
                        vUD = ml.ValueUD.ToString("N2");
                    }
                    else if (Math.Abs(ml.ValueUD) > 0)
                    {
                       vUD =  ml.ValueUD.ToString("N3") ;
                    }
                    else
                    {
                        vUD = "";
                    }

                    /*if ((Math.Abs(ml.ValueUD) <= 1) && (ml.ValueUD != 0))
                    {
                        vUD = ml.ValueUD.ToString("N1");
                    }
                    else
                    {
                        vUD = ml.ValueUD.ToString("N0");
                    }*/

                    if ((Math.Abs(ml.ValueEnterp) >= 100))  //
                    {
                        V_RE = ml.ValueEnterp.ToString("N0");
                    }
                    else if ((Math.Abs(ml.ValueEnterp) < 100) && (Math.Abs(ml.ValueEnterp) >= 10))
                    {
                        V_RE = ml.ValueEnterp.ToString("N1");
                    }
                    else if ((Math.Abs(ml.ValueEnterp) < 10) && (Math.Abs(ml.ValueEnterp) >= 1))
                    {
                        V_RE = ml.ValueEnterp.ToString("N2");
                    }
                    else if (Math.Abs(ml.ValueEnterp) > 0)
                    {
                        V_RE = ml.ValueEnterp.ToString("N3")  ;
                    }
                    else
                    {
                        V_RE = "";
                    }

                    /* if ((Math.Abs(ml.ValueEnterp) <= 1) && (ml.ValueEnterp != 0))
                     {
                         V_RE = ml.ValueEnterp.ToString("N1");
                     }
                     else
                     {
                         V_RE = ml.ValueEnterp.ToString("N0");
                     }*/

                    if ((Math.Abs(ml.ValueEnterpUD) >= 100))  //
                    {
                        VUD_RE = ml.ValueEnterpUD.ToString("N0");
                    }
                    else if ((Math.Abs(ml.ValueEnterpUD) < 100) && (Math.Abs(ml.ValueEnterpUD) >= 10))
                    {
                        VUD_RE = ml.ValueEnterpUD.ToString("N1");
                    }
                    else if ((Math.Abs(ml.ValueEnterpUD) < 10) && (Math.Abs(ml.ValueEnterpUD) >= 1))
                    {
                        VUD_RE = ml.ValueEnterpUD.ToString("N2");
                    }
                    else if (Math.Abs(ml.ValueEnterpUD) > 0)
                    {
                        VUD_RE = ml.ValueEnterpUD.ToString("N3");
                    }
                    else
                    {
                        VUD_RE = "";
                    }

                    /*if ((Math.Abs(ml.ValueEnterpUD) <= 1) && (ml.ValueEnterpUD != 0))
                    {
                        VUD_RE = ml.ValueEnterpUD.ToString("N1");
                    }
                    else
                    {
                        VUD_RE = ml.ValueEnterpUD.ToString("N0");
                    }*/

                    foreach (DateTime dt in DTList)
                    {
                        ml.AlterValueS = ml.AlterValueS + vUD + "" + RowSplitter;
                        ml.AlterValueUDS = ml.AlterValueUDS + vUD + "" + RowSplitter;
                        ml.AlterValueS_RE = ml.AlterValueS_RE + V_RE + " " + RowSplitter;
                        ml.AlterValueUDS_RE = ml.AlterValueUDS_RE + VUD_RE + " " + RowSplitter;
                    }

                    for (int k = 0; k < DTList.Count; k++)
                    {
                        double T2 = 0;
                        double T1 = 0;

                        DateTime dt = DTList[k];
                        DateTime dtl;

                        if (k + 1 < DTList.Count)
                        {
                            dtl = DTList[k + 1];
                        }
                        else
                        {
                            dtl = DTList.Last().AddYears(1);
                        }

                        var Delta = (dtl.Date - dt.Date).TotalDays;

                        var FDTList = DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();

                        if ((Math.Abs(Q) >= 100))  //
                        {
                            V = Q.ToString("N0");
                        }
                        else if ((Math.Abs(Q) < 100) && (Math.Abs(Q) >= 10))
                        {
                            V = Q.ToString("N1");
                        }
                        else if ((Math.Abs(Q) < 10) && (Math.Abs(Q) >= 1))
                        {
                            V = Q.ToString("N2");
                        }
                        else if (Math.Abs(Q) > 0)
                        {
                            V = Q.ToString("N3");
                        }
                        else
                        {
                            V = "";
                        }

                        /* if ((Math.Abs(Q) <= 1) && (Q != 0))
                         {
                             V = Q.ToString("N1");
                         }
                         else
                         {
                             V = Q.ToString("N0");
                         }*/

                        if ((Math.Abs(QUD) >= 100))  //
                        {
                            vUD = QUD.ToString("N0");
                        }
                        else if ((Math.Abs(QUD) < 100) && (Math.Abs(QUD) >= 10))
                        {
                            vUD = QUD.ToString("N1");
                        }
                        else if ((Math.Abs(QUD) < 10) && (Math.Abs(QUD) >= 1))
                        {
                            vUD = QUD.ToString("N2");
                        }
                        else if (Math.Abs(QUD) > 0)
                        {
                            vUD = QUD.ToString("N3");
                        }
                        else
                        {
                            vUD = "";
                        }

                        /* if ((Math.Abs(QUD) <= 1) && (QUD != 0))
                         {
                             vUD = QUD.ToString("N1");
                         }
                         else
                         {
                             vUD = QUD.ToString("N0");
                         }*/

                        if ((Math.Abs(QRE) >= 100))  //
                        {
                            V_RE = QRE.ToString("N0");
                        }
                        else if ((Math.Abs(QRE) < 100) && (Math.Abs(QRE) >= 10))
                        {
                            V_RE = QRE.ToString("N1");
                        }
                        else if ((Math.Abs(QRE) < 10) && (Math.Abs(QRE) >= 1))
                        {
                            V_RE = QRE.ToString("N2");
                        }
                        else if (Math.Abs(QRE) > 0)
                        {
                            V_RE = QRE.ToString("N3");
                        }
                        else
                        {
                            V_RE = "";
                        }

                        /*if ((Math.Abs(QRE) <= 1) && (QRE != 0))
                        {
                            V_RE = QRE.ToString("N1");
                        }
                        else
                        {
                            V_RE = QRE.ToString("N0");
                        }*/

                        if ((Math.Abs(QUDRE) >= 100))  //
                        {
                            VUD_RE = QUDRE.ToString("N0");
                        }
                        else if ((Math.Abs(QUDRE) < 100) && (Math.Abs(QUDRE) >= 10))
                        {
                            VUD_RE = QUDRE.ToString("N1");
                        }
                        else if ((Math.Abs(QUDRE) < 10) && (Math.Abs(QUDRE) >= 1))
                        {
                            VUD_RE = QUDRE.ToString("N2");
                        }
                        else if (Math.Abs(QUDRE) > 0)
                        {
                            VUD_RE = QUDRE.ToString("N3");
                        }
                        else
                        {
                            VUD_RE = "";
                        }

                        /*if ((Math.Abs(QUDRE) <= 1) && (QUDRE != 0))
                        {
                            VUD_RE = QUDRE.ToString("N1");
                        }
                        else
                        {
                            VUD_RE = QUDRE.ToString("N0");
                        }*/


                        ml.ValueS = ml.ValueS + vUD;
                        ml.ValueUDS = ml.ValueUDS + V;
                        ml.ValueS_RE = ml.ValueS_RE + V_RE;
                        ml.ValueUDS_RE = ml.ValueUDS_RE + VUD_RE;

                        List<Raws> tData2 = new List<Raws>();  //asks
                        if  (Delta > 1)        //((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                        {
                            tData2 = ml.TaskNavList2.Where(x => x.AlterDate >= dt && x.AlterDate < dtl).ToList();
                        }
                        else
                        {
                            tData2 = ml.TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                        }
                        T2 = T2 + tData2.Sum(x => x.Quantity);

                        List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();
                        if  (Delta > 1)    //((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == ml.MaterialCode) && (x.AlterDate >= dt && x.AlterDate < dtl)).ToList();   //(((DateTime)x.AlterDate).Date == dt.Date)
                        }
                        else
                        {
                            Cdata = fm.ContList.Where(x => (x.MaterialCode == ml.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                        }
                        T2 = T2 + Cdata.Sum(x => x.Quantity);

                        List<Raws> tData1 = new List<Raws>(); //ask in progress
                        if (Delta > 1)     //((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                        {
                            tData1 = ml.TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                        }
                        else
                        {
                            tData1 = ml.TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                        }
                        T1 = T1 + tData1.Sum(x => x.Quantity);

                        if (T2 > 0)
                        {
                            foreach (var t in tData2)
                            {
                                if (Math.Abs(t.Quantity) >= 100)
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueS_RE = ml.ValueS_RE + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " (" + t.Quantity.ToString("N0") + ") ";
                                }
                                else if (Math.Abs(t.Quantity) < 100 & Math.Abs(t.Quantity) >= 10)
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueS_RE = ml.ValueS_RE + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " (" + t.Quantity.ToString("N1") + ") ";
                                }
                                else if (Math.Abs(t.Quantity) < 10 && Math.Abs(t.Quantity) >= 1)
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N2") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N2") + ") ";
                                    ml.ValueS_RE = ml.ValueS_RE + " (" + t.Quantity.ToString("N2") + ") ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " (" + t.Quantity.ToString("N2") + ") ";
                                }
                                else if (Math.Abs(t.Quantity) >0)
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N3") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N3") + ") ";
                                    ml.ValueS_RE = ml.ValueS_RE + " (" + t.Quantity.ToString("N3") + ") ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " (" + t.Quantity.ToString("N3") + ") ";
                                }

                                /*if ((Math.Abs(t.Quantity) <= 1) && (t.Quantity != 0))
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueS_RE = ml.ValueS_RE + " (" + t.Quantity.ToString("N1") + ") ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " (" + t.Quantity.ToString("N1") + ") ";
                                }
                                else
                                {
                                    ml.ValueS = ml.ValueS + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueUDS = ml.ValueUDS + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueS_RE = ml.ValueS_RE + " (" + t.Quantity.ToString("N0") + ") ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " (" + t.Quantity.ToString("N0") + ") ";
                                }*/
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

                        if (T1 > 0)
                        {
                            foreach (var t in tData1)
                            {
                                if (Math.Abs(t.Quantity) >= 100)
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueS_RE = ml.ValueS_RE + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " [" + t.Quantity.ToString("N0") + "] ";
                                }
                                else if (Math.Abs(t.Quantity) < 100 & Math.Abs(t.Quantity) >= 10)
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueS_RE = ml.ValueS_RE + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " [" + t.Quantity.ToString("N1") + "] ";
                                }
                                else if (Math.Abs(t.Quantity) < 10 && Math.Abs(t.Quantity) >= 1)
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N2") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N2") + "] ";
                                    ml.ValueS_RE = ml.ValueS_RE + " [" + t.Quantity.ToString("N2") + "] ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " [" + t.Quantity.ToString("N2") + "] ";
                                }
                                else if (Math.Abs(t.Quantity) > 0)
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N3") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N3") + "] ";
                                    ml.ValueS_RE = ml.ValueS_RE + " [" + t.Quantity.ToString("N3") + "] ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " [" + t.Quantity.ToString("N3") + "] ";
                                }


                                /*if ((Math.Abs(t.Quantity) <= 1) && (t.Quantity != 0))
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueS_RE = ml.ValueS_RE + " [" + t.Quantity.ToString("N1") + "] ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " [" + t.Quantity.ToString("N1") + "] ";
                                }
                                else
                                {
                                    ml.ValueS = ml.ValueS + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueUDS = ml.ValueUDS + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueS_RE = ml.ValueS_RE + " [" + t.Quantity.ToString("N0") + "] ";
                                    ml.ValueUDS_RE = ml.ValueUDS_RE + " [" + t.Quantity.ToString("N0") + "] ";
                                }*/
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
                        ml.ValueUDS_RE = ml.ValueUDS_RE + RowSplitter;
                        ml.ValueS_RE = ml.ValueS_RE + RowSplitter;
                    }
                }
            }

            CFMList = CFMList.OrderBy(x => x.MaterialName).ToList();  //Where(x => x.MaterialCode == "8СИ00009").ToList();       //

            AppList = fm.tdb.tApplication.ToList();
            StorageTaskList = fm.tdb.tStorageTask.ToList();

            dt1 = new DataTable();

            GC.Collect();
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

        #region Hidden
        /*private void ShowLocalData()
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
                            foreach (var sl in cfm.RawStorageListAlter)
                            {
                                Value = Value + sl.Quantity;
                            }
                        }
                        else
                        {
                            foreach (var sl in cfm.RawStorageListAlter.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                            {
                                Value = Value + sl.Quantity;
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

                                    //  Value = Value + Q;
                                }
                            }

                            if (Value < 0)
                            {
                              //  if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                {
                                    //  CFMBool[i] = false;
                                    if (cfm.MaterialStatus > 4)    //(!CFMListGold.Contains(cfm.MaterialCode))  //-
                                    {
                                        // CFMListGold.Add(cfm.MaterialCode);
                                        CFMListRed.Add(cfm.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                        cfm.MaterialStatus = 4;
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
                            foreach (var sl in cfm.RawStorageListAlter)
                            {
                                Value = Value + sl.Quantity;
                            } 
                        }
                        else
                        {
                            foreach (var sl in cfm.RawStorageListAlter.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                            {
                                Value = Value + sl.Quantity;
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
                                }
                            }

                            if (Value < 0)
                            {
                               // if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                {
                                    //  CFMBool[i] = false;
                                    if   (cfm.MaterialStatus > 4)   //(!CFMListGold.Contains(cfm.MaterialCode))  //-
                                    {
                                      //  CFMListGold.Add(cfm.MaterialCode);
                                        CFMListRed.Add(cfm.MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cfm.MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        CFMListMinus.Add(Col);
                                        cfm.MaterialStatus = 4;
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
        }*/

        # endregion

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

            CFMListGold = new List<string>(); //Value<0
        //    CFMListGray = new List<string>();
            CFMListLine = new List<Colorres>(); //[
            CFMListPink = new List<string>();//[
            CFMListPlus = new List<Colorres>(); //+
            CFMListGreen = new List<string>();//+
            CFMListMinus = new List<Colorres>();//-
            CFMListRed = new List<string>(); //-
            CFMListLinePlus = new List<Colorres>();  //[+
            CFMListGoldens = new List<string>();  //[+
            CFMListPlum = new List<Colorres>(); //<
            CFMListPlumes = new List<string>(); //<

            StorageTaskList = fm.tdb.tStorageTask.ToList();
            CFMListPlumes = StorageTaskList.Select(x => x.MaterialCode).Distinct().ToList();
            

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

            dt1.Columns.Add("MatStatus");
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
                    if (cbRawEnterprise.Checked == true)
                    {
                        Value = cfm1.ValueEnterp;
                    }
                    else
                    {
                        Value = cfm1.Value_;
                    }

                    sVal = cfm1.ShortString;
                }
                else
                {
                    if (cbRawEnterprise.Checked == true)
                    {
                        Value = cfm1.ValueEnterpUD;
                    }
                    else
                    {
                        Value = cfm1.ValueUD;
                    }

                    sVal = cfm1.ShortStringUD;
                }

                /*   if ((cfm1.WaitingDaysString == "") || (cfm1.MaterialMultiplyString == "") )  //||(cfm1.StorageQuantityString=="")
                {
                    CFMListGray.Add(cfm1.MaterialCode);
                }*/

                if (cbViews.Text == "Остаток")
                {
                    Cols = 16 + WorkMonthList.Count;  //14

                    if (cfm1.DontShow == true)
                    {
                        if (cbUseMaterialPlanning.Checked == true)
                        {
                            if (cbMaterialDelay.Checked == true)
                            {
                                if (cbRawEnterprise.Checked == true)
                                {
                                    sVal = sVal + cfm1.ValueS_RE;
                                }
                                else
                                {
                                    sVal = sVal + cfm1.ValueS;
                                }                              
                            }
                            else
                            {
                                if (cbRawEnterprise.Checked == true)
                                {
                                    sVal = sVal + cfm1.ValueUDS_RE;
                                }
                                else
                                {
                                    sVal = sVal + cfm1.ValueUDS;
                                }                             
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
                                if (cbRawEnterprise.Checked == true)
                                {
                                    sVal = sVal + cfm1.AlterValueS_RE;
                                }
                                else
                                {
                                    sVal = sVal + cfm1.AlterValueS;
                                }                             
                            }
                            else
                            {
                                if (cbRawEnterprise.Checked == true)
                                {
                                    sVal = sVal + cfm1.AlterValueUDS_RE;
                                }
                                else
                                {
                                    sVal = sVal + cfm1.AlterValueUDS;
                                }                               
                            }
                        }
                    }
                    else
                    {
                        double OldValue = 0;
                        string StringValue = "";
                        var xlData = XLDataList.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                        
                        var adata = AppList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();    // AppList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();
                        var cListData = fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode)).ToList();

                        var PlumData = StorageTaskList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();

                        // foreach (DateTime dt in DTList)
                        for (int j = 0; j < DTList.Count; j++)
                        {
                            //income
                            double T2 = 0;
                            double T1 = 0;
                            DateTime dt = DTList[j];
                            DateTime dtl;

                            if (j + 1 < DTList.Count)
                            {
                                dtl = DTList[j + 1];                                
                            }
                            else
                            {
                               dtl = DTList.Last().AddYears(1);
                            }

                            var Delta = (dtl.Date - dt.Date).TotalDays;

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

                                List<Raws> tData3 = new List<Raws>();
                                if (Delta > 1)         //((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                                {
                                    tData3 = cfm1.TaskNavList2.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();  //=
                                }

                                //  foreach (var td2 in tData2)
                                //  {
                                T2 = T2 + tData2.Sum(x => x.Quantity) + tData3.Sum(x => x.Quantity);     // td2.Quantity;
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

                                List<tPlannedMeatContainers> Cdata1 = new List<tPlannedMeatContainers>();
                                if   (Delta > 1)       //((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                                {
                                    Cdata1 = cListData.Where(x => ((DateTime)x.AlterDate).Date > dt.Date && ((DateTime)x.AlterDate).Date < dtl.Date).ToList();  //=      // fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                                }

                                // foreach (var cd in Cdata)
                                // {
                                //     T2 = T2 + cd.Quantity;
                                // }
                                T2 = T2 + Cdata.Sum(x => x.Quantity) + Cdata1.Sum(x => x.Quantity);

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

                                List<Raws> tDataA = new List<Raws>();
                                if  (Delta > 1)        //((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                                {
                                    tDataA = cfm1.TaskNavList1.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();  //=
                                }

                                //   foreach (var td1 in tData1)
                                //   {
                                //       T1 = T1 + td1.Quantity;
                                //   }
                                T1 = T1 + tData1.Sum(x => x.Quantity) + tDataA.Sum(x => x.Quantity);

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
                            List<tApplication> bData1 = new List<tApplication>();

                            if  (Delta> 1)        //((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                            {
                                bData1 = adata.Where(x => (DateTime)x.AlterDate > dt.Date && (DateTime)x.AlterDate < dtl.Date).ToList();  //=
                            }

                            Q = Q + bdata.Sum(x => x.Quantity) + bData1.Sum(x => x.Quantity);

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
                                    if  (!CFMListGold.Contains(cfm1.MaterialCode))  //-
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

                                        if (cfm1.MaterialStatus > 4)
                                        {
                                            cfm1.MaterialStatus = 4;
                                        }
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

                                if (Math.Abs(Value) >=100)
                                {
                                    StringValue = Value.ToString("N0") + " ";
                                }
                                else if (Math.Abs(Value) <100 && Math.Abs(Value) >= 10)
                                {
                                    StringValue = Value.ToString("N1") + " ";
                                }
                                else  if (Math.Abs(Value) < 10 && Math.Abs(Value) >=1)
                                {
                                    StringValue = Value.ToString("N2") + " ";
                                }
                                else if (Math.Abs(Value) > 0)
                                {
                                    StringValue = Value.ToString("N3") + " ";
                                }
                                else
                                {
                                    StringValue = " ";
                                }

                               /* if ((Math.Abs(Value) <= 1) && (Value != 0))
                                {
                                    StringValue = Value.ToString("N1") + " ";
                                }
                                else
                                {
                                    StringValue = Value.ToString("N0") + " ";
                                }*/
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
                                if (Math.Abs(T2) >= 100)
                                {
                                    sVal = sVal + " (" + T2.ToString("N0") + ") ";
                                }
                                else if (Math.Abs(T2) <100 && Math.Abs(T2) >=10)
                                {
                                    sVal = sVal + " (" + T2.ToString("N1") + ") ";
                                }
                                else if (Math.Abs(T2) < 10 && Math.Abs(T2) >=1)
                                {
                                    sVal = sVal + " (" + T2.ToString("N2") + ") ";
                                }
                                else if (Math.Abs(T2) > 0)
                                {
                                    sVal = sVal + " (" + T2.ToString("N3") + ") ";
                                }
                                                                 
                                /*if ((Math.Abs(T2) <= 1) && (T2 != 0))
                                {
                                    sVal = sVal + " (" + T2.ToString("N1") + ") ";
                                }
                                else
                                {
                                    sVal = sVal + " (" + T2.ToString("N0") + ") ";
                                }*/

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
                                if (Math.Abs(T1) >= 100)
                                {
                                    sVal = sVal + " [" + T1.ToString("N0") + "] ";
                                }
                                else if (Math.Abs(T1) <100 && Math.Abs(T1) >=10)
                                {
                                    sVal = sVal + " [" + T1.ToString("N1") + "] ";
                                }
                                else if (Math.Abs(T1) < 10 && Math.Abs(T1) >= 1)
                                {
                                    sVal = sVal + " [" + T1.ToString("N2") + "] ";
                                }
                                else if (Math.Abs(T1) > 0)
                                {
                                    sVal = sVal + " [" + T1.ToString("N3") + "] ";
                                }
                                
                                /*if ((Math.Abs(T1) <= 1) && (T1 != 0))
                                {
                                    sVal = sVal + " [" + T1.ToString("N1") + "] ";
                                }
                                else
                                {
                                    sVal = sVal + " [" + T1.ToString("N0") + "] ";
                                }*/

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
                                
                                if  (Delta > 1)    //((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                {
                                    bdata = adata.Where(x => (DateTime)x.AlterDate > dt.Date && (DateTime)x.AlterDate < dtl.Date).ToList();  //=
                                }

                                foreach (var b in bdata)
                                {
                                    if (Math.Abs(b.Quantity) >= 100)
                                    {
                                        sVal = sVal + " {" + b.Quantity.ToString("N0") + "} ";
                                    }
                                    else if (Math.Abs(b.Quantity) < 100 && Math.Abs(b.Quantity) >= 10)
                                    {
                                        sVal = sVal + " {" + b.Quantity.ToString("N1") + "} ";
                                    }
                                    else if (Math.Abs(b.Quantity) < 10 && Math.Abs(b.Quantity) >= 1)
                                    {
                                        sVal = sVal + " {" + b.Quantity.ToString("N2") + "} ";
                                    }
                                    else if (Math.Abs(b.Quantity) >0)
                                    {
                                        sVal = sVal + " {" + b.Quantity.ToString("N3") + "} ";
                                    }

                                    /*if ((Math.Abs(b.Quantity) <= 1) && (b.Quantity != 0))
                                    {
                                        sVal = sVal + " {" + b.Quantity.ToString("N1") + "} ";
                                    }
                                    else
                                    {
                                        sVal = sVal + " {" + b.Quantity.ToString("N0") + "} ";
                                    }*/
                                }
                            }

                            //Plum
                            var PPData = PlumData.Where(x => x.DateWork.Date == dt.Date).ToList();
                            if  (Delta > 1)    //((PPData.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                            {
                                PPData = PlumData.Where(x => (DateTime)x.AlterDate > dt.Date && (DateTime)x.AlterDate < dtl.Date).ToList();  //=
                            }

                            if (PPData.Count > 0)
                            {
                                foreach (var PP in PPData)
                                {
                                    if (Math.Abs(PP.Quantity) >= 100)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N0") + "> ";
                                    }
                                    else if (Math.Abs(PP.Quantity) < 100 && Math.Abs (PP.Quantity) >= 10)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N1") + "> ";
                                    }
                                    else if (Math.Abs(PP.Quantity) < 10 && Math.Abs(PP.Quantity) >= 1)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N2") + "> ";
                                    }
                                    else if (Math.Abs(PP.Quantity) > 0)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N3") + "> ";
                                    }


                                   /* if ((Math.Abs(PP.Quantity) <= 1) && (PP.Quantity != 0))
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N1") + "> ";
                                    }
                                    else
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N0") + "> ";
                                    }*/
                                }

                                if (!CFMListPlumes.Contains(cfm1.MaterialCode))  //+
                                {
                                    CFMListPlumes.Add(cfm1.MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm1.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListPlum.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListPlum.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
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
                                        CFMListPlum.Add(Col);
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

                    for (int alfa = 0; alfa< DTList.Count; alfa++)
                    {
                        var xlD = xlData.Where(x => x.DateWork.Date == DTList[alfa]).ToList();

                        if (xlD.Count > 0)
                        {
                            double summ = xlD.Sum(x => x.Quantity);

                            if (Math.Abs(summ) >= 100)
                            {
                                sVal = sVal + summ.ToString("N0") + RowSplitter;
                            }
                            else if (Math.Abs(summ) < 100 && Math.Abs(summ) >= 10)
                            {
                                sVal = sVal + summ.ToString("N1") + RowSplitter;
                            }
                            else if (Math.Abs(summ) < 10 && Math.Abs(summ) >= 1)
                            {
                                sVal = sVal + summ.ToString("N2") + RowSplitter;
                            }
                            else if (Math.Abs(summ) > 0)
                            {
                                sVal = sVal + summ.ToString("N3") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + "" + RowSplitter;
                            }

                           /* if ((Math.Abs(summ) <= 1) && (summ != 0))
                            {
                                sVal = sVal + summ.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + summ.ToString("N0") + RowSplitter;
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
                    var PlumData = StorageTaskList.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();

                    var xlData = XLDataList.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
                    var cListData = fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode)).ToList();
                    Cols = 16 + WorkMonthList.Count;  //14
                                                      //   foreach (DateTime dt in DTList)
                    for (int j = 0; j < DTList.Count; j++)
                    {
                        if (cfm1.DontShow == true)
                        {
                            sVal = sVal + "" + RowSplitter;
                        }
                        else
                        {
                            DateTime dt = DTList[j];
                            DateTime dtl;

                            if (j + 1 < DTList.Count)
                            {
                                dtl = DTList[j + 1];
                            }
                            else
                            {
                                dtl = DTList.Last().AddYears(1);
                            }

                            var delta = (dtl.Date - dt.Date).TotalDays;

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
                            List<Raws> tData3 = new List<Raws>();
                            if  (delta > 1)    // ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                            {
                                tData3 = cfm1.TaskNavList2.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();  //=
                            }

                            //  foreach (var td2 in tData2)
                            //  {
                            //      T2 = T2 + td2.Quantity;
                            //  }
                            T2 = T2 + tData2.Sum(x => x.Quantity) + tData3.Sum(x => x.Quantity);

                            List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                            if (cbAddDay.Checked == false)
                            {
                                Cdata = cListData.Where(x => x.ExpDT.Date == dt.Date).ToList();    //fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                            }
                            else
                            {
                                Cdata = cListData.Where(x => x.ExpDT.Date.AddDays(1) == dt.Date).ToList();    //fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                            }

                            List<tPlannedMeatContainers> Cdata1 = new List<tPlannedMeatContainers>();
                            if (delta > 1) //((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                            {
                                Cdata1 = cListData.Where(x => ((DateTime)x.AlterDate).Date > dt.Date && ((DateTime)x.AlterDate).Date < dtl.Date).ToList();  //=   // fm.ContList.Where(x => (x.MaterialCode == cfm1.MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                            }

                            //  foreach (var cd in Cdata)
                            //  {
                            //      T2 = T2 + cd.Quantity;
                            //  }
                            T2 = T2 + Cdata.Sum(x => x.Quantity) + Cdata1.Sum(x => x.Quantity);

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

                            List<Raws> tDataA = new List<Raws>();
                            if  (delta > 1)    //((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                            {
                                tDataA = cfm1.TaskNavList1.Where(x => x.AlterDate > dt && x.AlterDate <dtl).ToList();  //=
                            }

                            T1 = T1 + tData1.Sum(x => x.Quantity) + tDataA.Sum(x => x.Quantity);

                            if (T1 > 0)
                            {
                                if (!AskList.Contains(cfm1.MaterialCode))
                                {
                                    AskList.Add(cfm1.MaterialCode);
                                }
                            }

                            //credit
                            var dll = xlData.Where(x => x.DateWork.Date == dt).ToList();

                            Value = Value - dll.Sum(x => x.Quantity);

                            //future planning Asks
                            double Q = 0;
                            //   var adata = AppList;    // fm.tdb.tApplication.ToList();
                            //  if (adata.Count > 0)
                            //    adata = adata.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();
                            //    if (adata.Count > 0)
                            {
                                var bdata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();
                                List<tApplication> bData1 = new List<tApplication>();

                                if (delta > 1)//((bdata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                {
                                    //  adata = AppList;
                                    //  adata = adata.Where(x => x.MaterialCode == cfm1.MaterialCode).ToList();
                                    bData1 = adata.Where(x => (DateTime)x.AlterDate > dt.Date && (DateTime)x.AlterDate < dtl.Date).ToList();  //=
                                }

                                Q = Q + bdata.Sum(x => x.Quantity) + bData1.Sum(x => x.Quantity);
                                
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
                                    if  (!CFMListGold.Contains(cfm1.MaterialCode))  //-
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

                                        if (cfm1.MaterialStatus > 4)
                                        {
                                            cfm1.MaterialStatus = 4;
                                        } 
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

                                if (Math.Abs(Value) >= 100)
                                {
                                    sVal = sVal + Value.ToString("N0") + " ";
                                }
                                else if (Math.Abs(Value) < 100 & Math.Abs(Value) >= 10)
                                {
                                    sVal = sVal + Value.ToString("N1") + " ";
                                }
                                else if (Math.Abs(Value) < 10 && Math.Abs(Value ) >= 1)
                                {
                                    sVal = sVal + Value.ToString("N2") + " ";
                                }
                                else if (Math.Abs(Value) > 0)
                                {
                                    sVal = sVal + Value.ToString("N3") + " ";
                                }

                                /*if ((Math.Abs(Value) <= 1) && (Value != 0))
                                {
                                    sVal = sVal + Value.ToString("N1") + " ";
                                }
                                else
                                {
                                    sVal = sVal + Value.ToString("N0") + " ";
                                }*/
                            }

                            //Plum
                            var PPData = PlumData.Where(x => x.DateWork.Date == dt.Date).ToList();
                            List<tStorageTask> PPData1 = new List<tStorageTask>();

                            if (delta >1) //((PPData.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                            {
                                PPData1 = PlumData.Where(x => (DateTime)x.AlterDate > dt.Date && (DateTime)x.AlterDate < dtl.Date).ToList();  //=
                            }

                            PPData.AddRange(PPData1);

                            if (PPData.Count > 0)
                            {
                                foreach (var PP in PPData)
                                {
                                    if (Math.Abs(PP.Quantity) >= 100)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N0") + "> ";
                                    }
                                    else if (Math.Abs(PP.Quantity) < 100 && Math.Abs(PP.Quantity)>= 10)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N1") + "> ";
                                    }
                                    else if (Math.Abs(PP.Quantity) < 10 && Math.Abs(PP.Quantity) >= 1)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N2") + "> ";
                                    }
                                    else if (Math.Abs(PP.Quantity) > 0)
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N3") + "> ";
                                    }

                                    /*if ((Math.Abs(PP.Quantity) <= 1) && (PP.Quantity != 0))
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N1") + "> ";
                                    }
                                    else
                                    {
                                        sVal = sVal + " <" + PP.Quantity.ToString("N0") + "> ";
                                    }*/
                                }

                                if (!CFMListPlumes.Contains(cfm1.MaterialCode))  //+
                                {
                                    CFMListPlumes.Add(cfm1.MaterialCode);
                                    Colorres Col = new Colorres
                                    {
                                        ProdCode = cfm1.MaterialCode,
                                        ColNo = new List<int>()
                                    };
                                    Col.ColNo.Add(Cols);
                                    CFMListPlum.Add(Col);
                                }
                                else
                                {
                                    var Coll = CFMListPlum.Where(x => x.ProdCode == cfm1.MaterialCode).ToList();
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
                                        CFMListPlum.Add(Col);
                                    }
                                }
                            }
                        }

                        sVal = sVal + "" + RowSplitter;
                        Cols = Cols + 1;
                    }
                }

                sVal = sVal + cfm1.MaterialStatus.ToString() + RowSplitter;

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
            dataGridViewMain.Columns[dataGridViewMain.ColumnCount - 2].Visible = false;
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
            XLDataList = fm.ExcelDataList ; //.Where(x => x.DateWork >= DateTime.Today.Date).ToList();
            XLDataList = XLDataList.Where(x => x.Quantity > 0).ToList();

            List<lData> DataListTempOld = new List<lData>();
            List<lData> DataListTempNew = new List<lData>();
        //    WorkMonthList = new List<DateTime>();
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
            button1_h.BackColor = Color.Tan;
            button1_i.BackColor = Color.Thistle;
            button1_j.BackColor = Color.Lavender;
            button1_k.BackColor = Color.Aqua;

            dateTimePicker2.Value = DateTime.Today.Date;

            comboBoxFiltering.Items.Clear();
            comboBoxFiltering.Items.Add("Выберите план");

            TBP = fm.tdb.tBlockedMaterial.Where(x => x.Using == false)
                    .Select(x => new ProductRecipes
                    {
                        ProductCode = x.ProductCode,
                        ProductName = x.ProductName
                    })
                    .ToList();

            XLDataListFilter = fm.ExcelDataList;

            foreach (var XL in XLDataListFilter)  //.OrderBy(x => x.DateWork).ThenBy(x => x.Line).ToList()
            {
                comboBoxFiltering.Items.Add(/*XL.RowIndex.ToString() + " " +*/ XL.DateWork.ToString("dd.MM.yyyy") + " " + XL.Line + " " + XL.ProdCodeStr + " " + XL.ProdName + " " + XL.Quantity.ToString("N0") + " кг");
            }

            /*foreach (var XLD in fm.ExcelDataList)  //.OrderBy(x => x.DateWork).ThenBy(x => x.Line).ToList()
            {
                comboBoxFiltering.Items.Add(XLD.DateWork.ToString("dd.MM.yyyy") + " " + XLD.Line + " " + XLD.ProdCodeStr + " " + XLD.AlterProdName + " " + XLD.Quantity.ToString("N0") + " кг");
            }*/

            comboBoxFiltering.Text = "Выберите план";

            cbLineList.Items.Clear();
            cbLineList.Items.Add("Линия");

            List<string> LL = fm.ExcelDataList.Select(x => x.Line).ToList();
            LL = LL.Distinct().ToList();
            LL.Sort();

            for (int i = 0; i < LL.Count; i++)
            {
                cbLineList.Items.Add(LL[i]);
            }

            cbLineList.Text = "Линия";

            Int32 Iteration = 0;
            MaterialNamesList = new List<string>();
            /*   MaterialData = new List<DailyResult>();
               FilteredMaterialData = new List<DailyResult>();*/
            DTList = new List<DateTime>();

          //  XLDataList = XLDataList.Where(x => x.Quantity > 0).ToList();
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
            PRList = fm.ProdRecipeList; //.Where(x => ProductNamesList.Contains(x.ProductCode)).ToList();

         //   var PPR = PRList.Where(x => x.ProductCode == "1010023051").ToList();

            foreach (DateTime dt in DTList)
            {
                var DataList = XLDataList.Where(x => x.DateWork == dt).ToList();

                foreach (var DL in DataList)
                {
                    var Recipe = PRList.Where(x => x.ProductCode.Trim() == DL.ProdCode.Trim()).ToList();
                    if (Recipe.Count > 0)
                    {
                        var RecLine = Recipe[0].RecLineList.Where(x => x.LineNumber.Trim() == DL.Line.Trim()).ToList();
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
                                            LD.ProdCode = RL.MaterialCode.Trim();
                                            LD.ParentCode = DL.ProdCode;
                                            LD.ProdCodeStr = RL.MaterialCode.Trim();
                                            LD.DateWork = dt;
                                            LD.ProdName = RL.MaterialName.Trim();
                                            LD.Quantity = DL.Quantity * RL.MaterialQuantity;
                                            LD.Line = DL.Line.Trim();
                                            LD.IPG = "тара";
                                            LD.Spices = false;
                                            LD.BOM = RL.BOM;
                                            LD.IsSemoProd = RL.IsSemiProd;
                                        }
                                    }
                                    else
                                    {
                                        LD.ProdCode = RL.MaterialCode.Trim();
                                        LD.ProdCodeStr = RL.MaterialCode.Trim();
                                        LD.ParentCode = DL.ProdCode;
                                        LD.DateWork = dt;
                                        LD.ProdName = RL.MaterialName.Trim();
                                        LD.Quantity = DL.Quantity * RL.MaterialQuantity;
                                        LD.Line = DL.Line.Trim();
                                        LD.IPG = RL.IPG;
                                        LD.Spices = DL.Spices;   //control: split 1-st level only!!!!!!!!!
                                        LD.BOM = RL.BOM;
                                        LD.IsSemoProd = RL.IsSemiProd;
                                    }

                                    if (LD.ProdCodeStr != "_____")
                                    {
                                        fl = false;

                                        for (int i = 0; i < DataListTempNew.Count; i++)
                                        {
                                            if (DataListTempNew[i].ProdCode.Trim() == LD.ProdCode.Trim())
                                            {
                                                if (DataListTempNew[i].ParentCode== LD.ParentCode)  //Line
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

                                  /*  if (LD.ProdCode == "1030100136")
                                    { }*/
                                }
                            }
                        }
                    }
                }
            }
            //end 1st step
         //   var d = DataListTempNew.Where(x => x.ProdCode == "3000106479").ToList();

            //10 и больше итераций полуфабрикатов - предложение выхода
            {
                Iteration = 0;
                XLDataList.Clear();

                while (DataListTempNew.Count > 0)
                {
                    Iteration = Iteration + 1;
                    for (int i = DataListTempNew.Count; i > 0; i--)
                    {

                        #region Hidden
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
                         else*/
                        #endregion

                        //all component add into List
                        {
                            fl = false;
                            for (int j = 0; j < XLDataList.Count; j++)
                            {
                                if (XLDataList[j].ProdCode.Trim() == DataListTempNew[i - 1].ProdCode.Trim())
                                {
                                    if (XLDataList[j].DateWork == DataListTempNew[i - 1].DateWork)
                                    {
                                        if (XLDataList[j].ParentCode == DataListTempNew[i - 1].ParentCode)   //Line
                                        {
                                            if (XLDataList[j].Spices == DataListTempNew[i - 1].Spices)
                                            {
                                                 /*if (XLDataList[j].ProdCode == "1021001064")
                                                 { }*/

                                                /*if (XLDataList[j].ProdCode == "1031014059")
                                                { }*/

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
                               /* if (DataListTempNew[i - 1].ProdCode == "1021001064")
                                { }*/

                                /*if (DataListTempNew[i - 1].ProdCode == "1031014059")
                                { }*/

                                LD = new lData
                                {
                                    ProdCode = DataListTempNew[i - 1].ProdCode.Trim(),
                                    ProdName = DataListTempNew[i - 1].ProdName.Trim(),
                                    ParentCode = DataListTempNew[i - 1].ParentCode,
                                    DateWork = DataListTempNew[i - 1].DateWork,
                                    Quantity = DataListTempNew[i - 1].Quantity,
                                    IPG = DataListTempNew[i - 1].IPG,
                                    Line = DataListTempNew[i - 1].Line.Trim(),
                                    Spices = DataListTempNew[i - 1].Spices,
                                    BOM = DataListTempNew[i - 1].BOM,
                                    IsSemoProd = DataListTempNew[i - 1].IsSemoProd
                                };

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

                        if ( DataListTempNew[i - 1].IsSemoProd != 1   )        //IPG.ToLower().Contains(@"п/ф"))         //(DataListTempNew[i - 1].IPG.ToLower() != "полуфабрикат")
                        {
                            DataListTempNew.RemoveAt(i - 1);
                        }
                        else
                        {
                           if (DataListTempNew[i - 1].Quantity == 0)
                            {
                                DataListTempNew.RemoveAt(i - 1);
                            }
                        }
                    }

                    DataListTempOld.Clear();
                    DataListTempOld.AddRange(DataListTempNew);
                    DataListTempNew.Clear();


                    for (int i = 0; i < DataListTempOld.Count; i++)
                    {
   
                        /*if (DataListTempOld[i].ProdCode == "1030100136")
                        {                        }*/
                       // var RRR = fm.ProdRecipeList.Where(x => x.ProductCode == "1021000706").ToList();

                        var RecList = fm.ProdRecipeList.Where(x => x.ProductCode.Trim() == DataListTempOld[i].ProdCode.Trim()).ToList();

                        if (RecList.Count > 0)
                        {
                            var RecLineList = RecList[0].RecLineList; //.Where(x => x.LineNumber.Trim() == DataListTempOld[i].Line.Trim()).ToList();

                            if (RecLineList.Count > 0)
                            {
                                var RecLineDateList = RecLineList[0].RecList.Where(x => (x.WorkDateStart <= DataListTempOld[i].DateWork) && (x.WorkDateEnd >= DataListTempOld[i].DateWork)).ToList();

                                if (RecLineDateList.Count > 0)
                                {
                                    var RList = RecLineDateList[0].rList;
                                    foreach (var RR in RList)
                                    {
                                        /*if (RR.MaterialCode== "1031014059")
                                        { }*/

                                        /*if (RR.MaterialCode == "1021001064")
                                        { }*/

                                        LD = new lData
                                        {
                                            ProdCode = RR.MaterialCode.Trim(),
                                            ProdCodeStr = RR.MaterialCode.Trim(),
                                            DateWork = DataListTempOld[i].DateWork,
                                            ParentCode = DataListTempOld[i].ProdCode,
                                            ProdName = RR.MaterialName.Trim(),
                                            Quantity = RR.MaterialQuantity * DataListTempOld[i].Quantity,
                                            Line =  RecLineList[0].LineNumber.Trim(),    //DataListTempOld[i].Line.Trim(),
                                            IPG = RR.IPG,
                                            BOM = RR.BOM,
                                            IsSemoProd = RR.IsSemiProd
                                        };

                                        fl = false;
                                        for (int j = 0; j < DataListTempNew.Count; j++)
                                        {
                                            if (DataListTempNew[j].ProdCode == LD.ProdCode)
                                            {
                                                if (DataListTempNew[j].ParentCode == LD.ParentCode)   //Line
                                                {
                                                    if (DataListTempNew[j].DateWork == LD.DateWork)
                                                    {
                                                        /*if (XLDataList[j].ProdCode == "1021001064")
                                                        { }*/

                                                        /*if (DataListTempNew[j].ProdCode == "1031014059")
                                                        { }*/

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

                  /*  if (Iteration > 10)
                    {
                        DialogResult dr = MessageBox.Show("Проведено " + Iteration.ToString() + " итераций. Полуфабрикаты не разложены на компоненты!" + System.Environment.NewLine + "Желаете прекратить обработку?", "Сообщение системы", MessageBoxButtons.OKCancel);
                        if (dr == DialogResult.OK)
                        {
                            DataListTempNew.Clear();
                        }
                    }*/
                }//while

            }
            //end iteration
            DataListTempNew.Clear();
            DataListTempOld.Clear();
            //    XLDataList = XLDataList.Where(x => x.IPG.Trim().ToLower() != "гп").ToList();
            MaterialNamesList = XLDataList.Select(x => x.ProdCode).ToList();
          //  XLDataList = XLDataList.Where(x => x.ProdCode == "1031014495").ToList();

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

            AppList = fm.tdb.tApplication.ToList();
            StorageTaskList = fm.tdb.tStorageTask.ToList();

            Int64 RCount = AppList.Count + StorageTaskList.Count;        //fm.tdb.tApplication.ToList().Count+;

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

        public void UpdateStorageTask(Int32 row, Int32 col, double Quant)
        {
            StorageTaskList = fm.tdb.tStorageTask.ToList();
            CFMListPlumes = StorageTaskList.Select(x => x.MaterialCode).Distinct().ToList();
            
            var ut = USCTList.Where(x => (x.RowIndex == row) && (x.ColIndex == col)).ToList();
            double Q = Quant;

            if (ut.Count > 0)  //exist!
            {
                ut[0].Quantity = ut[0].Quantity + Quant;
                string s = dataGridViewMain.Rows[row].Cells[col].Value.ToString();

                if (s.IndexOf("<") > 0)
                {
                    s = s.Substring(0, s.IndexOf("<") - 1);
                }

                s = s + " <" + ut[0].Quantity.ToString("N0") + ">";

                dataGridViewMain.Rows[row].Cells[col].Value = s;
                Q = ut[0].Quantity;
            }
            else  //non exists
            {
                USCTList.Add(new UserSelectedCell
                {
                    RowIndex = row,
                    ColIndex = col,
                    Quantity = Quant
                });


                var CFMPl = CFMListPlum.Where(x => x.ProdCode == UserSelectedCode).ToList();

                if (CFMPl.Count>0)
                {
                    CFMPl[0].ColNo.Add(col);
                }
                else
                {
                    Colorres Coll = new Colorres();

                    Coll.ProdCode = UserSelectedCode;
                    Coll.ColNo = new List<int>();

                    Coll.ColNo.Add(col);
                    CFMListPlum.Add(Coll);
                }

                dataGridViewMain.Rows[row].Cells[col].Value = dataGridViewMain.Rows[row].Cells[col].Value.ToString() + " <" + Quant.ToString("N0") + ">";
                if (dataGridViewMain.Rows[row].Cells[col].Style.BackColor == Color.White)
                {
                    dataGridViewMain.Rows[row].Cells[col].Style.BackColor = Color.Plum;
                }
            }

         //   string Code = dataGridViewMain.Rows[row].Cells[0].Value.ToString();
            var Rec = CFMList.Where(x => x.MaterialCode == UserSelectedCode).FirstOrDefault();

            if (Rec.DontShow == true)
            {
                Int32 C = col - 16 - WorkMonthList.Count;
                if (C >= 0)
                {
                  /*  var ul = USCList.Where(x => (x.RowIndex == row) && (x.ColIndex == col)).ToList();
                    string Append = "";

                    if (ul.Count > 0)
                    {
                        Append = " {" + ul[0].Quantity.ToString("N0") + "}";
                    }*/

                    string V;
                    string vUD;
                    Rec.ValueS = "";
                    Rec.ValueUDS = "";
                    char RowSplitter = '|';

                    if (Math.Abs(Rec.Value_) >= 100)
                    {
                        V = Rec.Value_.ToString("N0");
                    }
                    else if (Math.Abs(Rec.Value_) >= 10 && Math.Abs(Rec.Value_) < 100)
                    {
                        V = Rec.Value_.ToString("N1");
                    }
                    else if (Math.Abs(Rec.Value_) >= 1 && Math.Abs(Rec.Value_) < 10)
                    {
                        V = Rec.Value_.ToString("N2");
                    }
                    else if (Math.Abs(Rec.Value_) > 0)
                    {
                        V = Rec.Value_.ToString("N3");
                    }
                    else
                    {
                        V = "";
                    }

                    /*if ((Math.Abs(Rec.Value_) <= 1) && (Rec.Value_ != 0))
                    {
                        V = Rec.Value_.ToString("N1");
                    }
                    else
                    {
                        V = Rec.Value_.ToString("N0");
                    }*/

                    if (Math.Abs(Rec.ValueUD) >= 100)
                    {
                        vUD = Rec.ValueUD.ToString("N0");
                    }
                    else if (Math.Abs(Rec.ValueUD) >= 10 && Math.Abs(Rec.ValueUD) < 100)
                    {
                        vUD = Rec.ValueUD.ToString("N1");
                    }
                    else if (Math.Abs(Rec.ValueUD) >= 1 && Math.Abs(Rec.ValueUD) < 10)
                    {
                        vUD = Rec.ValueUD.ToString("N2");
                    }
                    else if (Math.Abs(Rec.ValueUD) > 0)
                    {
                        vUD = Rec.ValueUD.ToString("N3");
                    }
                    else
                    {
                        vUD = "";
                    }

                    /*if ((Math.Abs(Rec.ValueUD) <= 1) && (Rec.ValueUD != 0))
                    {
                        vUD = Rec.ValueUD.ToString("N1");
                    }
                    else
                    {
                        vUD = Rec.ValueUD.ToString("N0");
                    }*/

                    //foreach (DateTime dt in DTList)
                    for (int i = 0; i < DTList.Count; i++)
                    {
                        Rec.ValueS = Rec.ValueS + vUD + ""; // + RowSplitter;
                        Rec.ValueUDS = Rec.ValueUDS + V + ""; // + RowSplitter;

                        if (i == C)
                        {
                            Rec.ValueS = Rec.ValueS + " <" + Q.ToString("N0") + "> ";
                            Rec.ValueUDS = Rec.ValueUDS + " <" + Q.ToString("N0") + "> ";
                        }

                        Rec.ValueS = Rec.ValueS + RowSplitter;
                        Rec.ValueUDS = Rec.ValueUDS + RowSplitter;
                    }
                }
            }
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

                    if (Math.Abs(Rec.Value_) >= 100)
                    {
                        V = Rec.Value_.ToString("N0");
                    }
                    else if (Math.Abs(Rec.Value_) >= 10 && Math.Abs(Rec.Value_) < 100)
                    {
                        V = Rec.Value_.ToString("N1");
                    }
                    else if (Math.Abs(Rec.Value_) >= 1 && Math.Abs(Rec.Value_) < 10)
                    {
                        V = Rec.Value_.ToString("N2");
                    }
                    else if (Math.Abs(Rec.Value_) > 0)
                    {
                        V = Rec.Value_.ToString("N3");
                    }
                    else
                    {
                        V = "";
                    }

                    /*if ((Math.Abs(Rec.Value_) <= 1) && (Rec.Value_ != 0))
                    {
                        V = Rec.Value_.ToString("N1");
                    }
                    else
                    {
                        V = Rec.Value_.ToString("N0");
                    }*/

                    if (Math.Abs(Rec.ValueUD) >= 100)
                    {
                        vUD = Rec.ValueUD.ToString("N0");
                    }
                    else if (Math.Abs(Rec.ValueUD) >= 10 && Math.Abs(Rec.ValueUD) < 100)
                    {
                        vUD = Rec.ValueUD.ToString("N1");
                    }
                    else if (Math.Abs(Rec.ValueUD) >= 1 && Math.Abs(Rec.ValueUD) < 10)
                    {
                        vUD = Rec.ValueUD.ToString("N2");
                    }
                    else if (Math.Abs(Rec.ValueUD ) > 0)
                    {
                        vUD = Rec.ValueUD.ToString("N3");
                    }
                    else
                    {
                        vUD = "";
                    }

                    /*if ((Math.Abs(Rec.ValueUD) <= 1) && (Rec.ValueUD != 0))
                    {
                        vUD = Rec.ValueUD.ToString("N1");
                    }
                    else
                    {
                        vUD = Rec.ValueUD.ToString("N0");
                    }*/

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
            if (Application.OpenForms["FormReport"] == null)
            {
                Cursor = Cursors.WaitCursor;
                frt = new FormReport(fm, this );
                frt.Visible = true;
                //   fs.MdiParent = this;
                frt.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                frt.BringToFront();
                frt.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
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
            UpdateFilters();
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
                try
                {
                    UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                    UserSelectedProdName = dataGridViewMain.Rows[e.RowIndex].Cells[1].Value.ToString();
                }
                catch (Exception xx)
                {
                    UserSelectedCode = "";
                    UserSelectedProdName = "";
                }

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
               /* dtpFact1.Value = DateTime.Today.Date.AddMonths(-1);
                dtpFact2.Value = DateTime.Today.Date.AddMonths(1);
                dtpPlan1.Value = DateTime.Today.Date.AddMonths(-1);
                dtpPlan2.Value = DateTime.Today.Date.AddMonths(1);*/

                /*   List<DateTime> DList = new List<DateTime>();
                   DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();
                   DList = DList.Distinct().OrderBy(x => x.Date).ToList();*/

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

                var CurProduct = fm.lRec; //.Where(x => x.MaterialCode == ProductCode && x.IsSemiProd == 1).ToList();
                var CCP = fm.lRec1; //.Where(x => x.MaterialCode == ProductCode && x.IsSemiProd == 1).ToList();

                CurProduct.AddRange(CCP);
                //   DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();

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
                            {
                                if (cbFilter.Checked == false)
                                {
                                    DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();
                                }
                                else
                                {
                                    DDList = DList.Where(x => x.Date >= dtpPlan1.Value.Date && x.Date <= dtpPlan2.Value.Date).OrderBy(x => x.Date).ToList();
                                }

                                UsingComponentPlan(ProductCode, DDList, CurProduct, "", "");
                            }
                            break;
                        case 0:
                            {
                                if (cbFilter.Checked == false)
                                {
                                    DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();
                                }
                                else
                                {
                                    DDList = DList.Where(x => x.Date >= dtpFact1.Value.Date && x.Date <= dtpFact2.Value.Date).OrderBy(x => x.Date).ToList();
                                }

                                var FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, DDList.Min()).ToList();                                
                                UsingComponentFact(ProductCode, DDList, FactData, "", "");
                            }
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
                                List<PR_GetRecipesDataMainProductBis_Result> LResult = new List<PR_GetRecipesDataMainProductBis_Result>();
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
                                            ((double)((decimal)Q/(decimal)lr.Quantity)).ToString("N2"),
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

                    //и для складов
                    {
                        //storages
                        var ListAA = cfm1[0].RawStorageListAlter;

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
                                string DString = aa.ProdDate.ToString("dd.MM.yyyy") + @"/" + aa.BBFDate.ToString("dd.MM.yyyy") + @"/" + aa.TestQuality;
                                string Vstring = "";

                                if (Math.Abs(aa.Quantity)>= 100)
                                {
                                    Vstring = aa.Quantity.ToString("N0");
                                }
                                else if (Math.Abs(aa.Quantity) >= 10 && Math.Abs(aa.Quantity) < 100)
                                {
                                    Vstring = aa.Quantity.ToString("N1");
                                }
                                else if (Math.Abs(aa.Quantity) >= 1 && Math.Abs(aa.Quantity) < 100)
                                {
                                    Vstring = aa.Quantity.ToString("N2");
                                }
                                else if (aa.Quantity != 0)
                                {
                                    Vstring = aa.Quantity.ToString("N3");
                                }

                                dataString = new string[]
                                 {
                                    aa.Storage,
                                    DString,               // aa.LotName,
                                    Vstring,        //aa.Quantity.ToString("N1"),
                                    aa.ProdDate.ToString("dd.MM.yyyy") ,    //ss.ValidFrom.ToString("dd.MM.yyyy"),
                                    aa.BBFDate.ToString("dd.MM.yyyy"),    // ss.ValidTo.ToString("dd.MM.yyyy"),
                                    aa.Color            //ss.Status
                                 };

                                dgvStorages.Rows.Add(dataString);
                            }

                            if (ListAA.Count > 1)
                            {
                                double Q = ListAA.Sum(x => x.Quantity);
                                string Vstring = "";

                                if (Math.Abs(Q) >= 100)
                                {
                                    Vstring = Q.ToString("N0");
                                }
                                else if (Math.Abs(Q) >= 10 && Math.Abs(Q) < 100)
                                {
                                    Vstring = Q.ToString("N1");
                                }
                                else if (Math.Abs(Q) >= 1 && Math.Abs(Q) < 10)
                                {
                                    Vstring = Q.ToString("N2");
                                }
                                else if (Q != 0)
                                {
                                    Vstring = Q.ToString("N3");
                                }

                                dataString = new string[]
                                {
                                "",
                                "Итого остатки:",
                                Vstring,    // ListAA.Sum(x=>x.Quantity).ToString("N1"),
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
                            var EL = cfm1[0].RawEnterpriseListAlter;

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
                                    string DString = el.ProdDate.ToString("dd.MM.yyyy") + @"/" + el.BBFDate.ToString("dd.MM.yyyy") + @"/" + el.TestQuality;
                                    string Vstring = "";

                                    if (Math.Abs(el.Quantity) >= 100)
                                    {
                                        Vstring = el.Quantity.ToString("N0");
                                    }
                                    else if (Math.Abs(el.Quantity) >= 10 && Math.Abs(el.Quantity) < 100)
                                    {
                                        Vstring = el.Quantity.ToString("N1");
                                    }
                                    else if (Math.Abs(el.Quantity) >= 1 && Math.Abs(el.Quantity) < 10)
                                    {
                                        Vstring = el.Quantity.ToString("N2");
                                    }
                                    else if (el.Quantity != 0)
                                    {
                                        Vstring = el.Quantity.ToString("N3");
                                    }

                                    dataString = new string[]
                                     {
                                    el.Storage,
                                    DString,        // el.LotName,
                                    Vstring,        // el.Quantity.ToString("N1"),
                                    el.ProdDate.ToString("dd.MM.yyyy") ,    //ss.ValidFrom.ToString("dd.MM.yyyy"),
                                    el.BBFDate.ToString("dd.MM.yyyy"),    // ss.ValidTo.ToString("dd.MM.yyyy"),
                                    el.Color            //ss.Status
                                     };

                                    dgvStorages.Rows.Add(dataString);
                                }

                                if (EL.Count > 1)
                                {
                                    double Q = EL.Sum(x => x.Quantity);
                                    string VString = "";

                                    if (Math.Abs(Q) >= 100)
                                    {
                                        VString = Q.ToString("N0");
                                    }
                                    else if (Math.Abs(Q) >= 10 && Math.Abs(Q) < 100)
                                    {
                                        VString = Q.ToString("N1");
                                    }
                                    else if (Math.Abs(Q) >= 1 && Math.Abs(Q) < 10)
                                    {
                                        VString = Q.ToString("N2");
                                    }
                                    else if (Q != 0)
                                    {
                                        VString = Q.ToString("N3");
                                    }

                                    dataString = new string[]
                                    {
                                "",
                                "Итого остатки:",
                                 VString,           //EL.Sum(x=>x.Quantity).ToString("N1"),
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

                        var sListBB = sListB.Where(x => /*(x.PlanOperDate.Date >= DateTime.Today.Date) && */(x.StatusMZP.Trim().ToLower() == "заказано" || x.StatusMZP.Trim().ToLower().Contains("в пути"))).ToList();

                 

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
                                string Vstring = "";

                                if (Math.Abs(ed.Quantity) >= 100)
                                {
                                    Vstring = ed.Quantity.ToString("N0");
                                }
                                else if (Math.Abs(ed.Quantity) >= 10 && Math.Abs(ed.Quantity) < 100)
                                {
                                    Vstring = ed.Quantity.ToString("N1");
                                }
                                else if (Math.Abs(ed.Quantity) >= 1 && Math.Abs(ed.Quantity) < 10)
                                {
                                    Vstring = ed.Quantity.ToString("N2");
                                }
                                else if (ed.Quantity != 0)
                                {
                                    Vstring = ed.Quantity.ToString("N3");
                                }

                                LName = Vstring + " => " + ed.Line;   //ed.Quantity.ToString("N1") 
                                if (ed.RePack == true)
                                {
                                    LName = LName + " П/УП";
                                }

                                contextMenuStripData.Items.Add(LName);
                            }
                        }

                        if (ExData.Count > 1)
                        {
                            string Vstring = "";
                            double Q = ExData.Sum(x => x.Quantity);

                            if (Math.Abs(Q) >= 100)
                            {
                                Vstring = Q.ToString("N0");
                            }
                            else if (Math.Abs(Q) >= 10 && Math.Abs(Q) < 100)
                            {
                                Vstring = Q.ToString("N1");
                            }
                            else if (Math.Abs(Q) >= 1 && Math.Abs(Q) < 10)
                            {
                                Vstring = Q.ToString("N2");
                            }
                            else if (Q != 0)
                            {
                                Vstring = Q.ToString("N3");
                            }

                            LName = "__________";
                            contextMenuStripData.Items.Add(LName);
                            LName = "Итого: " + Vstring;            //(ExData.Sum(x => x.Quantity)).ToString("N1");
                            contextMenuStripData.Items.Add(LName);
                        }
                    }
                    else
                    {
                        switch (e.ColumnIndex - Offset)
                        {
                            case 8:
                                if (cfm1.Count > 0)
                                {
                                    foreach (var cl in cfm1[0].ConsurmptionList)
                                    {
                                        string Vstring = "";

                                        if (Math.Abs( cl.Quantity) >= 100)
                                        {
                                            Vstring = cl.Quantity.ToString("N0");
                                        }
                                        else if (Math.Abs(cl.Quantity) >= 10 && Math.Abs(cl.Quantity) < 100)
                                        {
                                            Vstring = cl.Quantity.ToString("N1");
                                        }
                                        else if (Math.Abs(cl.Quantity) >= 1 && Math.Abs(cl.Quantity) < 10)
                                        {
                                            Vstring = cl.Quantity.ToString("N2");
                                        }
                                        else if (cl.Quantity != 0)
                                        {
                                            Vstring = cl.Quantity.ToString("N3");
                                        }

                                        m1 = new ToolStripMenuItem(cl.Storage + " => " + Vstring);          //cl.Quantity.ToString("N1"));
                                        contextMenuStripData.Items.Add(m1);
                                    }
                                }
                                break;

                            case 9:
                                {
                                    var ListAA = cfm1[0].RawStorageListAlter;

                                    if (cbMaterialDelay.Checked == true)
                                    {
                                        ListAA = ListAA.Where(x => x.BBFDate >= DateTime.Today.Date).ToList();
                                    }

                                    foreach (var sml in ListAA.OrderBy(x => x.ProdDate)) //sbrMesList)
                                    {
                                        string Vstring = "";

                                        if (Math.Abs(sml.Quantity) >= 100)
                                        {
                                            Vstring = sml.Quantity.ToString("N0");
                                        }
                                        else if (Math.Abs(sml.Quantity) >= 10 && Math.Abs(sml.Quantity) < 100)
                                        {
                                            Vstring = sml.Quantity.ToString("N1");
                                        }
                                        else if (Math.Abs(sml.Quantity) >= 1 && Math.Abs(sml.Quantity) < 10)
                                        {
                                            Vstring = sml.Quantity.ToString("N2");
                                        }
                                        else if (sml.Quantity != 0)
                                        {
                                            Vstring = sml.Quantity.ToString("N3");
                                        }

                                            m1 = new ToolStripMenuItem(sml.Storage + " [" + sml.LotName + @"/" + sml.LotDescr + @"/" +Vstring + "] => " + sml.Quantity.ToString("N1"));  //sml.TestQuality 
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
                                break;

                            case 10:
                                {
                                    var EL = cfm1[0].RawEnterpriseListAlter;

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
                                        string Vstring = "";

                                        if (Math.Abs(el.Quantity) >= 100)
                                        {
                                            Vstring = el.Quantity.ToString("N0");
                                        }
                                        else if (Math.Abs(el.Quantity) >= 10 && Math.Abs(el.Quantity) < 100)
                                        {
                                            Vstring = el.Quantity.ToString("N1");
                                        }
                                        else if (Math.Abs(el.Quantity) >= 1 && Math.Abs(el.Quantity) < 10)
                                        {
                                            Vstring = el.Quantity.ToString("N2");
                                        }
                                        else if (el.Quantity != 0)
                                        {
                                            Vstring = el.Quantity.ToString("N3");
                                        }

                                        ToolStripMenuItem m2 = new ToolStripMenuItem(el.Storage + " [" + el.LotName + @"/" + el.LotDescr + @"/" + Vstring + "] => " + el.Quantity.ToString("N1"));  //el.TestQuality
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
                                break;

                            case 14:
                                {
                                    var ListAA = cfm1[0].RawMesOut2List;

                                    if (cbMaterialDelay.Checked == true)
                                    {
                                        ListAA = ListAA.Where(x => x.BBFDate >= DateTime.Today.Date).ToList();
                                    }

                                    foreach (var aa in ListAA)
                                    {
                                        string Vstring = "";

                                        if (Math.Abs(aa.Quantity) >= 100)
                                        {
                                            Vstring = aa.Quantity.ToString("N0");
                                        }
                                        else if (Math.Abs(aa.Quantity) >= 10 && Math.Abs(aa.Quantity) < 100)
                                        {
                                            Vstring = aa.Quantity.ToString("N1");
                                        }
                                        else if (Math.Abs(aa.Quantity) >= 1 && Math.Abs(aa.Quantity) < 10)
                                        {
                                            Vstring = aa.Quantity.ToString("N2");
                                        }
                                        else if (aa.Quantity != 0)
                                        {
                                            Vstring = aa.Quantity.ToString("N3");
                                        }

                                        ToolStripMenuItem m2 = new ToolStripMenuItem(aa.Storage + " [" + aa.LotName + @"/" + aa.LotDescr + @"/" + aa.StatusMZP + "] => " + Vstring);       // + aa.Quantity.ToString("N1"));
                                        //      m2.ToolTipText = "Количество записей: " + sl1.Count.ToString();
                                        m2.ToolTipText = "Выпущен: " + aa.ProdDate.ToString("dd.MM.yyyy") + "; Годен до: " + aa.BBFDate.ToString("dd.MM.yyyy");
                                        contextMenuStripData.Items.Add(m2);
                                    }
                                }
                                break;

                            case 15:
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
                                            string Vstring = "";

                                            if (Math.Abs(c.PlanQuantity) >= 100)
                                            {
                                                Vstring = c.PlanQuantity.ToString("N0");
                                            }
                                            else if (Math.Abs(c.PlanQuantity) >= 10 && Math.Abs(c.PlanQuantity) < 100)
                                            {
                                                Vstring = c.PlanQuantity.ToString("N1");
                                            }
                                            else if (Math.Abs(c.PlanQuantity) >= 1 && Math.Abs(c.PlanQuantity) < 10)
                                            {
                                                Vstring = c.PlanQuantity.ToString("N2");
                                            }
                                            else if (c.PlanQuantity != 0)
                                            {
                                                Vstring = c.PlanQuantity.ToString("N3");
                                            }

                                            m1 = new ToolStripMenuItem(c.OrderNymber + "/" + c.LostDescr + " => " + Vstring);    // c.PlanQuantity.ToString("N0"));
                                            m1.ToolTipText = "Плановая дата поставки: " + c.PlanOperDate.ToString("dd.MM.yyyy") + "; К-во: " + c.Count;
                                            m1.ForeColor = Color.DarkRed;

                                            contextMenuStripData.Items.Add(m1);
                                        }
                                    }
                                }
                                break;

                            default:
                                break;
                        }

                         /* else if (e.ColumnIndex == (8 + Offset))
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
                            var ListAA = cfm1[0].RawStorageListAlter;

                            if (cbMaterialDelay.Checked == true)
                            {
                                ListAA = ListAA.Where(x => x.BBFDate >= DateTime.Today.Date).ToList();
                            }

                            foreach (var sml in ListAA.OrderBy(x => x.ProdDate)) //sbrMesList)
                            {
                                m1 = new ToolStripMenuItem(sml.Storage + " [" + sml.LotName + @"/" + sml.LotDescr + @"/" + sml.TestQuality + "] => " + sml.Quantity.ToString("N1"));
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
                        else if (e.ColumnIndex == (10 + Offset))   //остатки пр-во
                        {
                            if (cfm1.Count > 0)
                            {
                                var EL = cfm1[0].RawEnterpriseListAlter;

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
                                    ToolStripMenuItem m2 = new ToolStripMenuItem(el.Storage + " [" + el.LotName + @"/" + el.LotDescr + @"/" + el.TestQuality + "] => " + el.Quantity.ToString("N1"));
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
                        else if (e.ColumnIndex == (14 + Offset))  //Blocked Mes
                        {
                            var ListAA = cfm1[0].RawMesOut2List;

                            foreach (var aa in ListAA)
                            {
                                ToolStripMenuItem m2 = new ToolStripMenuItem(aa.Storage + " [" + aa.LotName + @"/" + aa.LotDescr + @"/" + aa.StatusMZP + "] => " + aa.Quantity.ToString("N1"));
                                //      m2.ToolTipText = "Количество записей: " + sl1.Count.ToString();
                                m2.ToolTipText = "Выпущен: " + aa.ProdDate.ToString("dd.MM.yyyy") + "; Годен до: " + aa.BBFDate.ToString("dd.MM.yyyy");
                                contextMenuStripData.Items.Add(m2);
                            }
                        }
                        else if (e.ColumnIndex == (15 + Offset))  //old Task Nav
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
                        }*/
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
                    try
                    {
                        UserSelectedCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                        UserSelectedProdName = dataGridViewMain.Rows[e.RowIndex].Cells[1].Value.ToString();
                    }
                    catch (Exception xx)
                    {
                        UserSelectedCode = "";
                        UserSelectedProdName = "";
                    }


                    if ((e.ColumnIndex == 0) || (e.ColumnIndex == 1))//color selection
                    {
                       
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
                                    string SelCode = dataGridViewMain.Rows[e.RowIndex].Cells[0].Value.ToString();
                                    var Data = fm.SMList.Where(x => x.MaterialCode == SelCode && x.BOM == "ВЫВОД").ToList();

                                    if (Data.Count > 0)
                                    {
                                        MessageBox.Show("Данный материал имеест статус ВЫВОД."+System.Environment.NewLine+"Оформление заявки невозможно!", "Сообщение системы");
                                    }
                                    else
                                    {
                                        USC = new UserSelectedCell
                                        {
                                            UserSelectedCode = SelCode,
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
                var ListAA = cfm1[0].RawStorageListAlter;

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
                        string DString = aa.ProdDate.ToString("dd.MM.yyyy") + @"/" + aa.BBFDate.ToString("dd.MM.yyyy") + @"/" + aa.TestQuality;
                        string Vstring = "";

                        if (Math.Abs(aa.Quantity) >= 100)
                        {
                            Vstring = aa.Quantity.ToString("N0");
                        }
                        else if (Math.Abs(aa.Quantity) >= 10 && Math.Abs(aa.Quantity) < 100)
                        {
                            Vstring = aa.Quantity.ToString("N1");
                        }
                        else if (Math.Abs(aa.Quantity) >= 1 && Math.Abs(aa.Quantity) < 10)
                        {
                            Vstring = aa.Quantity.ToString("N2");
                        }
                        else if (aa.Quantity != 0)
                        {
                            Vstring = aa.Quantity.ToString("N3");
                        }

                        dataString = new string[]
                         {
                                    aa.Storage,
                                    DString,        //aa.LotName,
                                    Vstring,        //aa.Quantity.ToString("N1"),
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
                    var EL = cfm1[0].RawEnterpriseListAlter;

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
                            string DString = el.ProdDate.ToString("dd.MM.yyyy") + @"/" + el.BBFDate.ToString("dd.MM.yyyy") + @"/" + el.TestQuality;
                            string Vstring = "";

                            if (Math.Abs(el.Quantity) >= 100)
                            {
                                Vstring = el.Quantity.ToString("N0");
                            }
                            else if (Math.Abs(el.Quantity) >= 10 && Math.Abs(el.Quantity) < 100)
                            {
                                Vstring = el.Quantity.ToString("N1");
                            }
                            else if (Math.Abs(el.Quantity) >= 1 && Math.Abs(el.Quantity) < 10)
                            {
                                Vstring = el.Quantity.ToString("N2");
                            }
                            else if (el.Quantity != 0)
                            {
                                Vstring = el.Quantity.ToString("N3");
                            }

                            dataString = new string[]
                             {
                                    el.Storage,
                                    DString,        //el.LotName,
                                    Vstring,        // el.Quantity.ToString("N1"),
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

                /*List<DateTime> DList = new List<DateTime>();
                DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();
                DList = DList.Distinct().OrderBy(x => x.Date).ToList();*/

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
                        var ELDListDate = fm.ExcelDataList.Where(x => (x.DateWork == dt1) && (x.Quantity != 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //  CurrDate

                        foreach (var ELD in ELDListDate)
                        {
                            /* var CP = CurProduct.Where(x => (x.ProductCode == ELD.ProdCode) && (x.LineNumber == ELD.Line)).ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
                             CP = CP.Where(x => x.WorkDate < dt1).OrderByDescending(x => x.WorkDate).ToList();  //Starting_Date

                             if (CP.Count > 0)*/

                            var CP = CurProduct.Where(x => (x.ProductCode == ELD.ProdCode) && (x.LineNumber == ELD.Line) && (x.WorkDate <= dt1)); //.ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
                            PR_GetRecipesDataMainProductBis_Result SRR;
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
            if (cbShortReport.Checked == true)
            {
                Cursor = Cursors.WaitCursor;

                int Interval = 10;
                var pt = fm.PTList.Where(x => x.PropertyName == "DayIntervalToSelect").ToList();

                if (pt.Count >0)
                {
                    Interval = Convert.ToInt32(pt[0].PropertyValue);
                }

                DateTime DStart = DateTime.Today.Date;
                DateTime DEnd = DStart.AddDays(Interval);
                DateTime dd = DStart;

               // var DTL = DTList.Where(x => x.Date >= DStart && x.Date <= DEnd).ToList();
                List<tMaterial> MatList = new List<tMaterial>();
                List<string> MatCode = new List<string>();

                while (dd <=DEnd)
                {
                    var XLD = XLDataList.Where(x => x.DateWork == dd).ToList();

                    for (int i =0; i<XLD.Count; i++)
                    {
                        if (XLD[i].Quantity >0)
                        {
                            MatCode.Add(XLD[i].ProdCode);
                        }
                    }

                    dd = dd.AddDays(1);
                }

                MatCode = MatCode.Distinct().ToList();

                if (MatCode.Count > 0)
                {
                    Excel1.Application xlApp = new Excel1.Application();
                    xlApp.Visible = true;
                    object misValue = System.Reflection.Missing.Value;

                    Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                    xlApp.DisplayAlerts = false;
                    Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                    wSheet.Name = "Потр-ние";
                    Excel1.Range range1;

                    range1 = (Excel1.Range)(wSheet.Cells[1, 1]);
                    range1.ColumnWidth = 15;
                    range1.EntireColumn.NumberFormat = "@";
                    wSheet.Cells[1, 1] = "Код материала";

                    range1 = (Excel1.Range)(wSheet.Cells[1, 2]);
                    range1.ColumnWidth = 50;
                    wSheet.Cells[1, 2] = "Наименование";

                    range1 = (Excel1.Range)(wSheet.Cells[1, 3]);
                    range1.ColumnWidth = 15;
                    wSheet.Cells[1, 3] = "Тип материала";
                    wSheet.Cells[1, 4] = "Ед.изм.";

                    dd = DStart;

                    for (int i = 5; i < 5 + Interval; i++)
                    {
                        range1 = (Excel1.Range)(wSheet.Cells[1, i]);
                        range1.ColumnWidth = 15;
                        range1.EntireColumn.NumberFormat = "### ### ##0.0";
                        wSheet.Cells[1, i] = "Потребление на " + dd.ToString("yyyy-MM-dd");
                        dd = dd.AddDays(1);
                    }

                    int CurrRec = 2;

                    for (int i = 0; i<MatCode.Count; i++)
                    {
                        string MC = MatCode[i];
                        var cfm = CFMList.Where(x => x.MaterialCode.Trim() == MC).ToList();

                        if (cfm.Count > 0)
                        {
                            wSheet.Cells[CurrRec, 1] = cfm[0].MaterialCode;
                            wSheet.Cells[CurrRec, 2] = cfm[0].MaterialName;
                            wSheet.Cells[CurrRec, 3] = cfm[0].MaterialGroup;

                            var sp = fm.SDataList.Where(x => x.AlterMaterialCode.Trim() == MC).ToList();

                            if (sp.Count >0)
                            {
                                wSheet.Cells[CurrRec, 4] = sp[0].UnitMeasure;
                            }

                            dd = DStart;

                            for (int j = 5; j < 5 + Interval; j++)
                            {
                                //wSheet.Cells[CurrRec, j] = 
                                var XLD = XLDataList.Where(x => x.DateWork.Date == dd && x.ProdCode == MC).ToList();

                                if (XLD.Count > 0)
                                {
                                    /*if (XLD[0].Quantity > 0)
                                    {*/
                                    wSheet.Cells[CurrRec, j] = XLD.Sum(x => x.Quantity);
                                   // }
                                }

                                dd = dd.AddDays(1);
                            }

                            CurrRec = CurrRec + 1;
                        }
                    }

                    range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[CurrRec - 1, 5 + Interval]];
                    range1.WrapText = true;
                    fm.SetBorders(range1, 1, true);

                    wSheet = null;
                    wBook = null;
                    range1 = null;
                    //   xlApp.Quit();
                    xlApp = null;

                    Cursor = Cursors.Default;
                    MessageBox.Show("Отчет успешно сформирован", "Сообщение системы");
                }
                else
                {
                    Cursor = Cursors.Default;
                    MessageBox.Show("Данные за ближайшие "+Interval.ToString()+" дней отсутствуют!", "Сообщение системы");
                }
            }
            else
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

                        foreach (var sl in CFMList[i].RawStorageListAlter)
                        {
                            Value = Value + sl.Quantity;
                        }

                        /* foreach (var el in CFMList[i].RawEnterpriseList)
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

                        for (int j = 0; j < DTList.Count; j++)
                        {
                            DateTime dt = DTList[j];
                            DateTime dtl;

                            if (j + 1 < DTList.Count)
                            {
                                dtl = DTList[j + 1];
                            }
                            else
                            {
                                dtl = DTList.Last().AddYears(1);
                            }

                            var delta = (dtl - dt).TotalDays;
                            T2 = 0;
                            var FDTList = DTList.Where(x => (x.Month == DTList[j].Month) && (x.Year == DTList[j].Year)).ToList();
                            //income
                            var tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == DTList[j].Date).ToList();

                            List<Raws> tData3 = new List<Raws>();
                            if (delta > 1)   //((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                            {
                                tData3 = CFMList[i].TaskNavList2.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();    // x.AlterDate == DTList[j]
                            }

                            // foreach (var td2 in tData2)
                            // {
                            T2 = T2 + tData2.Sum(x => x.Quantity) + tData3.Sum(x => x.Quantity);
                            //}

                            var Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == DTList[j].Date)).ToList();
                            List<tPlannedMeatContainers> Cdata1 = new List<tPlannedMeatContainers>();

                            if (delta > 1)    //((DTList[j].Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                            {
                                Cdata1 = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date > dt.Date && ((DateTime)x.AlterDate).Date < dtl.Date)).ToList();
                            }

                            // foreach (var cd in Cdata)
                            // {
                            T2 = T2 + Cdata.Sum(x => x.Quantity) + Cdata1.Sum(x => x.Quantity);
                            // }

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

                            List<Raws> tDataA = new List<Raws>();
                            if (delta > 1)      //((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                            {
                                tData1 = CFMList[i].TaskNavList1.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();
                            }

                            T1 = 0;
                            // foreach (var td1 in tData1)
                            //{
                            T1 = T1 + tData1.Sum(x => x.Quantity) + tDataA.Sum(x => x.Quantity);
                            // }
                            //   Value = Value + T1;
                            //credit
                            var dll = XLDataList.Where(x => (x.ProdCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();
                            foreach (var dl in dll)
                            {
                                Value = Value - dl.Quantity;
                            }

                            Q = 0;

                            var adata = AppList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();               // AppList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.DateWork.Date == DTList[j].Date)).ToList();

                            List<tApplication> adata1 = new List<tApplication>();

                            if (delta > 1)       //((adata.Count == 0) && (FDTList.Count == 1) && (DTList[j].Day == 1))
                            {
                                //   adata = AppList;
                                adata1 = AppList.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();          //adata.Where(x => x.MaterialCode == CFMList[i].MaterialCode).ToList();
                                adata1 = adata.Where(x => (DateTime)x.AlterDate > dt && (DateTime)x.AlterDate < dtl).ToList();
                            }

                            //  foreach (var ad in adata)
                            //  {
                            Q = Q + adata.Sum(x => x.Quantity) + adata1.Sum(x => x.Quantity);
                            //  }

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

                        foreach (var sl in CFMList[i].RawStorageListAlter)
                        {
                            Value = Value + sl.Quantity;
                        }

                        /*  foreach (var el in CFMList[i].RawEnterpriseListAlter)
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

                        foreach (var sl in CFMList[i].RawStorageListAlter)
                        {
                            Value = Value + sl.Quantity;
                        }

                        /*   foreach (var el in CFMList[i].RawEnterpriseListAlter)
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

                        for (int j = 0; j < DTList.Count; j++)
                        {
                            T2 = 0;
                            var FDTList = DTList.Where(x => (x.Month == DTList[j].Month) && (x.Year == DTList[j].Year)).ToList();

                            DateTime dt = DTList[j];
                            DateTime dtl;

                            if (j + 1 < DTList.Count)
                            {
                                dtl = DTList[j + 1];
                            }
                            else
                            {
                                dtl = DTList.Last().AddYears(1);
                            }

                            var Delta = (dtl - dt).TotalDays;

                            //income
                            var tData2 = CFMList[i].TaskNavList2.Where(x => x.BBFDate == DTList[j].Date).ToList();
                            List<Raws> tData1 = new List<Raws>();


                            if (Delta > 1)    //((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                            {
                                tData1 = CFMList[i].TaskNavList2.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();
                            }

                            //  foreach (var td2 in tData2)
                            //  {
                            T2 = T2 + tData1.Sum(x => x.Quantity) + tData2.Sum(x => x.Quantity);
                            //  }

                            var Cdata = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (x.ExpDT.Date == DTList[j].Date)).ToList();
                            List<tPlannedMeatContainers> Cdata1 = new List<tPlannedMeatContainers>();

                            if (Delta > 1)    //((DTList[j].Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                            {
                                Cdata1 = fm.ContList.Where(x => (x.MaterialCode == CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date > dt.Date && ((DateTime)x.AlterDate).Date < dtl)).ToList();
                            }

                            //  foreach (var cd in Cdata)
                            //  {
                            T2 = T2 + Cdata.Sum(x => x.Quantity) + Cdata1.Sum(x => x.Quantity);
                            //  }

                            if (T2 > 0)
                            {
                                if (Value < 0)
                                {
                                    Value = 0;
                                }
                            }
                            Value = Value + T2;

                            //ask in progress
                            var tDataA = CFMList[i].TaskNavList1.Where(x => x.BBFDate == DTList[j].Date).ToList();
                            if (Delta > 1)    //((DTList[j].Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                            {
                                tDataA = CFMList[i].TaskNavList1.Where(x => x.AlterDate > dt && x.AlterDate < dtl).ToList();
                            }

                            T1 = 0;
                            //   foreach (var td1 in tDataA)
                            //  {
                            T1 = T1 + tDataA.Sum(x => x.Quantity);
                            //  }
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

                    Cursor = Cursors.Default;
                    MessageBox.Show("Отчет сформирован", "Сообщение системы");
                }
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();          
        }

        private void cbUseMaterialPlanning_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFilters();
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
         //   if (cbPlanFilter.Checked == true)
            {
              //  ClearDGV();
                List<string> OprComp = fm.OPRTableList.Select(x => x.MatCode).ToList();

                Cursor = Cursors.WaitCursor;
                if ((comboBoxFiltering.Text.Trim() != "" || comboBoxFiltering.Text.Trim() != "Выберите план") && comboBoxFiltering.SelectedIndex > 0)
                {
                    var XLF = XLDataListFilter[comboBoxFiltering.SelectedIndex - 1];
                    int RowInd = XLF.RowIndex;

                    var XLD = fm.ExcelDataList.Where(x => x.RowIndex == RowInd).FirstOrDefault();

                    //var XLD = fm.ExcelDataList[comboBoxFiltering.SelectedIndex];
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
                                            LD.ParentCode = XLD.ProdCode;
                                            LD.ProdCodeStr = RL.MaterialCode;
                                            LD.DateWork = XLD.DateWork;
                                            LD.ProdName = RL.MaterialName;
                                            LD.Quantity = XLD.Quantity * RL.MaterialQuantity;
                                            LD.Line = XLD.Line;
                                            LD.IPG = "тара";
                                            LD.Spices = false;
                                            LD.BOM = RL.BOM;
                                            LD.IsSemoProd = RL.IsSemiProd;
                                        }
                                    }
                                    else
                                    {
                                        LD.ProdCode = RL.MaterialCode;
                                        LD.ProdCodeStr = RL.MaterialCode;
                                        LD.ParentCode = XLD.ProdCode;
                                        LD.DateWork = XLD.DateWork;
                                        LD.ProdName = RL.MaterialName;
                                        LD.Quantity = XLD.Quantity * RL.MaterialQuantity;
                                        LD.Line = XLD.Line;
                                        LD.IPG = RL.IPG;
                                        LD.Spices = XLD.Spices;
                                        LD.BOM = RL.BOM;
                                        LD.IsSemoProd = RL.IsSemiProd;
                                    }

                                    if (LD.ProdCodeStr != "_____")
                                    {
                                        fl = false;

                                        for (int i = 0; i < DataListTempNew.Count; i++)
                                        {
                                            if (DataListTempNew[i].ProdCode == LD.ProdCode)
                                            {
                                                if (DataListTempNew[i].ParentCode == LD.ParentCode)  //Line
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
                                {
                                    fl = false;
                                    for (int j = 0; j < LocalList.Count; j++)
                                    {
                                        if (LocalList[j].ProdCode == DataListTempNew[i - 1].ProdCode)
                                        {
                                            if (LocalList[j].DateWork == DataListTempNew[i - 1].DateWork)
                                            {
                                                if (LocalList[j].ParentCode == DataListTempNew[i - 1].ParentCode)   //.Line
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
                                            ParentCode = DataListTempNew[i - 1].ParentCode,
                                            DateWork = DataListTempNew[i - 1].DateWork,
                                            Quantity = DataListTempNew[i - 1].Quantity,
                                            IPG = DataListTempNew[i - 1].IPG,
                                            Line = DataListTempNew[i - 1].Line,
                                            Spices = DataListTempNew[i - 1].Spices,
                                            BOM = DataListTempNew[i - 1].BOM,
                                            IsSemoProd = DataListTempNew[i - 1].IsSemoProd
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

                                if (DataListTempNew[i - 1].IsSemoProd != 1)        //IPG.ToLower().Contains(@"п/ф"))         //(DataListTempNew[i - 1].IPG.ToLower() != "полуфабрикат")
                                {
                                    DataListTempNew.RemoveAt(i - 1);
                                }
                                else
                                {
                                    if (DataListTempNew[i - 1].Quantity == 0)
                                     {
                                         DataListTempNew.RemoveAt(i - 1);
                                     }
                                }
                            }

                            DataListTempOld.Clear();
                            DataListTempOld.AddRange(DataListTempNew);
                            DataListTempNew.Clear();

                            for (int i = 0; i < DataListTempOld.Count; i++)
                            {
                                var RecList = fm.ProdRecipeList.Where(x => x.ProductCode == DataListTempOld[i].ProdCode).ToList();

                                if (RecList.Count > 0)
                                {
                                    var RecLineList = RecList[0].RecLineList; //.Where(x => x.LineNumber == DataListTempOld[i].Line).ToList();

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
                                                    ParentCode = DataListTempOld[i].ProdCode,
                                                    DateWork = DataListTempOld[i].DateWork,
                                                    ProdName = RR.MaterialName,
                                                    Quantity = RR.MaterialQuantity * DataListTempOld[i].Quantity,
                                                    Line = Line = RecLineList[0].LineNumber.Trim(),                                            //        DataListTempOld[i].Line,       
                                                    IPG = RR.IPG,
                                                    BOM = RR.BOM,
                                                    IsSemoProd = RR.IsSemiProd
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

                            /*if (Iteration > 10)
                            {
                                DialogResult dr = MessageBox.Show("Проведено " + Iteration.ToString() + " итераций. Полуфабрикаты не разложены на компоненты!" + System.Environment.NewLine + "Желаете прекратить обработку?", "Сообщение системы", MessageBoxButtons.OKCancel);
                                if (dr == DialogResult.OK)
                                {
                                    DataListTempNew.Clear();
                                }
                            }*/
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

                comboBoxFiltering.Text = "Выберите план";
                Cursor = Cursors.Default;
            }
        }

        private void cbMaterialDelay_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFilters();
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
            UpdateFilters();
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
            int Selector = dataGridViewMain.ColumnCount - 2;
            //Font f1 = new Font(dataGridViewMain.DefaultCellStyle.Font, FontStyle.Bold);

            for (int j = 0; j < dataGridViewMain.Rows.Count; j++)  //     CFMList
            {
               // bool ttn = true;
                // if ((j >= LowPageIndex) && (j <= HighPageIndex))
                // {
                string PC = dataGridViewMain.Rows[j].Cells[0].Value.ToString();   //dataGridViewMain.Rows[j].Cells[0].Value.ToString();   CFMList[j].MaterialCode
                string MatStat = dataGridViewMain.Rows[j].Cells[Selector].Value.ToString();
                var cfm = CFMList.Where(x => x.MaterialCode == PC).First();

                switch (MatStat)
                {
                    case ("-1"):
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.Aqua;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.Aqua;
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "0":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.Lavender;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.Lavender;
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "1":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = SColor;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = SColor;                         
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "2":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = MColor;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = MColor;                        
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "3":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightCyan; //  .SteelBlue;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightCyan;                        
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "4":
                        if (cfm.MaterialGroup != "ГП") //(CFMBool[j] == false) 
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightYellow;   //.Gold;
                                dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightYellow;   //.Gold;                             
                            }
                            catch (Exception xx)
                            { }
                        }
                        break;

                    case "5":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.Tan; //  
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.Tan;
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "6":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGreen;   //.Lime;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGreen;
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "7":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.MistyRose;   //.Rose;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.MistyRose;
                        }
                        catch (Exception xx)
                        { }
                        break;

                    case "8":
                        if (cfm.MaterialGroup != "ГП")
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGray;  //DarkGray
                                dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGray;
                            }
                            catch (Exception xx)
                            { }
                        }
                        break;

                    case "9":
                        try
                        {
                            dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.Thistle;
                            dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.Thistle;
                        }
                        catch (Exception xx)
                        { }
                        break;

                    default:
                        break;
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

                if (CFMListPlumes.Contains(PC))
                {
                    var CMP = CFMListPlum.Where(x => x.ProdCode == PC).ToList();
                    if (CMP.Count>0)
                    {
                        foreach(var c in CMP[0].ColNo)
                        {
                            try
                            {
                                dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.Plum;   //.Violet
                            }
                            catch (Exception xx)
                            { }
                        }
                    }
                }

                //9/10 color
                try
                {
                   // var cfm = CFMList.Where(x => x.MaterialCode == PC).First();
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
              
            }

            Cursor = Cursors.Default;
        }

        private void tabControl2_SelectedIndexChanged(object sender, EventArgs e)
        {
            string[] dataString;
            bool fl;

            Cursor = Cursors.WaitCursor;

            if (RowInd == null)
            {
                RowInd = 0;
            }

            if (RowInd > 0)
            {
                /*List<DateTime> DList = new List<DateTime>();
                DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();
                DList = DList.Distinct().OrderBy(x => x.Date).ToList();*/

                List<fn_select_ProductionByCode_Result> FactData = new List<fn_select_ProductionByCode_Result>();
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
                CurProduct.AddRange(fm.lRec1.Where(x => x.MaterialCode == ProductCode).ToList());

                List<DateTime> DDList = new List<DateTime>();  // DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();

                switch (tabControl2.SelectedIndex)
                {
                    case 1:
                        if (dgvPlanned.RowCount <= 1)
                        {
                            dgvPlanned.Rows.Clear();

                            if (cbFilter.Checked == false)
                            {
                                DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();
                            }
                            else
                            {
                                DDList = DList.Where(x => x.Date >= dtpPlan1.Value.Date && x.Date <= dtpPlan2.Value.Date).OrderBy(x => x.Date).ToList();
                                //   CurProduct = CurProduct.Where(x => x.WorkDate >= dtpPlan1.Value.Date && x.WorkDate <= dtpPlan2.Value.Date).ToList();

                                // UpdateCBPlan();

                                //   CurProduct = fm.lRec.Where(x => x.MaterialCode == ProductCode).ToList();
                            }

                            UsingComponentPlan(ProductCode, DDList, CurProduct, "", "");
                        }
                        break;
                    case 0:
                        if (dgvFactis.RowCount <= 1)
                        {
                            dgvFactis.Rows.Clear();

                            if (cbFilter.Checked == false)
                            {
                                DDList = DList.Where(x => x.Date >= DateTime.Today.Date.AddDays(-100)).OrderBy(x => x.Date).ToList();
                            }
                            else
                            {
                                DDList = DList.Where(x => x.Date >= dtpFact1.Value.Date && x.Date <= dtpFact2.Value.Date).OrderBy(x => x.Date).ToList();
                            }

                            FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, DDList.Min()).ToList();

                            //}
                            UsingComponentFact(ProductCode, DDList, FactData, "", "");
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
                            List<PR_GetRecipesDataMainProductBis_Result> LResult = new List<PR_GetRecipesDataMainProductBis_Result>();
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
                                    ((double)((decimal)Q/(decimal)lr.Quantity)).ToString("N2"),
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

            }

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

        private void button1_f_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 2).ToList();

            if (CFMData.Count>0)
            { 
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Меркурий";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Истек срок годности по Меркурию";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;                
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i<CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }                          

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();           
        }

        private void button1_b_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 1).ToList();

            if (CFMData.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Пользователь";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Материалы выделенные пользователем";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button1_d_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 3).ToList();

            if (CFMData.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "ОПР";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Материалы ОПР";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button1_c_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 4).ToList();
            CFMData.AddRange(CFMList.Where(x => CFMListGold.Contains(x.MaterialCode)).ToList());
            CFMData = CFMData.Distinct().OrderBy(x => x.MaterialName).ToList();

            if (CFMData.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Дефицит";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Нет заявки, будет дефицит";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button1_h_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 5).ToList();

            if (CFMData.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Нет в рецептах";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Материалы не используемые в рецептуре";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button1_e_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 6).ToList();

            if (CFMData.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Есть заявки";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Не используемые материалы, есть заявки";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button1_g_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 7).ToList();

            if (CFMData.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Есть остатки";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Не используемые материалы, есть остатки";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button1_a_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var CFMData = CFMList.Where(x => x.MaterialStatus == 8).ToList();
            CFMData.AddRange(CFMList.Where(x => x.MaterialGroup == "" || x.MaterialGroup == "Н/Д").ToList());
            CFMData.AddRange(CFMList.Where(x => x.MaterialMultiplyString == "" || x.MaterialMultiplyString == "Н/Д").ToList());
            CFMData.AddRange(CFMList.Where(x => x.WaitingDaysString == "" || x.WaitingDaysString == "Н/Д").ToList());
            CFMData.AddRange(CFMList.Where(x => x.StorageQuantityString == "" || x.StorageQuantityString == "Н/Д").ToList());
            CFMData.AddRange(CFMList.Where(x => x.Sender == "" || x.Sender == "Н/Д").ToList());

            var CFMData1 = CFMData.GroupBy(x => new { x.MaterialCode, x.MaterialName, x.MaterialGroup, x.MaterialMultiplyString, x.WaitingDaysString, x.StorageQuantityString, x.Responsible, x.Sender }).
              Select(g => new AbsentMaterial
              {
                  MaterialCode = g.Key.MaterialCode,
                  MaterialName = g.Key.MaterialName,
                  MaterialGroup = g.Key.MaterialGroup,
                  MaterialMultiplyString = g.Key.MaterialMultiplyString,
                  WaitingDaysString = g.Key.WaitingDaysString,
                  StorageQuantityString = g.Key.StorageQuantityString,
                  Responsible = g.Key.Responsible,
                  Sender = g.Key.Sender
              }).ToList();

            CFMData1 = CFMData1.Distinct().OrderBy(x=>x.MaterialName).ToList();

            if (CFMData1.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Нет данных";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Материалы без параметров поставки";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < CFMData1.Count; i++)
                {
                    wSheet.Cells[CurrRow + i, 1] = CFMData1[i].MaterialCode;
                    wSheet.Cells[CurrRow + i, 2] = CFMData1[i].MaterialName;
                    wSheet.Cells[CurrRow + i, 3] = CFMData1[i].MaterialGroup;
                    wSheet.Cells[CurrRow + i, 4] = CFMData1[i].MaterialMultiplyString;
                    wSheet.Cells[CurrRow + i, 5] = CFMData1[i].WaitingDaysString;
                    wSheet.Cells[CurrRow + i, 6] = CFMData1[i].StorageQuantityString;
                    wSheet.Cells[CurrRow + i, 7] = CFMData1[i].Responsible;
                    wSheet.Cells[CurrRow + i, 8] = CFMData1[i].Sender;
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + CFMData1.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + CFMData1.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void btnSearchPlan_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.CurrentRow != null)
            {
                dgvPlanned.Rows.Clear();

                string ProductCode = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Trim();
                var CurProduct = fm.lRec.Where(x => x.MaterialCode == ProductCode).ToList();
                List<DateTime> DDList = DList.Where(x => x.Date >= dtpPlan1.Value.Date && x.Date <= dtpPlan2.Value.Date).OrderBy(x => x.Date).ToList();

                UsingComponentPlan(ProductCode, DDList, CurProduct, tbCodePlan.Text, tbNamePlan.Text);
            }
        }

        private void btnSearchFact_Click(object sender, EventArgs e)
        {
            if (dataGridViewMain.CurrentRow != null)
            {
                dgvFactis.Rows.Clear();

                string ProductCode = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Trim();
                List<DateTime> DDList = DList.Where(x => x.Date >= dtpFact1.Value.Date && x.Date <= dtpFact2.Value.Date).OrderBy(x => x.Date).ToList();

                DateTime StartDate1;
                DateTime EndDate1;

                try
                {
                    StartDate1 = DDList.Min();
                    EndDate1 = DDList.Max();
                }
                catch (Exception xx)
                {
                    StartDate1 = DateTime.Today.Date;
                    EndDate1 = StartDate1.AddDays(1);
                }

                var FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, StartDate1).ToList();

                UsingComponentFact(ProductCode, DDList, FactData, tbCodeFact.Text, tbNameFact.Text);
            }
        }

        private void UsingComponentPlan(string ProductCode, List<DateTime> DDList, List<PR_GetRecipesDataMainProductBis_Result> CurProduct, string ProductionCode, string ProductionName)
        {
            Cursor = Cursors.WaitCursor;

            string[] dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };
            dgvPlanned.Rows.Add(dataString);
            DateTime DT = DateTime.Today.Date;

            var E1 = XLDataList.Where(x => x.ProdCode == ProductCode && x.DateWork > DT).OrderBy(x => x.DateWork).ThenBy(x => x.Line).ThenBy(x=>x.ProdCode).ThenBy(x => x.ProdName).ToList();
            double Summ = 0;
            double PSumm = 0;

            for (int i = 0; i< E1.Count; i++)
            {
                string PC = E1[i].ParentCode;
                string MC = E1[i].ProdCode;
                string LineName = E1[i].Line;
                DT = E1[i].DateWork;
               // double PVal = 0;

                var CP = CurProduct.Where(x => x.MaterialCode == MC && x.ProductCode == PC ).ToList();   //Norma!  && x.LineNumber == LineName

                if (CP.Count > 0)
                {
                    var EE = XLDataList.Where(x => x.ProdCode == PC && x.DateWork == DT).ToList();

                    if (EE.Count > 0)
                    {
                        dataString = new string[]
                        {
                            E1[i].DateWork.ToString("dd.MM.yyyy"),
                            E1[i].Line,
                            EE[0].Quantity.ToString("N2"),
                            ((double)CP[0].Quantity*1000).ToString("N3"),
                            E1[i].Quantity.ToString("N3"),
                            EE[0].ProdCode,
                            EE[0].ProdName
                        };

                        dgvPlanned.Rows.Add(dataString);

                        Summ = Summ + E1[i].Quantity;
                        PSumm = PSumm + EE[0].Quantity;
                    }

                    var EEE = fm.ExcelDataList.Where(x => x.ProdCode == PC && x.DateWork == DT).ToList();

                    if (EEE.Count > 0)
                    {
                        dataString = new string[]                    {
                            E1[i].DateWork.ToString("dd.MM.yyyy"),
                            E1[i].Line,
                            EEE[0].Quantity.ToString("N2"),
                            ((double)CP[0].Quantity*1000).ToString("N3"),
                            E1[i].Quantity.ToString("N3"),
                            EEE[0].ProdCode,
                            EEE[0].ProdName
                        };

                        dgvPlanned.Rows.Add(dataString);

                        Summ = Summ + E1[i].Quantity;
                        PSumm = PSumm + EEE[0].Quantity;
                    }

                }
            }

            if (dgvPlanned.RowCount > 1)
            {
                dataString = new string[]
                    {
                      "",
                     "ИТОГО:",
                     PSumm.ToString("N2"),
                     "",
                     Summ.ToString("N3"),
                     ""
                    };

                dgvPlanned.Rows.Add(dataString);
            }

            #region Hidden1
            /*  var E1 = fm.ExcelDataList.Where(x => DDList.Contains(x.DateWork) && x.Quantity != 0).ToList();
                string[] dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };
                dgvPlanned.Rows.Add(dataString);

              var E2 = CurProduct.Select(x => x.ProductCode).ToList();
              E2 = E2.Distinct().ToList();

              var E3 = XLDataList.Where(x => E2.Contains(x.ProdCode)).ToList();           //CFMList.Where(x => E2.Contains(x.MaterialCode)).ToList();
           //   var E11 = E1.Where(x => E2.Contains(x.ProdCode)).ToList();


              if (cbFilter.Checked == true)
              {
                  if (!String.IsNullOrEmpty(ProductionCode))
                  {
                      E1 = E1.Where(x => x.ProdCode.Contains(ProductionCode)).ToList();
                  }

                  if (!String.IsNullOrEmpty(ProductionName))
                  {
                      ProductionName = ProductionName.ToLower();

                      E1 = E1.Where(x => x.ProdName.ToLower().Contains(ProductionName)).ToList();
                  }
              }

              double Summ = 0;
              double PSumm = 0;

              for (int i = 0; i < DDList.Count; i++)  // foreach (DateTime dt1 in DDList)   //DList
              {
                  //planned MainProd
                  //  if (dt1 >= DateTime.Today.Date)
                  {
                      bool fl = dataGridViewMain.Rows[RowInd].Cells[2].Value.ToString().ToLower() == "тара" ? true : false;  //Cells["Класс материала"]
                      var ELDListDate = E1.Where(x => (x.DateWork == DDList[i])).ToList();   //fm.ExcelDataList.Where(x => (x.DateWork == dt1) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //dt1

                      for (int j = 0; j < ELDListDate.Count; j++)   // foreach (var ELD in ELDListDate)
                      {
                          var CP = CurProduct.Where(x => (x.ProductCode == ELDListDate[j].ProdCode) &&             (x.WorkDate <= DDList[i])).ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
                        PR_GetRecipesDataMainProduct_Result SRR;
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
                            double PVal = 0;
                            //раскрутка рецептуры в обратном направлении! Продукт-Линия-совпадает ли дата?
                            var PRL = fm.ProdRecipeList.Where(x => x.ProductCode == ELDListDate[j].ProdCodeStr).ToList();
                            if (PRL.Count > 0)
                            {
                                var LL = PRL[0].RecLineList; //.Where(x => x.LineNumber == ELDListDate[j].Line).ToList();
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
                                                        PVal = (double)SRR.Quantity * ELDListDate[j].Quantity;

                                                        dataString = new string[] {
                                                                    DDList[i].ToString("dd.MM.yyyy"),  //  CurrDate
                                                                    ELDListDate[j].Line,
                                                                    ELDListDate[j].Quantity.ToString("N2"),
                                                                    ((double) SRR.Quantity*1000).ToString("N3"),   //CP[0].Quantity
                                                                    PVal.ToString("N3"),        //(((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                                                                    ELDListDate[j].ProdCodeStr,
                                                                    ELDListDate[j].AlterProdName
                                                                    };

                                                        dgvPlanned.Rows.Add(dataString);

                                                        PSumm = PSumm + PVal;
                                                        Summ = Summ + ELDListDate[j].Quantity;
                                                    }
                                                }
                                            }
                                            else
                                            {
                                               // if ((double)SRR.Quantity > 0)
                                                {
                                                    PVal = (double)SRR.Quantity * ELDListDate[j].Quantity;

                                                    dataString = new string[] {
                                                                DDList[i].ToString("dd.MM.yyyy"),
                                                                ELDListDate[j].Line,
                                                               ELDListDate[j].Quantity.ToString("N2"),
                                                                ((double) SRR.Quantity*1000).ToString("N3"),
                                                               PVal.ToString("N3"),          //  (((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                                                               ELDListDate[j].ProdCodeStr,
                                                               ELDListDate[j].AlterProdName
                                                                };

                                                    dgvPlanned.Rows.Add(dataString);

                                                    PSumm = PSumm + PVal;
                                                    Summ = Summ + ELDListDate[j].Quantity;
                                                }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                    ELDListDate = null;
                }                       // dt1
            }  //int i

            //  if (cbFilter.Checked == true)
            {
                if (dgvPlanned.RowCount > 1)
                {
                    dataString = new string[]
                        {
                                    "",
                                    "ИТОГО ГП:",
                                    Summ.ToString("N2"),
                                    "",
                                    PSumm.ToString("N3"),
                                    ""
                        };

                    dgvPlanned.Rows.Add(dataString);
                }

                dataString = new string[]
                    {
                        "","","","",""
                    };

                dgvPlanned.Rows.Add(dataString);
            }

            int RC = 0;
            Summ = 0;
            PSumm = 0;

            for (int i = 0; i < DDList.Count; i++)
            {
                //Planned semiProd
                var E34 = E3.Where(x => x.DateWork == DDList[i]).ToList();

                for (int j = 0; j < E34.Count; j++)
                {
                    string MC = E34[j].ProdCode;
                    var CP = CurProduct.Where(x => x.ProductCode == MC).ToList();

                    double PVal = (E34[j].Quantity * (double)CP[0].Quantity);

                    dataString = new string[] {
                             DDList[i].ToString("dd.MM.yyyy"),
                             CP[0].LineNumber,
                             E34[j].Quantity.ToString("N2"),
                             ((double)CP[0].Quantity*1000).ToString("N3"),
                             PVal.ToString("N3"),          //  (((double) SRR.Quantity) * ELDListDate[j].Quantity).ToString("N0"),
                             E34[j].ProdCode,
                             E34[j].ProdName
                         };

                    dgvPlanned.Rows.Add(dataString);

                    PSumm = PSumm + PVal;
                    Summ = Summ + E34[j].Quantity;

                    RC = RC + 1;
                }
            }

            if (RC > 0)
            {
                {
                    dataString = new string[]
                        {
                                    "",
                                    "ИТОГО ПФ:",
                                    Summ.ToString("N2"),
                                    "",
                                    PSumm.ToString("N3"),
                                    ""
                        };

                    dgvPlanned.Rows.Add(dataString);
                }
            }*/
            #endregion

            foreach (DataGridViewRow r in dgvPlanned.Rows)
            {
                if (r.Cells[0].Value.ToString() != "")
                {
                    r.Cells[0].ToolTipText = "Количество: " + r.Cells[2].Value.ToString() + "; Норма на тонну: " + r.Cells[3].Value.ToString() + " Итого расход: " + r.Cells[4].Value.ToString();
                    r.Cells[1].ToolTipText = "Количество: " + r.Cells[2].Value.ToString() + "; Норма на тонну: " + r.Cells[3].Value.ToString() + " Итого расход: " + r.Cells[4].Value.ToString();
                    r.Cells[2].ToolTipText = "Количество: " + r.Cells[2].Value.ToString() + "; Норма на тонну: " + r.Cells[3].Value.ToString() + " Итого расход: " + r.Cells[4].Value.ToString();
                    r.Cells[3].ToolTipText = "Количество: " + r.Cells[2].Value.ToString() + "; Норма на тонну: " + r.Cells[3].Value.ToString() + " Итого расход: " + r.Cells[4].Value.ToString();
                    r.Cells[4].ToolTipText = "Количество: " + r.Cells[2].Value.ToString() + "; Норма на тонну: " + r.Cells[3].Value.ToString() + " Итого расход: " + r.Cells[4].Value.ToString();
                }
            }

            Cursor = Cursors.Default;
            GC.Collect();
        }

        private void UsingComponentFact(string ProductCode, List<DateTime> DDList, List<fn_select_ProductionByCode_Result> FactData, string ProductionCode, string ProductionName)
        {
            Cursor = Cursors.WaitCursor;
            string[] dataString = new string[] { "", "", "", "", "", "", "Данный компонент используется в:" };

            DateTime StartDate1;
            DateTime EndDate1;

            dgvFactis.Rows.Add(dataString);
            {
                bool fl = dataGridViewMain.CurrentRow.Cells[2].Value.ToString().ToLower() == "тара" ? true : false;
                //  var MuaSection = fm.MUAList.Where(x => (x.DateTime >= DateStart) && (x.DateTime < DateEnd)).ToList();    

                FactData = FactData.OrderBy(x => x.LotForErp).ToList();       //fm.tdb.fn_select_ProductionByCode(ProductCode, StartDate1).OrderBy(x => x.LotForErp).ToList();  
                                                                              //   DateTime.Today.Date.AddYears(-2)               //fm.edb.fn_select_MaterialUsingByProduct(StartDate1, EndDate1, ProductCode).OrderBy(x => x.DateStart).ToList();
                if (cbFilter.Checked == true)
                {
                    if (!String.IsNullOrEmpty(ProductionCode))
                    {
                        FactData = FactData.Where(x => x.MaterialCode == ProductionCode).ToList();
                    }

                    if (!String.IsNullOrEmpty(ProductionName))
                    {
                        ProductionName = ProductionName.ToLower();

                        FactData = FactData.Where(x => x.MaterialName.ToLower().Contains(ProductionName)).ToList();
                    }
                }

                List<string> LotForErp = FactData.Select(x => x.LotForErp).Distinct().ToList();
                //  List<fn_select_ProductionByCode_Result> FactData1 = new List<fn_select_ProductionByCode_Result>();

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

            //    if (cbFilter.Checked == true)
                {
                    if (dgvFactis.RowCount > 1)
                    {
                        double Summ = FactData.Sum(x => (double)x.Quantity);
                        double RSumm = FactData.Sum(x => (double)x.RawQuant);

                        dataString = new string[]
                        {
                                       RSumm.ToString("N3"),
                                       "",
                                       "",
                                       "",
                                       Summ.ToString("N2"),
                                       ""
                        };

                        dgvFactis.Rows.Add(dataString);
                    }
                }

                FactData = null;
            }  //fact

            Cursor = Cursors.Default;
            GC.Collect();
        }

        private void dtpFact1_ValueChanged(object sender, EventArgs e)
        {
            if (dtpFact1.Value > dtpFact2.Value)
            {
                dtpFact2.Value = dtpFact1.Value.Date.AddDays(1);
            }
        }

        private void dtpFact2_ValueChanged(object sender, EventArgs e)
        {
            if (dtpFact2.Value < dtpFact1.Value)
            {
                dtpFact1.Value = dtpFact2.Value.Date.AddDays(-1);
            }
        }

        private void dtpPlan1_ValueChanged(object sender, EventArgs e)
        {
            if (dtpPlan1.Value > dtpPlan2.Value)
            {
                dtpPlan2.Value = dtpPlan1.Value.Date.AddDays(1);
            }
        }

        private void dtpPlan2_ValueChanged(object sender, EventArgs e)
        {
            if (dtpPlan2.Value < dtpPlan1.Value)
            {
                dtpPlan1.Value = dtpPlan2.Value.Date.AddDays(-1);
            }
        }

        private void checkBox1_CheckedChanged(object sender, EventArgs e)
        {
            if (cbFilter.Checked == true)
            {
                dtpFact1.Visible = true;
                dtpFact2.Visible = true;

                dtpPlan1.Visible = true;
                dtpPlan2.Visible = true;

                label10.Visible = true;
                label9.Visible = true;

                label8.Visible = true;
                label7.Visible = true;

                btnSearchFact.Visible = true;
                btnSearchPlan.Visible = true;

                tbCodePlan.Text = "";
                tbCodePlan.Visible = true;

                tbCodeFact.Text = "";
                tbCodeFact.Visible = true;

                tbNamePlan.Visible = true;
                tbNamePlan.Text = "";

                tbNameFact.Visible = true;
                tbNameFact.Text = "";
            }
            else
            {
                dtpFact1.Visible = false;
                dtpFact2.Visible = false;

                dtpPlan1.Visible = false;
                dtpPlan2.Visible = false;

                label10.Visible = false;
                label9.Visible = false;

                label8.Visible = false;
                label7.Visible = false;

                btnSearchFact.Visible = false;
                btnSearchPlan.Visible = false;

                tbCodeFact.Visible = false;
                tbCodePlan.Visible = false;

                tbNamePlan.Visible = false;
                tbNameFact.Visible = false;
            }
        }

        private void dataGridViewMain_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode.ToString() == "Escape")
            {
                var cfm1 = CFMList.Where(x => x.MaterialCode == UserSelectedCode).ToList();

                if (cfm1.Count > 0)
                {
                    if (ColInd > 15 + WorkMonthList.Count)
                    {
                        string LName = "";
                        DateTime CurrDate;

                        if ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Остаток на")) || (dataGridViewMain.Columns[ColInd].HeaderText.Contains("Потребление на")) || ((dataGridViewMain.Columns[ColInd].HeaderText.Contains("Дефицит на"))))
                        {
                            LName = dataGridViewMain.Columns[ColInd].HeaderText;
                            LName = LName.Substring(LName.Length - 11, 11).Trim();
                            CurrDate = fm.ConvertDataToDate(LName);
                        }
                        else
                        {
                            CurrDate = DateTime.Today.Date;
                        }

                        if (CurrDate >=DateTime.Today.Date)
                        {
                            USC = new UserSelectedCell
                            {
                                UserSelectedCode = UserSelectedCode,
                                UserSelectedProdName = UserSelectedProdName,
                                UserSelectedDateTime = CurrDate,
                                RowIndex = RowInd,
                                ColIndex = ColInd,
                            };

                            try
                            {
                                USC.Quantity = fm.ConvertStringToDouble(dataGridViewMain.Rows[RowInd].Cells[ColInd].Value.ToString().Trim());
                            }
                            catch (Exception xx)
                            {
                                USC.Quantity = 0;
                            }

                            Rectangle rect = dataGridViewMain.GetCellDisplayRectangle(ColInd, RowInd, true);
                         //   AxX = rect.X +/* fm.Left +*/ this.Left;
                         //   AxY = rect.Y +/* fm.Top + */this.Top + this.panel1.Height + 60;

                            if (Application.OpenForms["FormTransfer"] == null)
                            {
                                Cursor = Cursors.WaitCursor;
                                ft = new FormTransfer(fm, this, USC, cfm1);
                                ft.Visible = true;
                                //  fr.MdiParent = fm;
                                ft.WindowState = FormWindowState.Normal;
                                Cursor = Cursors.Default;
                            }
                            else
                            {
                                ft.BringToFront();
                                //   cf.UpdateData();
                            }

                            if (AxX > ft.Width / 2)
                            {
                                ft.Left = AxX - 25;
                            }
                            else
                            {
                                ft.Left = AxX;
                            }
                            if (AxY > ft.Height / 2)
                            {
                                ft.Top = AxY - 25;
                            }
                            else
                            {
                                ft.Top = AxY;
                            }

                            this.Enabled = false;
                        }
                    }
                }
            }
        }

        private void UpdateFormWW()
        {
            try
            {
                UW = new UpdateWindow();
                UW.ShowInTaskbar = false;
                UW.ShowDialog();
            }
            catch (Exception xx)
            {
                if (UW != null)
                {
                    UW.Close();
                }
                UW = null;
            }

            while (LoadFlag == true)
            {
                System.Threading.Thread.Sleep(100);
            }

            UW.Close();
        }

        private void button11_Click(object sender, EventArgs e)
        {
            UpdateData();
        }

        private void UpdateData()
        {
            Cursor = Cursors.WaitCursor;

            LoadFlag = true;
            this.Enabled = false;
            System.Threading.Thread thread1 = new System.Threading.Thread(UpdateFormWW);
            thread1.Start();

            fm.LoadStoragesData();

            PrepareData();

            ClearDGV();
            ShowData();

            DGV_Get_CellColor();

            try
            {
                LoadFlag = false;
                System.Threading.Thread.Sleep(250);
                thread1.Abort();
            }
            catch (Exception xx)
            { }

            this.Enabled = true;
            Cursor = Cursors.Default;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormTask"] == null)
            {
                Cursor = Cursors.WaitCursor;
                ftt = new FormTask(fm, this);
                ftt.Visible = true;
                //   fs.MdiParent = this;
                ftt.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                ftt.BringToFront();
                ftt.Visible = true;
                //   cf.UpdateData();
            }

            this.Enabled = false;
        }

        private void UpdateFilters()
        {
            Cursor = Cursors.WaitCursor;
            panel1.Enabled = false;
            ClearDGV();
            RecalculateData();
            //ShowLocalData();
            //  DGV_UpdateColumnZero();
            DGV_Get_CellColor();

            //   LowPageIndex = memoryCache.GetLowIndex();
            //   HighPageIndex = memoryCache.GetHihgIndex();
            //   OldLowPageIndex = LowPageIndex;
            //   OldHighPageIndex = HighPageIndex;

            panel1.Enabled = true;
            Cursor = Cursors.Default;

        }

        private void checkBox2_CheckedChanged(object sender, EventArgs e)
        {
            UpdateFilters();
        }

        private void button1_j_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var Data = fm.SMList.Where(x => x.BOM.Trim() == "ВЫВОД").ToList();
            Data = Data.Distinct().ToList();

            if (Data.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "ВЫВОД";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Материалы со статусом ВЫВОД";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < Data.Count; i++)
                {
                    var CFMData =  CFMList.Where(x => x.MaterialCode == Data[i].MaterialCode).ToList();

                    wSheet.Cells[CurrRow + i, 1] = Data[i].MaterialCode;

                    if (CFMData.Count > 0)
                    {
                        wSheet.Cells[CurrRow + i, 2] = CFMData[0].MaterialName;
                        wSheet.Cells[CurrRow + i, 3] = CFMData[0].MaterialGroup;
                        wSheet.Cells[CurrRow + i, 4] = CFMData[0].MaterialMultiplyString;
                        wSheet.Cells[CurrRow + i, 5] = CFMData[0].WaitingDaysString;
                        wSheet.Cells[CurrRow + i, 6] = CFMData[0].StorageQuantityString;
                        wSheet.Cells[CurrRow + i, 7] = CFMData[0].Responsible;
                        wSheet.Cells[CurrRow + i, 8] = CFMData[0].Sender;
                    }
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + Data.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + Data.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();

        }

        private void button1_i_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var Data = fm.SMList.Where(x => x.BOM.Trim() != "ВЫВОД").ToList();
            Data = Data.Distinct().ToList();

            if (Data.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Нет данных";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Материалы со статусом N/A";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < Data.Count; i++)
                {
                    var CFMData = CFMList.Where(x => x.MaterialCode == Data[i].MaterialCode).ToList();

                    wSheet.Cells[CurrRow + i, 1] = Data[i].MaterialCode;

                    if (CFMData.Count > 0)
                    {
                        wSheet.Cells[CurrRow + i, 2] = CFMData[0].MaterialName;
                        wSheet.Cells[CurrRow + i, 3] = CFMData[0].MaterialGroup;
                        wSheet.Cells[CurrRow + i, 4] = CFMData[0].MaterialMultiplyString;
                        wSheet.Cells[CurrRow + i, 5] = CFMData[0].WaitingDaysString;
                        wSheet.Cells[CurrRow + i, 6] = CFMData[0].StorageQuantityString;
                        wSheet.Cells[CurrRow + i, 7] = CFMData[0].Responsible;
                        wSheet.Cells[CurrRow + i, 8] = CFMData[0].Sender;
                    }
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + Data.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + Data.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();
        }

        private void button13_Click(object sender, EventArgs e)
        {
            string ProductCode = "";

            if (InputBox.Query("Введите код продукции либо нажмите Cancel", "Текущее значение:", ref ProductCode))
            {
                lData LD1;
                List<lData> DataListTempNew1 = new List<lData>();
                List<lData> DataListTempOld1 = new List<lData>();

                var Recipe = fm.ProdRecipeList.Where(x => x.ProductCode.Trim() == ProductCode).ToList();

                if (Recipe.Count > 0)
                {
                    var RecLine = Recipe[0].RecLineList;
                    if (RecLine.Count > 0)
                    {
                        var RecLineDate = RecLine[0].RecList;
                        if (RecLineDate.Count > 0)
                        {
                            var RList = RecLineDate[0].rList;

                            foreach (var RL in RList)
                            {
                                LD1 = new lData();

                                LD1.ProdCode = RL.MaterialCode.Trim();
                              //  LD1.ProdCodeStr = RL.MaterialCode.Trim();
                                LD1.ParentCode = ProductCode;
                                LD1.DateWork = RecLineDate[0].WorkDateStart;
                                LD1.ProdName = RL.MaterialName.Trim();
                                LD1.Quantity = 1000 * RL.MaterialQuantity;
                                LD1.ParentQuant = 1000;
                                LD1.ResValue = RL.MaterialQuantity;
                                LD1.Line = RecLine[0].LineNumber;
                                LD1.IPG = RL.IPG;
                                LD1.Spices = false;   //control: split 1-st level only!!!!!!!!!
                                LD1.BOM = RL.BOM;
                                LD1.IsSemoProd = RL.IsSemiProd;

                                bool fl = false;

                                for (int i = 0; i < DataListTempNew1.Count; i++)
                                {
                                    if (DataListTempNew1[i].ProdCode.Trim() == LD1.ProdCode.Trim() && DataListTempNew1[i].ParentCode == LD1.ParentCode)
                                    {
                                      //  if (DataListTempNew1[i].Line.Trim() == LD1.Line.Trim())
                                        {
                                           // if (DataListTempNew1[i].DateWork == LD1.DateWork)
                                            {
                                                if (DataListTempNew1[i].Spices == LD1.Spices)
                                                {
                                                    DataListTempNew1[i].Quantity = DataListTempNew1[i].Quantity + LD1.Quantity;
                                                    fl = true;
                                                    i = DataListTempNew1.Count;
                                                }
                                            }
                                        }
                                    }
                                }
                                if (fl == false)
                                {
                                    DataListTempNew1.Add(LD1);
                                }
                            }
                        }
                    }
                }

                int Iteration = 1;
                var XLDataList1 = new List<lData>();

                //create Excel

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Передел";
                wSheet.Cells[1, 2] = "Код ГП Родитель";
                range1 = wSheet.Cells[1, 2] as Excel1.Range;
                range1.EntireColumn.NumberFormat = "@";
                
                wSheet.Cells[1, 3] = "Код материала/ПФ";
                range1 = wSheet.Cells[1, 3] as Excel1.Range;
                range1.EntireColumn.NumberFormat = "@";

                wSheet.Cells[1, 4] = "Наименование";

                wSheet.Cells[1, 5] = "Количество Родитель";
                range1 = wSheet.Cells[1, 5] as Excel1.Range;
                range1.EntireColumn.NumberFormat = "### ### ###.000";

                wSheet.Cells[1, 6] = "Норма";
                range1 = wSheet.Cells[1, 6] as Excel1.Range;
                range1.EntireColumn.NumberFormat = "### ### ###.000";

                wSheet.Cells[1, 7] = "Количество материал";
                range1 = wSheet.Cells[1, 6] as Excel1.Range;
                range1.EntireColumn.NumberFormat = "### ### ###.000";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 7]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Bold = true;
                range1.EntireColumn.AutoFit();

                range1 = wSheet.Cells[1, 4] as Excel1.Range;
                range1.EntireColumn.ColumnWidth = 50;

                int CurrRow = 2;

                while (DataListTempNew1.Count > 0)
                {
                    Iteration = Iteration - 1;
                    for (int i = DataListTempNew1.Count; i > 0; i--)
                    {
                        bool fl = false;
                        for (int j = 0; j < XLDataList1.Count; j++)
                        {
                            if (XLDataList1[j].ProdCode.Trim() == DataListTempNew1[i - 1].ProdCode.Trim() && DataListTempNew1[i -1].ParentCode == XLDataList1[j].ParentCode)
                            //          if (DataListTempNew1[i].ProdCode.Trim() == LD1.ProdCode.Trim() && DataListTempNew1[i].ParentCode == LD1.ParentCode)
                            {
                                //  if (XLDataList1[j].DateWork == DataListTempNew1[i - 1].DateWork)
                                {
                                  //  if (XLDataList1[j].Line.Trim() == DataListTempNew1[i - 1].Line.Trim())
                                    {
                                        if (XLDataList1[j].Spices == DataListTempNew1[i - 1].Spices)
                                        {
                                            fl = true;
                                            XLDataList1[j].Quantity = XLDataList1[j].Quantity + DataListTempNew1[i - 1].Quantity;
                                            j = XLDataList1.Count;
                                        }
                                    }
                                }
                            }
                        }


                        wSheet.Cells[CurrRow, 1] = Iteration;
                        wSheet.Cells[CurrRow, 2] = DataListTempNew1[i - 1].ParentCode;
                        wSheet.Cells[CurrRow, 3] = DataListTempNew1[i - 1].ProdCode;
                        wSheet.Cells[CurrRow, 4] = DataListTempNew1[i - 1].ProdName;
                        wSheet.Cells[CurrRow, 5] = DataListTempNew1[i - 1].ParentQuant;
                        wSheet.Cells[CurrRow, 6] = DataListTempNew1[i - 1].ResValue;
                        wSheet.Cells[CurrRow, 7] = DataListTempNew1[i - 1].Quantity;
                        CurrRow = CurrRow + 1;

                        if (fl == false)
                        {
                            LD1 = new lData
                            {
                                ProdCode = DataListTempNew1[i - 1].ProdCode.Trim(),
                                ProdName = DataListTempNew1[i - 1].ProdName.Trim(),
                                ParentCode = DataListTempNew1[i - 1].ParentCode,
                                DateWork = DataListTempNew1[i - 1].DateWork,
                                Quantity = DataListTempNew1[i - 1].Quantity,
                                ParentQuant = DataListTempNew1[i - 1].ParentQuant,
                                ResValue = DataListTempNew1[i - 1].ResValue,
                                IPG = DataListTempNew1[i - 1].IPG,
                                Line = DataListTempNew1[i - 1].Line.Trim(),
                                Spices = DataListTempNew1[i - 1].Spices,
                                BOM = DataListTempNew1[i - 1].BOM,
                                IsSemoProd = DataListTempNew1[i - 1].IsSemoProd
                            };

                            if (LD1.Quantity > 0)         // ((LD.IPG != "ГП") && (LD.Quantity > 0))
                            {
                                XLDataList1.Add(LD1);
                            }
                        }

                        if (DataListTempNew1[i - 1].IsSemoProd != 1)
                        {
                            DataListTempNew1.RemoveAt(i - 1);
                        }
                    }

                    DataListTempOld1.Clear();
                    DataListTempOld1.AddRange(DataListTempNew1);
                    DataListTempNew1.Clear();

                    for (int i = 0; i < DataListTempOld1.Count; i++)
                    {
                        var RecList = fm.ProdRecipeList.Where(x => x.ProductCode.Trim() == DataListTempOld1[i].ProdCode.Trim()).ToList();

                        if (RecList.Count > 0)
                        {
                            var RecLineList = RecList[0].RecLineList;
                            {
                                if (RecLineList.Count > 0)
                                {
                                    var RecLineDateList = RecLineList[0].RecList;

                                    if (RecLineDateList.Count > 0)
                                    {
                                        var RList = RecLineDateList[0].rList;

                                        foreach (var RR in RList)
                                        {
                                            LD1 = new lData
                                            {
                                                ProdCode = RR.MaterialCode.Trim(),
                                                ParentCode = DataListTempOld1[i].ProdCode,
                                                ProdCodeStr = RR.MaterialCode.Trim(),
                                                DateWork = DataListTempOld1[i].DateWork,
                                                ProdName = RR.MaterialName.Trim(),
                                                ParentQuant = DataListTempOld1[i].Quantity,
                                                Quantity = RR.MaterialQuantity * DataListTempOld1[i].Quantity,
                                                ResValue = RR.MaterialQuantity,
                                                Line = RecLineList[0].LineNumber.Trim(),    //DataListTempOld[i].Line.Trim(),
                                                IPG = RR.IPG,
                                                BOM = RR.BOM,
                                                IsSemoProd = RR.IsSemiProd
                                            };

                                            bool fl = false;
                                            for (int j = 0; j < DataListTempNew1.Count; j++)
                                            {
                                                if (DataListTempNew1[j].ProdCode == LD1.ProdCode && DataListTempNew1[j].ParentCode == LD1.ParentCode)
                                                {
                                                  //  if (DataListTempNew[j].Line == LD.Line)
                                                    {
                                                    //    if (DataListTempNew[j].DateWork == LD.DateWork)
                                                        {
                                                            DataListTempNew1[j].Quantity = DataListTempNew1[j].Quantity + LD1.Quantity;
                                                            fl = true;
                                                            j = DataListTempNew1.Count;
                                                        }
                                                    }
                                                }
                                            }

                                            if (fl == false)
                                            {
                                                DataListTempNew1.Add(LD1);
                                                //  if ((LD.ProdCode == "1031008522") || (LD.ProdCodeStr == "1031008522"))
                                                //  { }
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }
                }

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[CurrRow - 1, 7]];
                range1.WrapText = true;

                fm.SetBorders(range1, 1, true);

                wSheet = null;
                wBook = null;
                xlApp = null;

            }
            else
            {
                //No data
            }
                
        }

        private void button1_k_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;
            var Data = CFMList.Where(x => x.MaterialStatus == -1).ToList();
            Data = Data.Distinct().ToList();

            if (Data.Count > 0)
            {
                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Новые";
                Excel1.Range range1;

                wSheet.Cells[1, 1] = "Новые материалы";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Merge();
                range1.Font.Size = 14;
                range1.Font.Bold = true;

                wSheet.Cells[3, 1] = "Код материала";
                wSheet.Cells[3, 2] = "Наименование";
                wSheet.Cells[3, 3] = "Класс материала";
                wSheet.Cells[3, 4] = "Кратность поставки";
                wSheet.Cells[3, 5] = "Мин.срок пост-ки, дней";
                wSheet.Cells[3, 6] = "Страховой запас";
                wSheet.Cells[3, 7] = "Исполнитель";
                wSheet.Cells[3, 8] = "Поставщик";

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3, 8]];
                range1.HorizontalAlignment = Excel1.Constants.xlCenter;
                range1.VerticalAlignment = Excel1.Constants.xlCenter;
                range1.Font.Size = 12;
                range1.Font.Bold = true;

                wSheet.Columns["A:H"].ColumnWidth = 25;
                wSheet.Columns["B:B"].ColumnWidth = 35;
                wSheet.Columns["C:C"].ColumnWidth = 35;
                wSheet.Columns["H:H"].ColumnWidth = 35;

                int CurrRow = 4;

                for (int i = 0; i < Data.Count; i++)
                {
                    var CFMData = CFMList.Where(x => x.MaterialCode == Data[i].MaterialCode).ToList();

                    wSheet.Cells[CurrRow + i, 1] = Data[i].MaterialCode;

                    if (CFMData.Count > 0)
                    {
                        wSheet.Cells[CurrRow + i, 2] = CFMData[0].MaterialName;
                        wSheet.Cells[CurrRow + i, 3] = CFMData[0].MaterialGroup;
                        wSheet.Cells[CurrRow + i, 4] = CFMData[0].MaterialMultiplyString;
                        wSheet.Cells[CurrRow + i, 5] = CFMData[0].WaitingDaysString;
                        wSheet.Cells[CurrRow + i, 6] = CFMData[0].StorageQuantityString;
                        wSheet.Cells[CurrRow + i, 7] = CFMData[0].Responsible;
                        wSheet.Cells[CurrRow + i, 8] = CFMData[0].Sender;
                    }
                }

                range1 = wSheet.Range[wSheet.Cells[4, 1], wSheet.Cells[3 + Data.Count, 8]];
                range1.WrapText = true;
                range1.Rows.AutoFit();

                range1 = wSheet.Range[wSheet.Cells[3, 1], wSheet.Cells[3 + Data.Count, 8]];
                fm.SetBorders(range1, 1, true);

                range1 = null;
                wSheet = null;
                wBook = null;
                xlApp = null;

                Cursor = Cursors.Default;
                MessageBox.Show("Данные сформированы.", "Сообщение системы");
            }
            else
            {
                Cursor = Cursors.Default;
                MessageBox.Show("Данные отсутствуют!", "Сообщение системы");
            }

            GC.Collect();

        }

        private void checkBox6_CheckedChanged(object sender, EventArgs e)
        {
            dateTimePicker2.Visible = checkBox6.Checked;
            dateTimePicker2.Enabled = checkBox6.Checked;
        }

        private void button15_Click(object sender, EventArgs e)
        {
            XLDataListFilter = fm.ExcelDataList;

            if (checkBox6.Checked == true)
            {
                DateTime DT = dateTimePicker2.Value.Date;
                XLDataListFilter = XLDataListFilter.Where(x => x.DateWork == DT).ToList();
            }

            if (tb_Code.Text.Trim() != "")
            {
                XLDataListFilter = XLDataListFilter.Where(x => x.ProdCodeStr.ToLower().Contains(tb_Code.Text.ToLower().Trim())).ToList();
            }

            if (tbName.Text.Trim() != "")
            {
                XLDataListFilter = XLDataListFilter.Where(x => x.ProdName.ToLower().Contains(tbName.Text.ToLower().Trim())).ToList();
            }

            if (cbLineList.Text != "Линия")
            {
                string L = cbLineList.Text;
                XLDataListFilter = XLDataListFilter.Where(x => x.Line == L).ToList();
            }

            comboBoxFiltering.Items.Clear();
            comboBoxFiltering.Items.Add("Выберите план");


            foreach (var XL in XLDataListFilter)  //.OrderBy(x => x.DateWork).ThenBy(x => x.Line).ToList()
            {
                comboBoxFiltering.Items.Add(/*XL.RowIndex.ToString()+ " " +*/ XL.DateWork.ToString("dd.MM.yyyy") + " " + XL.Line + " " + XL.ProdCodeStr + " " + XL.ProdName + " " + XL.Quantity.ToString("N0") + " кг");
            }

            comboBoxFiltering.Text = "Выберите план";
        }

        private void button14_Click(object sender, EventArgs e)
        {
            checkBox6.Checked = false;
            cbLineList.Text = "Линия";
            tb_Code.Text = "";
            tbName.Text = "";

            button15_Click(sender, e);
        }

        private void cbShortReport_CheckedChanged(object sender, EventArgs e)
        {
            //
        }
    }
}
