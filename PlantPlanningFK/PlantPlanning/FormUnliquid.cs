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
    public partial class FormUnliquid : Form
    {
        FormMain fm;

        List<Material> MaterialDataList;

        public List<StockBalances> SBList;
        public List<StockBalancesResult> SBRList;  //ResultListData!!!!!
        public List<StockBalancesMes> SBMList;
        public List<StockBalancesPlant> SBPList;
        public List<lData> XLDataList;
        public List<ProductRecipes> TBP;
        public List<ProductRecipes> PRList;
        public List<DailyResult> FilteredMaterialDataIn;

        public List<tUserRecordSelected> URSList;
        public List<tMercuryRecordSelected> URMList;
        public List<string> URSCodeList;
        public List<string> URMCodeList;
        public List<tColorSelection> CSList;

        public Color SColor = Color.LightPink;   //.LightPink;  Color color = (Color)ColorConverter.ConvertFromString("#FFDFD991");
        public Color MColor = Color.RosyBrown; //Mercury!

        List<string> MaterialNamesList;
        List<string> ProductNamesList;
        List<DateTime> WorkMonthList;
        List<DateTime> DTList;
        List<DailyResult> MaterialData;

        public BindingSource bs;
        public DataTable dt1;
        public DateTime StartDate;
        public DateTime EndDate;
        Int32 CurrIndex;
        public Int32 SelRowIndes;

        List<string> CFMListGray;



        public FormUnliquid(FormMain FM)
        {
            InitializeComponent();

            fm = FM;

            label6.Visible = false;
            label5.Visible = false;
            button6.Visible = false;
            button5.Visible = false;
            button9.Visible = false;

            MaterialDataList = new List<Material>();
            WorkMonthList = new List<DateTime>();
            MaterialData = new List<DailyResult>();
            XLDataList = new List<lData>();
            dt1 = new DataTable();
            TBP = new List<ProductRecipes>();
            DTList = new List<DateTime>();
            ProductNamesList = new List<string>();
            MaterialNamesList = new List<string>();
            FilteredMaterialDataIn = new List<DailyResult>();

            button1_a.BackColor = Color.LightGray; //    Color.DarkGray; #FFFF00 
            button1_b.BackColor = Color.LightPink;    // Color.LightPink;
            button1_c.BackColor = Color.LightYellow;      //Color.Gold;
            button1_d.BackColor = Color.LightCyan;     // Color.SteelBlue;
            button1_e.BackColor = Color.LightGreen;    //Color.Lime;
            button1_f.BackColor = Color.RosyBrown;

            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;
            dgvFactis.ReadOnly = true;
            dgvFactis.AllowUserToAddRows = false;
            dgvPlanned.ReadOnly = true;
            dgvPlanned.AllowUserToAddRows = false;
            dgvStorages.ReadOnly = true;
            dgvStorages.AllowUserToAddRows = false;

            LoadData();
            GetStorage();
            CalculateData();

            ShowData();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void CalculateData()
        {
            XLDataList = fm.ExcelDataList; //.Where(x => x.DateWork >= DateTime.Today.Date).ToList();
            XLDataList = XLDataList.Where(x => x.Quantity > 0).ToList();

            List<lData> DataListTempOld = new List<lData>();
            List<lData> DataListTempNew = new List<lData>();
            WorkMonthList = new List<DateTime>();
            lData LD;
            bool fl;
            Int32 Iteration = 0;
            List<string> OprComp = fm.OPRTableList.Select(x => x.MatCode).ToList();

            TBP = fm.tdb.tBlockedMaterial.Where(x => x.Using == false)
                  .Select(x => new ProductRecipes
                  {
                      ProductCode = x.ProductCode,
                      ProductName = x.ProductName
                  })
                .ToList();

            ProductNamesList = XLDataList.Select(x => x.ProdCode).ToList();
            ProductNamesList = ProductNamesList.Distinct().ToList();
            DTList = XLDataList.OrderBy(x => x.DateWork).Where(x => x.DateWork <= DateTime.Today.Date).Select(x => x.DateWork).ToList();
            DTList = DTList.Distinct().ToList();

            try
            {
                StartDate = DTList.Min();
                EndDate = DTList.Max();
            }
            catch (Exception xx)
            {
                StartDate = DateTime.Today.Date.AddDays(-1);
                EndDate = StartDate.AddDays(1);
            }

            //Reciepts
            PRList = new List<ProductRecipes>();
            PRList = fm.ProdRecipeList.Where(x => ProductNamesList.Contains(x.ProductCode)).ToList();

            foreach (DateTime dt in DTList)
            {
                var DataList = XLDataList.Where(x => x.DateWork == dt).ToList();

                foreach (var DL in DataList)
                {
                    //   if (DL.ProdCode == "1030001456")
                    //   { }

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

                        if (OprComp.Contains(DataListTempNew[i - 1].ProdCode))
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
                                                /*  if (LD.Spices == false)
                                                  {*/
                                                XLDataList.Add(LD);
                                                /*  }
                                                  else
                                                  {
                                                      if (LD.IPG.ToLower() != "пф")
                                                      {
                                                          XLDataList.Add(LD);
                                                      }
                                                  }*/


                                                //      if ((LD.ProdCode == "1031008522") || (LD.ProdCodeStr == "1031008522"))
                                                //      { }
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        else
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
                                XLDataList.Add(LD);
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


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void GetStorage()
        {
            SBList = new List<StockBalances>();
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
            List<StockBalancesPlant> sbpln = new List<StockBalancesPlant>();

            //nav
            var sbDataList = fm.StockBaklanceList.Where(x => /*(x.TestQualityGr != "3Блок")  &&  (x.StorageType == "Собств") &&*/  (x.WorkDate.Date == LocalStartDate.Date) && (x.WorkDate.Hour >= 18) && (x.WorkDate.Hour <= 22)).ToList();
            //  sbDataList = sbDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
            sbDataList = sbDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
            foreach (SelectStockBalances1_Result ss in sbDataList)
            {
                SBList.Add(fm.NewStockBalance(ss));
            }

            sbDataList = fm.StockBaklanceList.Where(x => /*(x.TestQualityGr != "3Блок") && (x.StorageType == "Собств") && */ (x.WorkDate.Date == LocalEndDate.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            //    sbDataList = sbDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
            sbDataList = sbDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
            foreach (SelectStockBalances1_Result ss in sbDataList)
            {
                SBList.Add(fm.NewStockBalance(ss));
            }

            sbDataList.Clear();
            sbDataList = null;

            //mes
            var sbmDataList = fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == LocalStartDate.Date) && (x.WorkDate.Hour >= 18) && (x.WorkDate.Hour <= 22)).ToList();
            sbmDataList = sbmDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
            sbmDataList = sbmDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
            foreach (SelectStockBalancesMesBis_Result ssm in sbmDataList)
            {
                SBMList.Add(fm.NewStockBalancesMes(ssm));
            }

            sbmDataList = fm.StockBalancesMesList.Where(x => (x.WorkDate.Date == LocalEndDate.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            sbmDataList = sbmDataList.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();
            sbmDataList = sbmDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
            foreach (SelectStockBalancesMesBis_Result ssm in sbmDataList)
            {
                SBMList.Add(fm.NewStockBalancesMes(ssm));
            }

            sbmDataList.Clear();
            sbmDataList = null;

            var sbpDataList = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == LocalEndDate.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10)).ToList();
            sbpDataList = sbpDataList.Where(x => fm.StorageNamesList.Contains(x.LineName)).ToList();

            sbpDataList = sbpDataList.Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();
            foreach (var z in sbpDataList)
            {
                SBPList.Add(fm.NewStockBalancesPlant(z));
            }

            sbpDataList.Clear();
            sbpDataList = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void LoadData()
        {
            var DataList = fm.MaterialDataList1.Where(x => x.LASTPRODUCER == "НЕЛИКВИД").ToList();

            foreach (var dl in DataList)
            {
                Material mm = new Material
                {
                    MaterialCode = dl.MaterialCode,
                    MaterialName = dl.NavMaterialName,
                    MaterialGroup = dl.MaterialGroup,
                    MaterialCategory = dl.SIGname,
                    Responsible = dl.SIGresp,
                    MaterialMultiply = (Int32)dl.Multiplicity,
                    MaterialMultiplyString = (Int32)dl.Multiplicity == 0 ? "" : dl.Multiplicity.ToString(),
                    WaitingDays = (Int32)dl.MinTransferDays,
                    WaitingDaysString = (Int32)dl.MinTransferDays == 0 ? "" : ((Int32)dl.MinTransferDays).ToString(),
                    StorageQuantity = (Int32)dl.SafetyStock,
                    StorageQuantityString = (Int32)dl.SafetyStock == 0 ? "" : ((Int32)dl.SafetyStock).ToString("0"),
                    Sender = " ",

                    RawEnterpriseList = new List<Raws>(),
                    RawStorageList = new List<Raws>(),
                    RawNav1List = new List<Raws>(),
                    RawNav2List = new List<Raws>(),
                    RawMesOut1List = new List<Raws>(),
                    RawMesOut2List = new List<Raws>(),
                    OutList = new List<Raws>(),
                    FactList = new List<Raws>(),
                    OldTaskNavList = new List<Raws>(),
                    ConsurmptionList=new List<Raws>(),

                    RawStorage = 0,
                    RawEnterprise = 0,
                    RawNav1 = 0,
                    RawNav2 = 0,
                    RawMesOut1 = 0,
                    RawMesOut2 = 0
                };

                MaterialNamesList.Add(dl.MaterialCode);

                MaterialDataList.Add(mm);
            }

            MaterialNamesList = MaterialNamesList.Distinct().ToList();

            var ConsList = fm.ConsDataList;    //fm.edb.fn_select_FactConsurmption().Where(x => MaterialNamesList.Contains(x.MaterialCode)).ToList();


            foreach (var m in MaterialDataList)
            {
                m.CurrentConsumption = ConsList.Where(x => x.MaterialCode == m.MaterialCode).Sum(x => (double)x.Quantity);

                foreach (var cl in ConsList.Where(x => x.MaterialCode == m.MaterialCode).ToList())
                {
                    Raws r = new Raws
                    {
                        Storage = cl.JobName,
                        Quantity = (double)cl.Quantity
                    };

                    m.ConsurmptionList.Add(r);
                }

                var sbList8 = fm.StockBaklanceList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == m.MaterialCode)).ToList();
                //   sbList8 = sbList8.Where(x => fm.StorageNamesList.Contains(x.Storage)).ToList();

                var sbpList8 = fm.StockBalancesPlantList.Where(x => (x.WorkDate.Date == DateTime.Today.Date) && (x.WorkDate.Hour >= 6) && (x.WorkDate.Hour <= 10) && (x.MaterialCode == m.MaterialCode)).ToList();
                sbpList8 = sbpList8.Where(x => fm.StorageNamesList.Contains(x.LineName)).ToList();

                var sbListIn = sbList8.Where(x => x.StorageType == "Собств").ToList();
                var sbListOut = sbList8.Where(x => x.StorageType != "Собств").ToList();

                var sbListInBlock = sbListIn.Where(x => x.TestQualityGr == "3Блок").ToList();
                var sbListOutBlock = sbListOut.Where(x => x.TestQualityGr == "3Блок").ToList();

                var sList = fm.StorageList.Where(x => x.MaterialCode == m.MaterialCode).ToList();
                var sListOld = sList.Where(x => x.PlanOperDate < DateTime.Today.Date).ToList();
                var sListNew = sList.Where(x => (x.PlanOperDate >= DateTime.Today.Date) && ((x.StatusMZP.ToLower().Trim() == "заказано") || (x.StatusMZP.ToLower().Trim() == ""))).ToList();

                m.RawStorage = (double)(sbList8.Sum(x => x.Quantity) - sbListInBlock.Sum(x => x.Quantity) /*- sbListOutBlock.Sum(x => x.Quantity)*/);
                m.RawEnterprise = (double)sbpList8.Sum(x => x.Quantity);
                m.RawNav1 = (double)sbListOut.Sum(x => x.Quantity);
                m.RawNav2 = (double)sbListIn.Sum(x => x.Quantity);
                m.RawMesOut1 = (double)sbListOutBlock.Sum(x => x.Quantity);
                m.RawMesOut2 = (double)sbListInBlock.Sum(x => x.Quantity);
                m.OldTaskNav = sListOld.Sum(x => (x.PlanQuantity - x.FactQuantity));

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

                        m.RawStorageList.Add(r);
                    }
                }

                try
                {
                    DateTime DT = m.RawStorageList.Min(x => x.BBFDate);

                    if ((DT - DateTime.Today.Date).Days > 30)
                    {
                        m.Color1 = "Green";
                    }
                    else if ((DT - DateTime.Today.Date).Days > 20)
                    {
                        m.Color1 = "Magenta";
                    }
                    else if ((DT - DateTime.Today.Date).Days > 0)
                    {
                        m.Color1 = "DarkOrange";
                    }
                    else
                    {
                        m.Color1 = "Red";
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

                    m.RawEnterpriseList.Add(r);
                }

                try
                {
                    DateTime DT = m.RawEnterpriseList.Min(x => x.BBFDate);

                    if ((DT - DateTime.Today.Date).Days > 30)
                    {
                        m.Color2 = "Green";
                    }
                    else if ((DT - DateTime.Today.Date).Days > 20)
                    {
                        m.Color2 = "Magenta";
                    }
                    else if ((DT - DateTime.Today.Date).Days > 0)
                    {
                        m.Color2 = "DarkOrange";
                    }
                    else
                    {
                        m.Color2 = "Red";
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

                    m.RawNav1List.Add(r);
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

                    m.RawNav2List.Add(r);
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

                    m.RawMesOut1List.Add(r);
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

                    m.RawMesOut2List.Add(r);
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
                            Quantity = rs6.PlanQuantity - rs6.FactQuantity
                        };

                        m.OldTaskNavList.Add(r);
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
                            AlterDate = new DateTime(rs7.PlanOperDate.Year, rs7.PlanOperDate.Month, 1),
                            AddDate = rs7.PlanOperDate.AddDays(1)
                        };

                        if (rs7.StatusMZP.Trim() == "")
                        {
                            m.TaskNavList1.Add(r);
                        }
                        else
                        {
                            m.TaskNavList2.Add(r);
                        }
                    }
                }

            }


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void ShowData()
        {
            dataGridViewMain.DataSource = null;
            dataGridViewMain.Rows.Clear();
            dataGridViewMain.Columns.Clear();
            WorkMonthList.Clear();
            MaterialData = new List<DailyResult>();

            string[] NewString;
            char RowSplitter = '|';
            string sVal = "";
            DateTime CurrDate;
            bool fl;
            Int32 J;
            double Value = 0;

            CFMListGray = new List<string>();

            DateTime XlDataStart =  XLDataList.Min(x => x.DateWork);
            XlDataStart = new DateTime(XlDataStart.Year, XlDataStart.Month, 1);
            DateTime XlDataEnd = new DateTime(DateTime.Today.Year, DateTime.Today.Month, 1);
            XlDataEnd = XlDataEnd.AddMonths(1).AddDays(-1);

            XLDataList = XLDataList.Where(x => MaterialNamesList.Contains(x.ProdCode) && (x.Quantity > 0)).ToList();

            while (XlDataStart < XlDataEnd)
            {
                CurrDate = XlDataStart.AddMonths(1).AddDays(-1);

                foreach (var mat in MaterialDataList)  //неликвид! материалы не используются по плану в данный момент! 
                {
                    var xData = XLDataList.Where(x => (x.DateWork >= XlDataStart) && (x.DateWork <= CurrDate) && (x.ProdCode == mat.MaterialCode)).ToList();
                    Value = xData.Sum(x => x.Quantity);

                    mat.OutList.Add(new Raws
                    {
                        ProdDate = XlDataStart,
                        Quantity = Value
                    });

                    //факт
                    var yData = fm.tdb.fn_select_ProductionByCodeAndPeriod(mat.MaterialCode, XlDataStart.AddHours(8), CurrDate.AddDays(1).AddHours(8)).ToList();
                    Value = yData.Sum(x => (double)x.Quantity);

                    mat.FactList.Add(new Raws
                    {
                        ProdDate = XlDataStart,
                        Quantity = Value
                    });
                }

                WorkMonthList.Add(XlDataStart);

                XlDataStart = CurrDate.AddDays(1);
            }

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
            foreach (var ddd in WorkMonthList)
            {
                dt1.Columns.Add("Потребность на " + ddd.ToString("MM.yyyy")+" план/факт");
            }
            dt1.Columns.Add("Уже потреблено");//8
            dt1.Columns.Add("Остатки склады");//9
            dt1.Columns.Add("Остатки пр-во");//10
            dt1.Columns.Add("Внешние остатки (Собств по НАВ)");//11
            dt1.Columns.Add("Заблокировано (Внешние МЕС)");//12
            dt1.Columns.Add("Старые заявки НАВ");//13
            dt1.Columns.Add("ИТОГО:");//14
            dt1.Columns.Add("ZZZZZ");

            for (int i = 0; i < MaterialDataList.Count; i++)
            {
                if ((MaterialDataList[i].WaitingDaysString == "") || (MaterialDataList[i].MaterialMultiplyString == "") /*||(MaterialDataList[i].StorageQuantityString=="")*/)
                {
                    CFMListGray.Add(MaterialDataList[i].MaterialCode);
                }

                Value = 0;

                sVal = MaterialDataList[i].MaterialCode + RowSplitter + MaterialDataList[i].MaterialName + RowSplitter +
                         MaterialDataList[i].MaterialGroup + RowSplitter + MaterialDataList[i].MaterialMultiplyString + RowSplitter +
                         MaterialDataList[i].WaitingDaysString + RowSplitter + MaterialDataList[i].StorageQuantityString + RowSplitter +
                         MaterialDataList[i].Responsible + RowSplitter + MaterialDataList[i].Sender + RowSplitter;

                for (int j = 0; j < WorkMonthList.Count; j++)
                {
                    if (Math.Abs(MaterialDataList[i].OutList[j].Quantity) <= 1)
                    {
                        sVal = sVal + MaterialDataList[i].OutList[j].Quantity.ToString("N1") + "/"; // + RowSplitter;
                       // Value = Value + MaterialDataList[i].OutList[j].Quantity;
                    }
                    else
                    {
                        sVal = sVal + MaterialDataList[i].OutList[j].Quantity.ToString("N0") + "/"; // + RowSplitter;
                       // Value = Value + MaterialDataList[i].OutList[j].Quantity;
                    }

                    if (Math.Abs(MaterialDataList[i].FactList[j].Quantity) <= 1)
                    {
                        sVal = sVal + MaterialDataList[i].FactList[j].Quantity.ToString("N1")  + RowSplitter;
                       // Value = Value + MaterialDataList[i].FactList[j].Quantity;
                    }
                    else
                    {
                        sVal = sVal + MaterialDataList[i].FactList[j].Quantity.ToString("N0")  + RowSplitter;
                       // Value = Value + MaterialDataList[i].FactList[j].Quantity;
                    }
                }

                if (Math.Abs(MaterialDataList[i].CurrentConsumption) <= 1)  //уже потреблено
                {
                    sVal = sVal + MaterialDataList[i].CurrentConsumption.ToString("N1") + RowSplitter;
                    Value = Value + MaterialDataList[i].CurrentConsumption;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].CurrentConsumption.ToString("N0") + RowSplitter;
                    Value = Value + MaterialDataList[i].CurrentConsumption;
                }

                if (Math.Abs(MaterialDataList[i].RawStorage) <= 1)  //склад
                {
                    sVal = sVal + MaterialDataList[i].RawStorage.ToString("N1") + RowSplitter;
                    Value = Value + MaterialDataList[i].RawStorage;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].RawStorage.ToString("N0") + RowSplitter;
                    Value = Value + MaterialDataList[i].RawStorage;
                }

                if (Math.Abs(MaterialDataList[i].RawEnterprise) <= 1)  //пр-во
                {
                    sVal = sVal + MaterialDataList[i].RawEnterprise.ToString("N1") + RowSplitter;
                    Value = Value + MaterialDataList[i].RawEnterprise;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].RawEnterprise.ToString("N0") + RowSplitter;
                    Value = Value + MaterialDataList[i].RawEnterprise;
                }

                if (Math.Abs(MaterialDataList[i].RawNav1) <= 1)  //внешние
                {
                    sVal = sVal + MaterialDataList[i].RawNav1.ToString("N1") + " (";
                    Value = Value + MaterialDataList[i].RawNav1;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].RawNav1.ToString("N0") + " (";
                    Value = Value + MaterialDataList[i].RawNav1;
                }

                if (Math.Abs(MaterialDataList[i].RawNav2) <= 1)   //собств
                {
                    sVal = sVal + MaterialDataList[i].RawNav2.ToString("N1") + ") " + RowSplitter;
                 //   Value = Value + MaterialDataList[i].RawNav2;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].RawNav2.ToString("N0") + ") " + RowSplitter;
                   // Value = Value + MaterialDataList[i].RawNav2;
                }

                if (Math.Abs(MaterialDataList[i].RawMesOut2) <= 1)  //забл
                {
                    sVal = sVal + MaterialDataList[i].RawMesOut2.ToString("N1") + " (";
                    Value = Value - MaterialDataList[i].RawMesOut2;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].RawMesOut2.ToString("N0") + " (";
                    Value = Value - MaterialDataList[i].RawMesOut2;
                }

                if (Math.Abs(MaterialDataList[i].RawMesOut1) <= 1)  //внешние МЕС
                {
                    sVal = sVal + MaterialDataList[i].RawMesOut1.ToString("N0") + ") " + RowSplitter;
                    Value = Value - MaterialDataList[i].RawMesOut1;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].RawMesOut1.ToString("N0") + ") " + RowSplitter;
                    Value = Value - MaterialDataList[i].RawMesOut1;
                }

                if (Math.Abs(MaterialDataList[i].OldTaskNav) <= 1)
                {
                    sVal = sVal + MaterialDataList[i].OldTaskNav.ToString("N1") + RowSplitter;   //0-12
                    Value = Value + MaterialDataList[i].OldTaskNav;
                }
                else
                {
                    sVal = sVal + MaterialDataList[i].OldTaskNav.ToString("N0") + RowSplitter;   //0-12
                    Value = Value + MaterialDataList[i].OldTaskNav;
                }

                if (Math.Abs(Value) <=1)
                {
                    sVal = sVal + Value.ToString("N1") + RowSplitter;
                }
                else
                {
                    sVal = sVal + Value.ToString("N0") + RowSplitter;
                }


                NewString = sVal.Split(RowSplitter);
                dt1.Rows.Add(NewString);
            }

            bs = new BindingSource();
            bs.DataSource = dt1;
            dataGridViewMain.DataSource = bs;

            dataGridViewMain.Columns[0].Width = 80;
            dataGridViewMain.Columns[1].Width = 300;
            dataGridViewMain.Columns[2].Width = 120;
            dataGridViewMain.Columns[6].Width = 100;
            dataGridViewMain.Columns[7].Width = 100;
            dataGridViewMain.Columns[dataGridViewMain.ColumnCount - 1].Visible = false;

            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;

            foreach (DataGridViewColumn c in dataGridViewMain.Columns)
            {
                c.SortMode = DataGridViewColumnSortMode.NotSortable;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void DGV_Get_CellColor()
        {
            Int32 Offset = WorkMonthList.Count;
            //cells values
            Font f1 = new Font(dataGridViewMain.DefaultCellStyle.Font, FontStyle.Bold);
            //1-2 cells
            for (int j = 0; j < dataGridViewMain.Rows.Count; j++)
            {
                string PC = dataGridViewMain.Rows[j].Cells[0].Value.ToString();   //dataGridViewMain.Rows[j].Cells[0].Value.ToString();

                if (URSCodeList.Contains(PC))                   //(dataGridViewMain.Rows[j].Cells["Код материала"].Value.ToString().ToLower() == URSList[i].MaterialCode.ToLower())
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = SColor;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = SColor;
                }
                else if (URMCodeList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = MColor;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = MColor;
                }
                else if (fm.OPRList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightCyan; //  .SteelBlue;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightCyan;
                }
                /*  else if (CFMListGold.Contains(PC))
                  {
                      if (CFMList[j].MaterialGroup != "ГП") //(CFMBool[j] == false) 
                      {
                          dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightYellow;   //.Gold;
                          dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightYellow;   //.Gold;
                      }
                  }*/
                else if (CFMListGray.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGray;  //DarkGray
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGray;
                }

               /* if (CFMListRed.Contains(PC)) //-
                {
                    var CM = CFMListMinus.Where(x => x.ProdCode == PC).ToList();
                    if (CM.Count > 0)
                    {
                        foreach (var c in CM[0].ColNo)
                        {
                            dataGridViewMain.Rows[j].Cells[c].Style.ForeColor = Color.Red; // Color.DarkRed;
                            dataGridViewMain.Rows[j].Cells[c].Style.Font = f1;
                        }
                    }
                }*/

               /* if (CFMListGreen.Contains(PC))  //+
                {
                    var CP = CFMListPlus.Where(x => x.ProdCode == PC).ToList();
                    if (CP.Count > 0)
                    {
                        foreach (var c in CP[0].ColNo)
                        {
                            dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightGreen;
                        }
                    }
                }*/

               /* if (CFMListPink.Contains(PC))
                {
                    var CL = CFMListLine.Where(x => x.ProdCode == PC).ToList();
                    if (CL.Count > 0)
                    {
                        foreach (var c in CL[0].ColNo)
                        {
                            dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightPink;    //.LightPink;
                        }
                    }
                }*/

               /* if (CFMListGoldens.Contains(PC))
                {
                    var CLP = CFMListLinePlus.Where(x => x.ProdCode == PC).ToList();
                    if (CLP.Count > 0)
                    {
                        foreach (var c in CLP[0].ColNo)
                        {
                            dataGridViewMain.Rows[j].Cells[c].Style.BackColor = Color.LightYellow;   //.Gold;
                        }
                    }
                }*/

                /*if (TaskList.Contains(PC))
                {
                    dataGridViewMain.Rows[j].Cells[0].Style.BackColor = Color.LightGreen;   //.Lime;
                    dataGridViewMain.Rows[j].Cells[1].Style.BackColor = Color.LightGreen;
                }*/

                //9/10 color
               /* try
                {
                    var cfm = XLDataList.Where(x => x.ProdCodeStr == PC).First();
                   
                        dataGridViewMain.Rows[j].Cells[9 + Offset].Style.ForeColor = Color.FromName(cfm.Color1);
                        dataGridViewMain.Rows[j].Cells[10 + Offset].Style.ForeColor = Color.FromName(cfm.Color2);
                   
                }
                catch (Exception xxx)
                { }*/

                //8-13 headers

                for (int i = 8; i <= 13; i++)
                {
                    dataGridViewMain.Columns[i + Offset].HeaderCell.Style.Font = new Font(dataGridViewMain.ColumnHeadersDefaultCellStyle.Font.FontFamily, 9f, FontStyle.Bold);
                }

            }
            //  DGV_UpdateColumnZero();
        }

       /* private void DrawDGV()
        {

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }*/

        private void FormUnliquid_Activated(object sender, EventArgs e)
        {
          //  DrawDGV();
        }

        private void FormUnliquid_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }

        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
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
                Int32 CurCol = 14 + WorkMonthList.Count;
                DateTime WorkDate;
                List<DateTime> WorkDateList = new List<DateTime>();
                Int32 Interval;

                Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                wSheet.Name = "Остатки НЕЛИКВИД";
                Excel1.Range range1;

                List<DateTime> DList = new List<DateTime>();
                DList = DTList.Distinct().OrderBy(x => x.Date).ToList();

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

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, dataGridViewMain.Columns.Count ]];
                range1.Font.Bold = true;
                range1.WrapText = true;
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[dataGridViewMain.Rows.Count + 1, 1]];
                range1.NumberFormat = "@";

                for (int i = 0; i < dataGridViewMain.Rows.Count; i++)
                {
                    wSheet.Cells[i + 2, 1] = dataGridViewMain.Rows[i].Cells[0].Value.ToString();
                    wSheet.Cells[i + 2, 2] = dataGridViewMain.Rows[i].Cells[1].Value.ToString();
                    wSheet.Cells[i + 2, 3] = dataGridViewMain.Rows[i].Cells[2].Value.ToString();
                    wSheet.Cells[i + 2, 4] = dataGridViewMain.Rows[i].Cells[3].Value.ToString();
                    wSheet.Cells[i + 2, 5] = dataGridViewMain.Rows[i].Cells[4].Value.ToString();
                    wSheet.Cells[i + 2, 6] = dataGridViewMain.Rows[i].Cells[5].Value.ToString();
                    wSheet.Cells[i + 2, 7] = dataGridViewMain.Rows[i].Cells[6].Value.ToString();
                    wSheet.Cells[i + 2, 8] = dataGridViewMain.Rows[i].Cells[7].Value.ToString();

                    for (int j = 0; j < Cols; j++)
                    {
                        wSheet.Cells[i + 2, 9 + j] = dataGridViewMain.Rows[i].Cells[8 + j].Value.ToString();      //Math.Round(CFMList[i].OutList[j].Quantity, 0);
                    }

                    wSheet.Cells[i + 2, 9 + Cols] =fm.ConvertStringToDouble( dataGridViewMain.Rows[i].Cells[8 + Cols].Value.ToString());
                    wSheet.Cells[i + 2, 10 + Cols] = fm.ConvertStringToDouble(dataGridViewMain.Rows[i].Cells[9 + Cols].Value.ToString());
                    wSheet.Cells[i + 2, 11 + Cols] = fm.ConvertStringToDouble(dataGridViewMain.Rows[i].Cells[10 + Cols].Value.ToString());
                    wSheet.Cells[i + 2, 12 + Cols] = fm.ConvertStringToDouble(dataGridViewMain.Rows[i].Cells[11 + Cols].Value.ToString());
                    wSheet.Cells[i + 2, 13 + Cols] = fm.ConvertStringToDouble(dataGridViewMain.Rows[i].Cells[12 + Cols].Value.ToString());
                    wSheet.Cells[i + 2, 14 + Cols] = fm.ConvertStringToDouble(dataGridViewMain.Rows[i].Cells[13 + Cols].Value.ToString());
                    wSheet.Cells[i + 2, 15 + Cols] = fm.ConvertStringToDouble(dataGridViewMain.Rows[i].Cells[14 + Cols].Value.ToString());

                    Value = 0;                   
                }  //остаток

                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[dataGridViewMain.Rows.Count + 1, dataGridViewMain.Columns.Count ]];
                range1.EntireColumn.AutoFit();

                wSheet = null;
                wBook = null;
                range1 = null;
                //   xlApp.Quit();
                xlApp = null;

                MessageBox.Show("Отчет сформирован", "Сообщение системы");
            }
        }

        private void FormUnliquid_Shown(object sender, EventArgs e)
        {
            DGV_Get_CellColor();

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            List<DailyResult> FilteredMaterialData1 = new List<DailyResult>();
            bool fl = false;
            string Code = "";

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
                for (int i = 0; i < MaterialDataList.Count; i++)
                {
                    if (MaterialDataList[i].MaterialCode == Code)
                    {
                        RInd = i;
                        i = MaterialDataList.Count + 1;
                    }
                }
                SelRowIndes = RInd;
                dataGridViewMain.FirstDisplayedScrollingRowIndex = RInd;
                dataGridViewMain.Rows[SelRowIndes].Selected = true;

                button5.Visible = true;
                button6.Visible = true;
                button9.Visible = true;

                ShowInfo(Code, SelRowIndes);
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


        private void ShowInfo(string ProductCode, Int32 RowInd)
        {
            Cursor = Cursors.WaitCursor;


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            Cursor = Cursors.Default;
        }

        private void dataGridViewMain_CellClick(object sender, DataGridViewCellEventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            Int32 RowInd = e.RowIndex;
            Int32 ColIndex = e.ColumnIndex;

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

                dgvPlanned.Rows.Clear();
                dgvFactis.Rows.Clear();
                dgvStorages.Rows.Clear();

                //storages
                {
                    var LocSBList = SBList.Where(x => x.MaterialCode == ProductCode).ToList();
                    //  LocSBList = LocSBList.Where(x => x.DayPeriod == "день").ToList();
                    var LocSBMList = SBMList.Where(x => x.MaterialCode == ProductCode).ToList();
                    //    LocSBMList = LocSBMList.Where(x => x.DayPeriod == "день").ToList();
                    var LocSBPList = SBPList.Where(x => x.MaterialCode == ProductCode).ToList();
                    //    LocSBPList = LocSBPList.Where(x => x.DayPeriod == "день").ToList();
                    var sbmdList = LocSBMList.Where(d => (d.WorkDate.Hour >= 6) && (d.WorkDate.Hour <= 10)).ToList();
                    var sbmnList = LocSBMList.Where(d => (d.WorkDate.Hour >= 18) && (d.WorkDate.Hour <= 22)).ToList();

                    var sbrld = sbmdList.GroupBy(x => new { x.Storage, x.MaterialCode, x.MaterialLot }).
                                  Select(g => new StockBalancesResult
                                  {
                                      WorkDate = DateTime.Today.Date,
                                      Storage = g.Key.Storage.Trim(),
                                      MaterialCode = g.Key.MaterialCode.Trim(),
                                      LotName = g.Key.MaterialLot.Trim(),
                                      QuantityMesDay = g.Sum(x => x.Quantity),

                                  }).ToList();

                    var sbrln = sbmnList.GroupBy(x => new { x.Storage, x.MaterialCode, x.MaterialLot }).
                      Select(g => new StockBalancesResult
                      {
                          WorkDate = DateTime.Today.Date,
                          Storage = g.Key.Storage.Trim(),
                          MaterialCode = g.Key.MaterialCode.Trim(),
                          LotName = g.Key.MaterialLot.Trim(),
                          QuantityMesNight = g.Sum(x => x.Quantity)
                      }).ToList();
                    //mes
                    foreach (StockBalancesResult s in sbrld)
                    {
                        var m = new StockBalancesResult
                        {
                            WorkDate = s.WorkDate,
                            Storage = s.Storage,
                            LotName = s.LotName,
                            LotDescription = s.LotDescription,
                            MaterialCode = s.MaterialCode,
                            QuantityMesDay = s.QuantityMesDay
                        };

                        sbrMesList.Add(m);
                    }

                    foreach (var s in sbrln)
                    {
                        var m = new StockBalancesResult
                        {
                            WorkDate = s.WorkDate,
                            Storage = s.Storage,
                            LotName = s.LotName,
                            LotDescription = s.LotDescription,
                            MaterialCode = s.MaterialCode,
                            QuantityMesNight = s.QuantityMesNight
                        };

                        var sss = sbrMesList.Where(x => (x.Storage == m.Storage) && (x.MaterialCode == m.MaterialCode) && (x.LotName == m.LotName)).ToList();

                        if (sss.Count > 0)
                        {
                            foreach (var sl in sss)
                            {
                                sl.QuantityMesNight = m.QuantityMesNight;
                                sl.QuantityMesDelta = sl.QuantityMesDay - sl.QuantityMesNight;
                            }
                        }
                        else
                        {
                            m.QuantityMesDelta = (-1) * m.QuantityMesNight;
                            sbrMesList.Add(m);
                        }
                    }
                    //nav 
                    var sbList = LocSBList.Where(d => (d.WorkDate.Hour >= 18) && (d.WorkDate.Hour <= 22)).ToList();

                    var sbrl = sbList.GroupBy(x => new { x.Storage, x.MaterialCode, x.LotName }).
                                  Select(g => new StockBalancesResult
                                  {
                                      WorkDate = DateTime.Today.Date,
                                      Storage = g.Key.Storage.Trim(),
                                      MaterialCode = g.Key.MaterialCode.Trim(),
                                      LotName = g.Key.LotName.Trim(),
                                      QuantityStorageNight = (double)g.Sum(x => x.Quantity)
                                  }).ToList();

                    foreach (var s in sbrl)
                    {
                        var m = new StockBalancesResult
                        {
                            WorkDate = s.WorkDate,
                            Storage = s.Storage,
                            LotName = s.LotName,
                            LotDescription = s.LotDescription,
                            MaterialCode = s.MaterialCode,
                            QuantityStorageNight = s.QuantityStorageNight
                        };

                        var sss = sbrMesList.Where(x => (x.Storage == m.Storage) && (x.MaterialCode == m.MaterialCode) && (x.LotName == m.LotName)).ToList();

                        if (sss.Count > 0)
                        {
                            foreach (var sl in sss)
                            {
                                sl.QuantityStorageNight = m.QuantityStorageNight;
                                sl.QuantityStorageDay = m.QuantityStorageNight - sl.QuantityMesDelta;
                                sl.Storage = m.Storage;
                            }
                        }
                        else
                        {
                            m.QuantityStorageDay = m.QuantityStorageNight;
                            sbrMesList.Add(m);
                        }
                    }

                    //    if (cbMaterialDelay.Checked == false)
                    {
                        sbrMesList = sbrMesList.Where(x => x.QuantityStorageDay != 0).OrderBy(x => x.Storage).ThenBy(x => x.LotName).ToList();
                    }
                    /*   else
                       {
                           sbrMesList = sbrMesList.Where(x => (x.QuantityStorageDay != 0)&&(x.BBFDate>=DateTime.Today.Date)).OrderBy(x => x.Storage).ThenBy(x => x.LotName).ToList();
                       }*/

                    foreach (var s in sbrMesList)
                    {
                        List<StockBalances> sl = new List<StockBalances>();

                        sl = SBList.Where(x => (x.Storage == s.Storage) && (x.MaterialCode == s.MaterialCode) && (x.LotName == s.LotName)).ToList();

                        if (sl.Count > 0)
                        {
                            s.ProdDate = sl[sl.Count - 1].ProdDate;
                            s.BBFDate = sl[sl.Count - 1].BBFDate;
                            if ((s.BBFDate - DateTime.Today.Date).Days > 30)
                            {
                                s.Status = "Green";
                            }
                            else if ((s.BBFDate - DateTime.Today.Date).Days > 20)
                            {
                                s.Status = "Magenta";
                            }
                            else if ((s.BBFDate - DateTime.Today.Date).Days > 0)
                            {
                                s.Status = "DarkOrange";
                            }
                            else
                            {
                                s.Status = "Red";
                            }
                        }
                        else
                        {
                            s.Status = "GrayText";
                            s.ProdDate = DateTime.Today.Date.AddDays(-5);
                            s.BBFDate = DateTime.Today.Date.AddDays(-5);
                        }

                    }


                    sbrMesList = sbrMesList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList();

                    if (sbrMesList.Count > 0)
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

                        foreach (var ss in sbrMesList.OrderBy(x => x.ProdDate))
                        {
                            dataString = new string[]
                            {
                                    ss.Storage,
                                    ss.LotName,
                                    ss.QuantityStorageDay.ToString("N1"),
                                    ss.ProdDate.ToString("dd.MM.yyyy"),
                                    ss.BBFDate.ToString("dd.MM.yyyy"),
                                    ss.Status
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
                    sbpList = SBPList.Where(x => (x.MaterialCode == ProductCode) && (x.Quantity != 0)).ToList();

                    foreach (var s1 in sbpList)
                    {
                        //  DateTime pDzate = s1.ValidFrom;
                        DateTime pBDzate = s1.ValidTo;

                        if ((pBDzate - DateTime.Today.Date).Days > 30)
                        {
                            s1.Status = "Green";
                        }
                        else if ((pBDzate - DateTime.Today.Date).Days > 20)
                        {
                            s1.Status = "Magenta";
                        }
                        else if ((pBDzate - DateTime.Today.Date).Days > 0)
                        {
                            s1.Status = "DarkOrange";
                        }
                        else
                        {
                            s1.Status = "Red";
                        }
                    }

                    if (sbpList.Count > 0)
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

                        foreach (var ss in sbpList.OrderBy(x => x.WorkDate))
                        {
                            dataString = new string[]
                             {
                                    ss.LineName,
                                    ss.MaterialLotName,
                                    ss.Quantity.ToString("N1"),
                                    ss.ValidFrom.ToString("dd.MM.yyyy"),
                                    ss.ValidTo.ToString("dd.MM.yyyy"),
                                    ss.Status
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
                    //tasks

                    sListA = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim() == "")).OrderBy(x => x.PlanOperDate).ToList();
                    sListB = fm.StorageList.Where(x => (x.MaterialCode == ProductCode) && (x.StatusMZP.Trim().ToLower() != "")).OrderBy(x => x.PlanOperDate).ToList();

                    var sListBB = sListB.Where(x => /*(x.PlanOperDate.Date >= DateTime.Today.Date) && */(x.StatusMZP.Trim().ToLower() == "заказано")).ToList();

                    //containers
                    var cList = fm.ContList.Where(x => (x.MaterialCode == ProductCode)).OrderBy(x => x.ExpDT).ToList();
                    cList = cList.Where(x => x.ExpDT >= DateTime.Today.Date).ToList();

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
                                    (s.PlanQuantity-s.FactQuantity).ToString("N1"),   //FactQuantity
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
                                    (s.PlanQuantity - s.FactQuantity).ToString("N1"),  //FactQuantity
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

                //planned, fact etc
                dataString = new string[] { "", "", "", "Данный компонент используется в:" };
                dgvPlanned.Rows.Add(dataString);
                dataString = new string[] { "", "", "", "", "", "Данный компонент используется в:" };
                dgvFactis.Rows.Add(dataString);

                /*   DList = fm.ExcelDataList.Select(x => x.DateWork).ToList();*/
                List<DateTime> DList = new List<DateTime>();
                DList = DTList.Distinct().OrderBy(x => x.Date).ToList();

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
                DDList = DList.Where(x => x.Date <= DateTime.Today.Date).OrderBy(x => x.Date).ToList();
                foreach (DateTime dt1 in DDList)
                {
                    //planned
                    {
                        fl = dataGridViewMain.Rows[RowInd].Cells["Класс материала"].Value.ToString().ToLower() == "тара" ? true : false;
                        var ELDListDate = fm.ExcelDataList.Where(x => (x.DateWork == dt1) && (x.Quantity > 0)).ToList();    // XLDataList.Where(x => (x.DateWork == CurrDate) && (x.Quantity > 0)).ToList();  //dt1

                        foreach (var ELD in ELDListDate)
                        {
                            var CP = CurProduct.Where(x => (x.ProductCode == ELD.ProdCode) && (x.LineNumber == ELD.Line) && (x.WorkDate <= dt1)).ToList();  //(x.Prod_No == ELD.ProdCode) && (x.WorkCenter == ELD.Line)
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
                                        var DL = LL[0].RecList.Where(x => x.WorkDateStart <= dt1).OrderByDescending(x => x.WorkDateStart).ToList();  //CurrDate
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

                        foreach (DataGridViewRow r in dgvPlanned.Rows)
                        {
                            if (r.Cells[0].Value.ToString() != "")
                            {
                                r.Cells[0].ToolTipText = "Количество: " + r.Cells[4].Value.ToString() + "; Норма на 1 кг: " + r.Cells[5].Value.ToString() + " Итого расход: " + r.Cells[6].Value.ToString();
                                r.Cells[1].ToolTipText = "Количество: " + r.Cells[4].Value.ToString() + "; Норма на 1 кг: " + r.Cells[5].Value.ToString() + " Итого расход: " + r.Cells[6].Value.ToString();
                                r.Cells[2].ToolTipText = "Количество: " + r.Cells[4].Value.ToString() + "; Норма на 1 кг: " + r.Cells[5].Value.ToString() + " Итого расход: " + r.Cells[6].Value.ToString();
                                //  r.Cells[3].ToolTipText = "Количество: " + r.Cells[4].Value.ToString() + "; Норма на 1 кг: " + r.Cells[5].Value.ToString() + " Итого расход: " + r.Cells[6].Value.ToString();
                            }
                        }
                        ELDListDate = null;
                    }  //planned
                }//foreach

                //fact
                {
                    fl = dataGridViewMain.Rows[RowInd].Cells["Класс материала"].Value.ToString().ToLower() == "тара" ? true : false;
                    //  var MuaSection = fm.MUAList.Where(x => (x.DateTime >= DateStart) && (x.DateTime < DateEnd)).ToList();    

                    var FactData = fm.tdb.fn_select_ProductionByCode(ProductCode, StartDate1).OrderBy(x => x.LotForErp).ToList();                      //fm.edb.fn_select_MaterialUsingByProduct(StartDate1, EndDate1, ProductCode).OrderBy(x => x.DateStart).ToList();

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

                    foreach (DataGridViewRow r in dgvFactis.Rows)
                    {
                        if (r.Cells[0].Value.ToString() != "")
                        {
                            r.Cells[0].ToolTipText = "Произведено: " + r.Cells[6].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                            r.Cells[1].ToolTipText = "Произведено: " + r.Cells[6].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                            r.Cells[2].ToolTipText = "Произведено: " + r.Cells[6].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                            //   r.Cells[3].ToolTipText = "Произведено: " + r.Cells[6].Value.ToString() + "; Итого расход: " + r.Cells[0].Value.ToString();
                        }
                    }

                    FactData = null;
                }
            }

            //Menu?

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
    }
}
