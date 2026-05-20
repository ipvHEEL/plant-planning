using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using NLog;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Excel1 = Microsoft.Office.Interop.Excel;
using PlantPlanning.Context;
using System.Runtime.InteropServices;
using System.Diagnostics;
using System.Data.SqlClient;
using System.Threading;
using System.IO;

namespace PlantPlanning
{
    public partial class FormMain : Form
    {
        public WorkPlanDataRawFKEntities tdb;
   //     public EDBEntities edb;
      //  public EquipPerfEntities eqp;
      //  public OLEG_DBEntities helg;
        public static Logger logger = LogManager.GetCurrentClassLogger();

        public List<string> LineList;
        public List<string> OPRList;
        public List<string> StorageNamesList;
        public List<string> SpisesCodeList;

        public List<DateTime> ExcelDatesList;
        public List<lData> ExcelDataList;
        public List<PlantOperList> POList;
        public List<ProductRecipes> ProdRecipeList;
        public List<ProductRecipes> SemiProdRecipeList;
        public List<ProductRecipes> ProdRecipeOPRList;

        public List<SelectStockBalances1_Result> StockBaklanceList;  //NAV
        public List<SelectStockBalancesMesBis_Result> StockBalancesMesList;
        public List<SelectStockBalancesPlantBis_Result> StockBalancesPlantList;
        public List<MesLocations> MesLocList;
        public List<NavPurchaseTable> StorageList;
        public List<tPlannedMeatContainers> ContList;
    //    public List<Material> mList;
        public List<tLines> LineDataList;
        public List<PlanOPZList> OPRTableList;
        public string ErrorMainStr = "";
        public string ErrorStr = "";
        public List<PR_GetRecipesDataMainProductBis_Result> lRec;
        public List<PR_GetRecipesDataMainProductBis_Result> lRec1;
        public List<PR_GetRecipesDataSemiProductBis_Result> NewRecList;
        public List<fn_select_PlannedMainData1_Result> MaterialDataList1;
        public List<tConsurptionData> ConsDataList;

        public List<ErrorGrid> CommonEGrid;
        public List<ErrorGrid> MainEGrid;
        public List<StatusesMaterials> SMList = new List<StatusesMaterials>();
        public List<MatNameAndAlterName> MMNANList = new List<MatNameAndAlterName>();

        public List<tPropertyTable> PTList = new List<tPropertyTable>();
        public List<PR_GetRecipesDataSemiProductBis_Result> SDataList;

        BindingSource bs;
        DataTable dt1;

        public string UserName;
        public char RowSplitter = '|';

        FormSettings fs;
        FormSaveCard fsc;
        FormLoadCard flc;
        FormAddCard fac;
        RecipesForm rf;
        FormLineList fll;
        StorageForm sf;
        Form_OPR fo;
        FormNavData fnd;
        MesSettingsForm msf;
        ShowForm sFm;
        Warning1 w1;
        FormUnliquid fu;
        FormCommdet fc;
        MercuryForm mf;
        ConsurmptionForm cf;
        WaitWindow WW;

      //  public string ConnectionString = "Data Source=192.168.91.162;Initial Catalog=OLEG DB;Persist Security Info=True;User ID=mes;Password=fdjn[eq";
        public string ConnectionStringTestData = "Data Source=192.168.91.162;Initial Catalog=WorkPlanDataRawFK;Persist Security Info=True;User ID=mes;Password=fdjn[eq";              // "Data Source=11-vm-dwh01;Initial Catalog=CSB_FK_REP;Persist Security Info=True;User ID=LinkMTZ;Password=mtz!Link2fk";    
        public SqlConnection connTest;

        private string FileName = "";
        private bool LoadFlag; 

        public FormMain()
        {
            InitializeComponent();
            LoadFlag = true;
            System.Threading.Thread thread = new System.Threading.Thread(WaitingShowForm);
            thread.Start();

            string td = GetConnectionString();

            if (!String.IsNullOrEmpty(td))
            {
                ConnectionStringTestData = td;
            }

            UserName = System.Environment.UserName;
            tdb = new WorkPlanDataRawFKEntities();

            connTest = new SqlConnection(ConnectionStringTestData);
            try
            {
                connTest.Open();
                connTest.Close();
            }
            catch (Exception xx)
            { }

            POList = new List<PlantOperList>();
            ConsDataList = new List<tConsurptionData>();
            MaterialDataList1 = new List<fn_select_PlannedMainData1_Result>();
            ExcelDataList = new List<lData>();
            ExcelDatesList = new List<DateTime>();
            lRec = new List<PR_GetRecipesDataMainProductBis_Result>();
            lRec1 = new List<PR_GetRecipesDataMainProductBis_Result>();
            NewRecList = new List<PR_GetRecipesDataSemiProductBis_Result>();
            LineList = new List<string>();
            logger.Info("WorkPlanning are started at " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss"));

            LoadRecipes();
            GetStorageList();
            LoadStoragesData();
            LoadData();

            try
            {
                LoadFlag = false;
                Thread.Sleep(250);
                thread.Abort();
            }
            catch (Exception xx)
            { }

            Cursor = Cursors.Default;
            this.BringToFront();
        }

        private string GetConnectionString()
        {
            string res = "";
            string FilePath = Application.StartupPath+@"\connection.udl";

            FileInfo f = new FileInfo(FilePath);
            string s = "";
            string Pass = "";
            string User = "";
            string DS = "";
            string IC = "";
            int Index = 0;

            if (f.Exists)
            {
                FileStream fs = new FileStream(FilePath, FileMode.Open);
                StreamReader sr = new StreamReader(fs);

                s = sr.ReadLine();
                s = sr.ReadLine();
                s = sr.ReadLine();

                sr.Close();

                Index = s.IndexOf("Password");
                Pass = s.Substring(Index);
                Index = Pass.IndexOf(";");
                Pass = Pass.Substring(0, Index);

                Index = s.IndexOf("User ID");
                User = s.Substring(Index);
                Index = User.IndexOf(";");
                User = User.Substring(0, Index);

                Index = s.IndexOf("Initial Catalog");
                IC = s.Substring(Index);
                Index = IC.IndexOf(";");
                IC = IC.Substring(0, Index);

                Index = s.IndexOf("Data Source");
                DS = s.Substring(Index);
                //    Index =DS.IndexOf(";");
                //    DS = DS.Substring(0, Index);

                res = DS + ";" + IC + ";Persist Security Info = True;" + User + ";" + Pass + ";";
            }
            
            f = null;
            GC.Collect();
            return res;            
        }

        private void WaitingShowFormWW()
        {
            try
            {
                WW = new WaitWindow();
                WW.ShowInTaskbar = false;
                WW.ShowDialog();
            }
            catch (Exception xx)
            {
                if (WW != null)
                {
                    WW.Close();
                }
                WW = null;
            }  

            while (LoadFlag == true)
            {
                System.Threading.Thread.Sleep(100);
            }

            WW.Close();
        }

        public BindingSource ConnectionSource(SqlConnection connection, string SqlString, string tname)
        {
            try
            {
                if (connection.State != ConnectionState.Open)
                {
                    connection.Open();
                }
                if (connection.State == ConnectionState.Open)
                {
                    SqlCommand comm1 = new SqlCommand(SqlString, connection);
                    comm1.CommandTimeout = 600;
                    SqlDataAdapter da1 = new SqlDataAdapter(comm1);
                    DataSet ds1 = new DataSet();
                    da1.Fill(ds1, tname);
                    BindingSource bs1 = new BindingSource { DataSource = ds1.Tables[tname] };

                    connection.Close();
                    return bs1;
                }
                else
                {
                    return null;
                }
            }
            catch (Exception xx)
            {
                return null;
            }
        }
        
        private void WaitingShowForm()
        {
            try
            {
                sFm = new ShowForm();
                sFm.ShowInTaskbar = false;
                sFm.ShowDialog();
                //   sFm.ShowDialog();
            }
            catch (Exception xx)
            {
                if (sFm != null)
                {
                    sFm.Close();
                }
                sFm = null;
            }

            while (LoadFlag == true)
            {
                Thread.Sleep(100);
            }

            sFm.Close();
        }

        public void GetStorageList()
        {
            var AList = tdb.MesAreas.Where(x => x.IsApproved == true).OrderBy(x=>x.AreaName).ToList();
            //MesLocList = new List<MesLocations>();

            /*  foreach (var al in AList)
              {
                  var MList = tdb.MesLocations.Where(x => (x.AreaID == al.ID) && (x.IsApproved == true)).ToList();
                  MesLocList.AddRange(MList);
              }*/

            StorageNamesList = AList.Select(x => x.AreaName).Distinct().ToList();
            
                //MesLocList.Select(x => x.LocationName).Distinct().ToList();
        }

        public void ControlAlarmWaste()
        {
            var WList = tdb.pr_GetAlarmWaste().ToList();

            if (WList!=null)
            {
                if (WList.Count > 0)
                {
                    dataGridView1.DataSource = null;

                    dataGridView1.Rows.Clear();
                    dataGridView1.Columns.Clear();

                    dataGridView1.Columns.Add("MaterialCode", "Код материала");
                    dataGridView1.Columns.Add("Name", "Наименование");
               //     dataGridView1.Columns.Add("Quantity", "Количество");
               //     dataGridView1.Columns.Add("L4Name", "Лот материала");
               //     dataGridView1.Columns.Add("MaterialLotName", "Альт.лот материала");

                    dataGridView1.EnableHeadersVisualStyles = false;
                    dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.ColumnHeadersDefaultCellStyle.Font.FontFamily, 16f, FontStyle.Bold | FontStyle.Italic);
                    dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;

                    for (int i=0; i<WList.Count; i++)
                    {
                        string[] s = new string[] { WList[i].MaterialCode,
                                                    WList[i].MaterialName/*,
                                                    WList[i].Quantity.ToString("N1"),
                                                    WList[i].L4Name,
                                                    WList[i].MaterialLotName*/ };

                        dataGridView1.Rows.Add(s);
                    }

                   Font f1 = new Font(dataGridView1.DefaultCellStyle.Font.FontFamily, 14f, FontStyle.Bold);


                    foreach (DataGridViewRow r in dataGridView1.Rows)
                    {
                        r.DefaultCellStyle.Font = f1;
                        r.DefaultCellStyle.BackColor = Color.Violet;
                    }

                   dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                  //  panel1.Enabled = false;
                  //  panel3.Enabled = false;
                }
            }

            GC.Collect();
        }

        public void LoadStoragesData()
        {
          //  Cursor = Cursors.WaitCursor;


            StockBaklanceList = new List<SelectStockBalances1_Result>(); //NAV!
            StockBaklanceList = tdb.SelectStockBalances1().ToList();
            StockBaklanceList = StockBaklanceList.Where(x => x.WorkDate.Date == DateTime.Today.Date).ToList();

            for (int i = 0; i < StockBaklanceList.Count; i++)
            {
                string s = StockBaklanceList[i].MaterialCode.Trim();
                var M = MMNANList.Where(x => x.MaterialCode == s).ToList();

                if (M.Count >0)
                {
                    if (M[0].MaterialCode != M[0].AlterMaterialCode)
                    {
                        StockBaklanceList[i].MaterialCode = M[0].AlterMaterialCode;
                    }
                }
            }

            DateTime DT = StockBaklanceList.Max(x => x.WorkDate);
            DT = DT.AddMinutes(-20);

            StockBaklanceList = StockBaklanceList.Where(x => x.WorkDate >= DT).ToList();

            StockBalancesMesList = new List<SelectStockBalancesMesBis_Result>();
            StockBalancesMesList = tdb.SelectStockBalancesMesBis().ToList();
            StockBalancesMesList = StockBalancesMesList.Where(x => x.WorkDate.Date == DateTime.Today.Date).ToList();

            DT = StockBalancesMesList.Max(x => x.WorkDate);
            DT = DT.AddMinutes(-20);

            StockBalancesMesList = StockBalancesMesList.Where(x => x.WorkDate >= DT).ToList();

            for (int i = 0; i<StockBalancesMesList.Count; i++)
            {
                string s = StockBalancesMesList[i].MaterialCode.Trim();
                var M = MMNANList.Where(x => x.MaterialCode == s).ToList();

                if (M.Count > 0)
                {
                    if (M[0].MaterialCode != M[0].AlterMaterialCode)
                    {
                        StockBalancesMesList[i].MaterialCode = M[0].AlterMaterialCode;
                    }
                }
            }

            StockBalancesPlantList = new List<SelectStockBalancesPlantBis_Result>();
            StockBalancesPlantList = tdb.SelectStockBalancesPlantBis().ToList();
            StockBalancesPlantList = StockBalancesPlantList.Where(x => x.WorkDate.Date == DateTime.Today.Date).ToList();

            DT = StockBalancesPlantList.Max(x => x.WorkDate);
            DT = DT.AddMinutes(-20);

            StockBalancesPlantList = StockBalancesPlantList.Where(x => x.WorkDate >= DT).ToList();

            for (int i = 0; i < StockBalancesPlantList.Count; i++)
            {
                string s = StockBalancesPlantList[i].MaterialCode.Trim();
                var M = MMNANList.Where(x => x.MaterialCode == s).ToList();

                if (M.Count > 0)
                {
                    if (M[0].MaterialCode != M[0].AlterMaterialCode)
                    {
                        StockBalancesPlantList[i].MaterialCode = M[0].AlterMaterialCode;
                    }
                }
            }

            try
            {
                StorageList = tdb.NavPurchaseTable.ToList();  //.Where(x => x.StatusMZP.ToLower().Trim() == "заказано")
            }
            catch (Exception xx)
            {
                StorageList = new List<NavPurchaseTable>();
            }
            StorageList = StorageList.Where(x => x.DateWork.Date == DateTime.Today.Date).ToList();

            DT = StorageList.Max(x => x.DateWork);
            DT = DT.AddMinutes(-20);

            StorageList = StorageList.Where(x => x.DateWork >= DT).ToList();
            StorageList = StorageList.Where(x => x.StatusMZP.ToLower().Trim() != "тмц на складе").ToList();
            StorageList = StorageList.Where(x => x.StatusMZP.ToLower().Trim() != "отказ").ToList();

            StorageList = StorageList.Where(x => x.PlanQuantity != x.FactQuantity).ToList();    //(x => x.PlanQuantity > 0)
            //  StorageList = StorageList.Where(x => x.NessDateOrder >= DateTime.Today.Date).ToList();

            GC.Collect();
        //    Cursor = Cursors.Default;
        }

        public void LoadData()
        {
            Cursor = Cursors.WaitCursor;
          //  mList = new List<Material>();
            OPRList = new List<string>();
            Material mm;
            string s;

            OPRList = tdb.PlanOPZList.Select(x => x.MatCode).ToList();

            try
            {
                ContList = tdb.tPlannedMeatContainers.ToList();
                ContList = ContList.Where(x => x.DateWork.Date == DateTime.Today.Date).ToList();

                foreach (var cl in ContList)
                {
                    cl.AlterDate = new DateTime(cl.ExpDT.Year, cl.ExpDT.Month, 1);
                }
            }
            catch (Exception xx)
            {
                ContList = new List<tPlannedMeatContainers>();
            }

            MaterialDataList1 = tdb.fn_select_PlannedMainData1().ToList();

            var DataList = MaterialDataList1.ToList();     //.Where(x => x.LASTPRODUCER != "НЕЛИКВИД").ToList();

            for (int i = 0; i < DataList.Count; i++)
            {
                mm = new Material();

                mm.MaterialCode = DataList[i].MaterialCode;
                mm.MaterialName = DataList[i].NavMaterialName;
                mm.MaterialGroup = DataList[i].MaterialGroup;
                mm.MaterialCategory = DataList[i].SIGname;
                mm.Responsible = DataList[i].SIGresp;
                mm.MaterialMultiply = (Int32)DataList[i].Multiplicity;
                mm.MaterialMultiplyString = mm.MaterialMultiply == 0 ? "" : mm.MaterialMultiply.ToString();
                mm.WaitingDays = (Int32)DataList[i].MinTransferDays;
                mm.WaitingDaysString = mm.WaitingDays == 0 ? "" : mm.WaitingDays.ToString();
                mm.StorageQuantity = (Int32)DataList[i].SafetyStock;
                mm.StorageQuantityString = mm.StorageQuantity == 0 ? "" : mm.StorageQuantity.ToString("0");
                mm.Sender = DataList[i].LASTPRODUCER;

                mm.ListUBP = new List<UsingByPlans>();
                //   mm.ListRec = new List<Recipe>();
                mm.ListSQBD = new List<StorageQuantByDate>();
                mm.ListUBD = new List<UsingByDate>();

                var SM = SMList.Where(x => x.MaterialCode == mm.MaterialCode).ToList();

                if (SM.Count>0)
                {
                    mm.BOM = SM[0].BOM;

                    if (mm.BOM == "N/A")
                    {
                        mm.MaterialStatus = 9;
                    }
                    else
                    {
                        mm.MaterialStatus = 0;
                    }
                }

              /*  mm.RawEnterpriseList = new List<Raws>();
                mm.RawStorageList = new List<Raws>();
                mm.RawNav1List = new List<Raws>();
                mm.RawNav2List = new List<Raws>();
                mm.RawMesOut1List = new List<Raws>();
                mm.RawMesOut2List = new List<Raws>();
                mm.OldTaskNavList = new List<Raws>();
                mm.TaskNavList1 = new List<Raws>();
                mm.TaskNavList2 = new List<Raws>();
                mm.StorageQuantList = new List<Raws>();
                mm.OutList = new List<Raws>();
                mm.ConsurmptionList = new List<Raws>();

                mm.ListLine = new List<int>();
                mm.ListPlus = new List<int>();*/

             //   mList.Add(mm);
            }

            DataList = new List<fn_select_PlannedMainData1_Result>();
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            Cursor = Cursors.Default;
        }

        public string  ReplaceLineName(string OldName)
        {
            string NewName = "";

             try
             {
                 NewName = LineDataList.Where(x => x.LineCode.ToString() == OldName).First().LineName;
             }
             catch (Exception xx)
             {
                 NewName = OldName;
             }

           // NewName = OldName;

            return NewName;
        }

        public void LoadRecipes()
        {
            Recipe rr;
            RecipeList rl;
            ProductLinesRecipes plr;
            ProductRecipes pr;
            bool fl = false;

            DateTime wDate;
            string lNumber;
            string RCode;
            string MaterialCode = "";

            int alfa;
            int bravo;
            int charlie;            

            Cursor = Cursors.WaitCursor;
            ProdRecipeList = new List<ProductRecipes>();
            SemiProdRecipeList = new List<ProductRecipes>();
            ProdRecipeOPRList = new List<ProductRecipes>();

            PTList = tdb.tPropertyTable.ToList();

            ConsDataList = tdb.tConsurptionData.ToList();
            ConsDataList = ConsDataList.Where(x => x.WorkDate.Date == DateTime.Today.Date).ToList();

            OPRTableList = tdb.PlanOPZList.Where(x => (bool)x.IsBase == true).ToList();  //empty!!!!!!!
            List<string> OprComp = OPRTableList.Select(x => x.MatCode).ToList();

            List<PR_GetRecipesDataMainProductBis_Result> OPRRec = new List<PR_GetRecipesDataMainProductBis_Result>();
            List<PR_GetRecipesDataMainProductBis_Result> RDataList = new List<PR_GetRecipesDataMainProductBis_Result>();
            SDataList = new List<PR_GetRecipesDataSemiProductBis_Result>();

            try
            {
                RDataList = tdb.PR_GetRecipesDataMainProductBis().ToList();
                RDataList = RDataList.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();

                SDataList = tdb.PR_GetRecipesDataSemiProductBis().ToList();
                SDataList = SDataList.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();
            }
            catch (Exception xx)
            {
                LoadFlag = false;
                Thread.Sleep(200);
                MessageBox.Show("Нет связи с сервером SQL80" + System.Environment.NewLine + "Приложение будет закрыто", "Собщение системы");

                this.Close();
            }

            MMNANList = new List<MatNameAndAlterName>();
            MMNANList =  RDataList.GroupBy(x => new { x.MaterialCode, x.AlterMaterialCode, x.IPG })
                .Select(g => new MatNameAndAlterName
                {
                    AlterMaterialCode = g.Key.AlterMaterialCode,
                    MaterialCode = g.Key.MaterialCode,
                    MaterialGroup = g.Key.IPG
                }).ToList();

            var SDL = SDataList.GroupBy(x => new { x.MaterialCode, x.AlterMaterialCode, x.IPG })
                .Select(g => new MatNameAndAlterName
                {
                    AlterMaterialCode = g.Key.AlterMaterialCode,
                    MaterialCode = g.Key.MaterialCode,
                    MaterialGroup = g.Key.IPG
                }).ToList();

            //   SDL = SDL.Where(x => x.MaterialCode == "1021000706").ToList();

            MMNANList.AddRange(SDL);

            //MatStatus
            var RDL = RDataList.Where(x => x.Bom.Trim() != "").ToList();

            SMList = RDL.GroupBy(x => new { x.Bom, x.AlterMaterialCode })
                .Select(g => new StatusesMaterials {
                    BOM = g.Key.Bom,
                    MaterialCode = g.Key.AlterMaterialCode
                }).ToList();

           /* for (int zz =0; zz<RDL.Count; zz++)
            {
                StatusesMaterials sm = new StatusesMaterials();

                sm.BOM = RDL[zz].Bom.Trim();
                sm.MaterialCode = RDL[zz].MaterialCode;

                SMList.Add(sm);
            }*/

            SMList = SMList.Distinct().ToList();
            MMNANList = MMNANList.Distinct().ToList();

            for (int zz =0; zz<RDataList.Count; zz++)
            {
                PR_GetRecipesDataMainProductBis_Result res = new PR_GetRecipesDataMainProductBis_Result();

                res.AlterProductCode = RDataList[zz].AlterProductCode;
                res.Bom = RDataList[zz].Bom;
                res.IPG = RDataList[zz].IPG;
                res.LineNumber = RDataList[zz].LineNumber;
                res.MaterialCode = RDataList[zz].AlterMaterialCode;
                res.MaterialName = RDataList[zz].MaterialName;
                res.ProductCode = RDataList[zz].ProductCode;
                res.ProductName = RDataList[zz].ProductName;
                res.Quantity = RDataList[zz].Quantity;
                res.UnitMeasure = RDataList[zz].UnitMeasure;
                res.WorkDate = RDataList[zz].WorkDate;
                res.IsSemiProd = RDataList[zz].IsSemiProd;
                res.CreationDate = RDataList[zz].CreationDate;

                OPRRec.Add(res);
            }
            OPRRec = OPRRec.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();


            /*   try
               {
                   OPRRec = tdb.fn_select_RecipesView1().ToList();
                   OPRRec = OPRRec.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();
               }
               catch (Exception xx)
               {
                   LoadFlag = false;
                   Thread.Sleep(200);
                   MessageBox.Show("Нет связи с сервером SQL80" + System.Environment.NewLine + "Приложение будет закрыто", "Собщение системы");

                   this.Close();
               }*/

            OPRRec = OPRRec.Where(x => OprComp.Contains(x.ProductCode)).ToList();

            RCode = "";
            wDate = Convert.ToDateTime("2000-01-01");
            lNumber = "ZZZZZ";
            MaterialCode = "";

            for (int i = 0; i < OPRRec.Count; i++)
            {
                if (RCode != OPRRec[i].ProductCode)
                {
                    pr = new ProductRecipes();

                    pr.RecLineList = new List<ProductLinesRecipes>();
                    pr.ProductCode = OPRRec[i].ProductCode;
                    pr.AlterProductCode = OPRRec[i].ProductCode;
                    pr.ProductCodeString = pr.ProductCode;
                    pr.ProductName = OPRRec[i].ProductName;
                    RCode = pr.ProductCode;

                    ProdRecipeOPRList.Add(pr);

                    plr = new ProductLinesRecipes();
                    plr.RecList = new List<Recipe>();
                    lNumber =  OPRRec[i].LineNumber;
                    plr.LineNumber = OPRRec[i].LineNumber;

                    alfa = ProdRecipeOPRList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList.Add(plr);

                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)OPRRec[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    bravo = ProdRecipeOPRList[alfa].RecLineList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = OPRRec[i].IPG;
                    rl.MaterialCode = OPRRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = OPRRec[i].MaterialName;
                    rl.MaterialQuantity = (double)OPRRec[i].Quantity;
                    rl.IsSemiProd =(int) OPRRec[i].IsSemiProd;
                    rl.UMC = OPRRec[i].UnitMeasure;
                    rl.BOM = OPRRec[i].Bom;
                    charlie = ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (OPRRec[i].LineNumber != lNumber)
                {
                    plr = new ProductLinesRecipes();
                    plr.RecList = new List<Recipe>();
                    lNumber = OPRRec[i].LineNumber;
                    plr.LineNumber = lNumber;
                    alfa = ProdRecipeOPRList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList.Add(plr);

                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)OPRRec[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    bravo = ProdRecipeOPRList[alfa].RecLineList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = OPRRec[i].IPG;
                    rl.MaterialCode = OPRRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = OPRRec[i].MaterialName;
                    rl.MaterialQuantity = (double)OPRRec[i].Quantity;
                    rl.IsSemiProd = (int)OPRRec[i].IsSemiProd;
                    rl.UMC = OPRRec[i].UnitMeasure;
                    rl.BOM = OPRRec[i].Bom;
                    charlie = ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (wDate.Date != OPRRec[i].WorkDate.Value.Date)
                {
                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)OPRRec[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    alfa = ProdRecipeOPRList.Count - 1;
                    bravo = ProdRecipeOPRList[alfa].RecLineList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = OPRRec[i].IPG;
                    rl.MaterialCode = OPRRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = OPRRec[i].MaterialName;
                    rl.IsSemiProd = (int)OPRRec[i].IsSemiProd;
                    rl.MaterialQuantity = (double)OPRRec[i].Quantity;
                    rl.UMC = OPRRec[i].UnitMeasure;
                    rl.BOM = OPRRec[i].Bom;
                    charlie = ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Count - 1;
                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (MaterialCode != OPRRec[i].MaterialCode)
                {
                    rl = new RecipeList();

                    rl.IPG = OPRRec[i].IPG;
                    rl.MaterialCode = OPRRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = OPRRec[i].MaterialName;
                    rl.MaterialQuantity = (double)OPRRec[i].Quantity;
                    rl.IsSemiProd = (int)OPRRec[i].IsSemiProd;
                    rl.UMC = OPRRec[i].UnitMeasure;
                    rl.BOM = OPRRec[i].Bom;

                    alfa = ProdRecipeOPRList.Count - 1;
                    bravo = ProdRecipeOPRList[alfa].RecLineList.Count - 1;
                    charlie = ProdRecipeOPRList[alfa].RecLineList[bravo].RecList.Count - 1;

                    ProdRecipeOPRList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                    //  ProdRecipeList[ProdRecipeList.Count - 1].recList[ProdRecipeList[ProdRecipeList.Count - 1].recList.Count - 1].rList.Add(rl);
                }
            }

            //set dateend in RecipeList

            for (int i = 0; i < ProdRecipeOPRList.Count; i++)
            {
                for (int j = 0; j < ProdRecipeOPRList[i].RecLineList.Count; j++)
                {
                    for (int k = 0; k < ProdRecipeOPRList[i].RecLineList[j].RecList.Count - 1; k++)
                    {
                        try
                        {
                            ProdRecipeOPRList[i].RecLineList[j].RecList[k].WorkDateEnd = ProdRecipeOPRList[i].RecLineList[j].RecList[k + 1].WorkDateStart.AddDays(-1);
                        }
                        catch (Exception xx)
                        { }
                    }
                }
            }

            LineDataList = tdb.tLines.ToList();               //tdb.tMappingTable.ToList();
            lRec = new List<PR_GetRecipesDataMainProductBis_Result>();

            for (int zz = 0; zz < RDataList.Count; zz++)
            {
                PR_GetRecipesDataMainProductBis_Result res = new PR_GetRecipesDataMainProductBis_Result();

                res.AlterProductCode = RDataList[zz].AlterProductCode.Trim();
                res.Bom = RDataList[zz].Bom.Trim();
                res.IPG = RDataList[zz].IPG.Trim();
                res.LineNumber = RDataList[zz].LineNumber.Trim();
                res.MaterialCode = RDataList[zz].AlterMaterialCode.Trim();
                res.MaterialName = RDataList[zz].MaterialName.Trim();
                res.ProductCode = RDataList[zz].ProductCode.Trim();
                res.ProductName = RDataList[zz].ProductName.Trim();
                res.Quantity = (decimal)RDataList[zz].Quantity;
                res.IsSemiProd = RDataList[zz].IsSemiProd;
                res.UnitMeasure = RDataList[zz].UnitMeasure;
                res.WorkDate = RDataList[zz].WorkDate;
                res.CreationDate = RDataList[zz].CreationDate;

                lRec.Add(res);
            }

            lRec = lRec.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();

      //      var lll = lRec.Where(x => x.ProductCode == "1010031627").ToList();  //    1031018925

            lRec1 = new List<PR_GetRecipesDataMainProductBis_Result>();

            for (int zz = 0; zz < SDataList.Count; zz++)
            {
                PR_GetRecipesDataMainProductBis_Result res = new PR_GetRecipesDataMainProductBis_Result();

                res.AlterProductCode = SDataList[zz].AlterProductCode.Trim();
                res.Bom = SDataList[zz].Bom.Trim();
                res.IPG = SDataList[zz].IPG.Trim();
                res.LineNumber = SDataList[zz].LineNumber.Trim();
                res.MaterialCode = SDataList[zz].AlterMaterialCode.Trim();
                res.MaterialName = SDataList[zz].MaterialName.Trim();
                res.ProductCode = SDataList[zz].ProductCode.Trim();
                res.ProductName = SDataList[zz].ProductName.Trim();
                res.Quantity = (decimal)SDataList[zz].Quantity;
                res.IsSemiProd = SDataList[zz].IsSemiProd;
                res.UnitMeasure = SDataList[zz].UnitMeasure;
                res.WorkDate = SDataList[zz].WorkDate;
                res.CreationDate = SDataList[zz].CreationDate;

                lRec1.Add(res);
            }

            lRec1 = lRec1.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();

        //    var lll2 = lRec1.Where(x => x.MaterialCode == "4500000461").ToList();

        //    lRec1 = lRec1.Where(x => x.ProductCode == "1021000706").ToList();

            /*  try
              {
                  lRec = tdb.fn_select_RecipesView1().ToList();
                  lRec = lRec.OrderBy(x => x.ProductCode).ThenBy(x => x.LineNumber).ThenBy(x => x.WorkDate).ToList();                

                  //    lRec = lRec.Where(x => x.ProductCode == "1010012017").ToList();
              }
              catch (Exception xx)
              {
                  LoadFlag = false;
                  Thread.Sleep(200);
                  MessageBox.Show("Нет связи с сервером SQL80" + System.Environment.NewLine + "Приложение будет закрыто", "Собщение системы");
                  this.Close();
              }*/

            //   lRec = lRec.Where(x => x.ProductCode == "8ПФ00798").ToList();
            foreach (var l in lRec)
            {
                l.LineNumber = ReplaceLineName(l.LineNumber.Trim());
            }

            foreach (var l in lRec1)
            {
             //   if (l.MaterialCode =="1021002325" || l.ProductCode == "1021002325" )
             //   { }

                l.LineNumber = ReplaceLineName(l.LineNumber.Trim());
            }

            LineList = lRec.Select(x => x.LineNumber).ToList();
            LineList.AddRange(lRec1.Select(x => x.LineNumber).ToList());
            LineList = LineList.Distinct().ToList();
            LineList = LineList.OrderBy(x => x).ToList();
            
          /* for (int i=0; i<LineList.Count; i++)
            {
                try
                {
                    var Listdata = LineDataList.Where(x => x.LineName == LineList[i]).First();
                    LineList[i] = Listdata.ResultName;
                }
                catch (Exception xx)
                { }
            }*/
            LineList = LineList.Distinct().ToList();

            RCode = "";
            wDate = Convert.ToDateTime("2000-01-01");
            lNumber = "ZZZ";
            MaterialCode = "";

          //  var LLRec = lRec.Where(x => x.MaterialCode == "1031014495").ToList();

            for (int i = 0; i < lRec.Count; i++)
            {
                if (RCode != lRec[i].ProductCode)
                {
                    pr = new ProductRecipes();

                    pr.RecLineList = new List<ProductLinesRecipes>();
                    pr.ProductCode = lRec[i].ProductCode;
                    pr.AlterProductCode = lRec[i].AlterProductCode;
                    pr.ProductCodeString = pr.ProductCode;
                    pr.ProductName = lRec[i].ProductName;
                    RCode = pr.ProductCode;

                    ProdRecipeList.Add(pr);

                    plr = new ProductLinesRecipes();
                    plr.RecList = new List<Recipe>();
                    lNumber = lRec[i].LineNumber; // ReplaceLineName(lRec[i].LineNumber);                    
                    plr.LineNumber = lNumber;                    
                  
                    alfa = ProdRecipeList.Count - 1;
                    ProdRecipeList[alfa].RecLineList.Add(plr);

                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)lRec[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    bravo = ProdRecipeList[alfa].RecLineList.Count - 1;
                    ProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = lRec[i].IPG;
                    rl.MaterialCode = lRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec[i].Quantity;
                    rl.IsSemiProd = (int)lRec[i].IsSemiProd;
                    rl.UMC = lRec[i].UnitMeasure;
                    rl.BOM = lRec[i].Bom;

                    charlie = ProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                    ProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (lRec[i].LineNumber != lNumber)      // (ReplaceLineName(lRec[i].LineNumber) != lNumber)
                {
                    plr = new ProductLinesRecipes();
                    plr.RecList = new List<Recipe>();
                    lNumber = lRec[i].LineNumber;    // ReplaceLineName(lRec[i].LineNumber);
                    plr.LineNumber = lNumber;
                    alfa = ProdRecipeList.Count - 1;
                    ProdRecipeList[alfa].RecLineList.Add(plr);

                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)lRec[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    bravo = ProdRecipeList[alfa].RecLineList.Count - 1;
                    ProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = lRec[i].IPG;
                    rl.MaterialCode = lRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec[i].Quantity;
                    rl.IsSemiProd = (int)lRec[i].IsSemiProd;
                    rl.UMC = lRec[i].UnitMeasure;
                    rl.BOM = lRec[i].Bom;

                    charlie = ProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                    ProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (wDate.Date != lRec[i].WorkDate.Value.Date)
                {
                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)lRec[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    alfa = ProdRecipeList.Count - 1;
                    bravo = ProdRecipeList[alfa].RecLineList.Count - 1;
                    ProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = lRec[i].IPG;
                    rl.MaterialCode = lRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec[i].Quantity;
                    rl.IsSemiProd = (int)lRec[i].IsSemiProd;
                    rl.UMC = lRec[i].UnitMeasure;
                    rl.BOM = lRec[i].Bom;

                    charlie = ProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                    ProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (MaterialCode != lRec[i].MaterialCode)
                {
                    rl = new RecipeList();

                    rl.IPG = lRec[i].IPG;
                    rl.MaterialCode = lRec[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec[i].Quantity;
                    rl.IsSemiProd = (int)lRec[i].IsSemiProd;
                    rl.UMC = lRec[i].UnitMeasure;
                    rl.BOM = lRec[i].Bom;

                    alfa = ProdRecipeList.Count - 1;
                    bravo = ProdRecipeList[alfa].RecLineList.Count - 1;
                    charlie = ProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;

                    ProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                    //  ProdRecipeList[ProdRecipeList.Count - 1].recList[ProdRecipeList[ProdRecipeList.Count - 1].recList.Count - 1].rList.Add(rl);
                }
            }
          //  lRec.Clear();
            //set dateend in RecipeList

         //   var PR11List = ProdRecipeList.Where(x => x.ProductCode == "1010025786").ToList();

            for (int i = 0; i < ProdRecipeList.Count; i++)
            {
                for (int j = 0; j < ProdRecipeList[i].RecLineList.Count; j++)
                {
                    for (int k = 0; k < ProdRecipeList[i].RecLineList[j].RecList.Count - 1; k++)
                    {
                        try
                        {
                            ProdRecipeList[i].RecLineList[j].RecList[k].WorkDateEnd = ProdRecipeList[i].RecLineList[j].RecList[k+1].WorkDateStart.AddDays(-1);
                        }
                        catch (Exception xx)
                        { }
                    }
                }
            }

            //SemiProducts
            RCode = "";
            wDate = Convert.ToDateTime("2000-01-01");
            lNumber = "ZZZ";
            MaterialCode = "";

            for (int i = 0; i < lRec1.Count; i++)
            {
                if (RCode != lRec1[i].ProductCode)
                {
                    pr = new ProductRecipes();

                    pr.RecLineList = new List<ProductLinesRecipes>();
                    pr.ProductCode = lRec1[i].ProductCode;
                    pr.AlterProductCode = lRec1[i].AlterProductCode;
                    pr.ProductCodeString = pr.ProductCode;
                    pr.ProductName = lRec1[i].ProductName;
                    RCode = pr.ProductCode;

                    SemiProdRecipeList.Add(pr);

                    plr = new ProductLinesRecipes();
                    plr.RecList = new List<Recipe>();
                    lNumber = lRec1[i].LineNumber; // ReplaceLineName(lRec[i].LineNumber);                    
                    plr.LineNumber = lNumber;

                    alfa = SemiProdRecipeList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList.Add(plr);

                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)lRec1[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    bravo = SemiProdRecipeList[alfa].RecLineList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = lRec1[i].IPG;
                    rl.MaterialCode = lRec1[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec1[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec1[i].Quantity;
                    rl.IsSemiProd = (int)lRec1[i].IsSemiProd;
                    rl.UMC = lRec1[i].UnitMeasure;
                    rl.BOM = lRec1[i].Bom;

                    charlie = SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (lRec1[i].LineNumber != lNumber)      // (ReplaceLineName(lRec[i].LineNumber) != lNumber)
                {
                    plr = new ProductLinesRecipes();
                    plr.RecList = new List<Recipe>();
                    lNumber = lRec1[i].LineNumber;    // ReplaceLineName(lRec[i].LineNumber);
                    plr.LineNumber = lNumber;
                    alfa = SemiProdRecipeList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList.Add(plr);

                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)lRec1[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    bravo = SemiProdRecipeList[alfa].RecLineList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = lRec1[i].IPG;
                    rl.MaterialCode = lRec1[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec1[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec1[i].Quantity;
                    rl.IsSemiProd = (int)lRec1[i].IsSemiProd;
                    rl.UMC = lRec1[i].UnitMeasure;
                    rl.BOM = lRec1[i].Bom;

                    charlie = SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (wDate.Date != lRec1[i].WorkDate.Value.Date)
                {
                    rr = new Recipe();
                    rr.rList = new List<RecipeList>();
                    rr.WorkDateStart = (DateTime)lRec1[i].WorkDate;
                    wDate = rr.WorkDateStart;
                    alfa = SemiProdRecipeList.Count - 1;
                    bravo = SemiProdRecipeList[alfa].RecLineList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                    rl = new RecipeList();
                    rl.IPG = lRec1[i].IPG;
                    rl.MaterialCode = lRec1[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec1[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec1[i].Quantity;
                    rl.IsSemiProd = (int)lRec1[i].IsSemiProd;
                    rl.UMC = lRec1[i].UnitMeasure;
                    rl.BOM = lRec1[i].Bom;

                    charlie = SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                }
                if (MaterialCode != lRec1[i].MaterialCode)
                {
                    rl = new RecipeList();

                    rl.IPG = lRec1[i].IPG;
                    rl.MaterialCode = lRec1[i].MaterialCode;
                    MaterialCode = rl.MaterialCode;
                    rl.MaterialName = lRec1[i].MaterialName;
                    rl.MaterialQuantity = (double)lRec1[i].Quantity;
                    rl.IsSemiProd = (int)lRec1[i].IsSemiProd;
                    rl.UMC = lRec1[i].UnitMeasure;
                    rl.BOM = lRec1[i].Bom;

                    alfa = SemiProdRecipeList.Count - 1;
                    bravo = SemiProdRecipeList[alfa].RecLineList.Count - 1;
                    charlie = SemiProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;

                    SemiProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);
                    //  ProdRecipeList[ProdRecipeList.Count - 1].recList[ProdRecipeList[ProdRecipeList.Count - 1].recList.Count - 1].rList.Add(rl);
                }
            }

            for (int i = 0; i < SemiProdRecipeList.Count; i++)
            {
                for (int j = 0; j < SemiProdRecipeList[i].RecLineList.Count; j++)
                {
                    for (int k = 0; k < SemiProdRecipeList[i].RecLineList[j].RecList.Count - 1; k++)
                    {
                        try
                        {
                            SemiProdRecipeList[i].RecLineList[j].RecList[k].WorkDateEnd = ProdRecipeList[i].RecLineList[j].RecList[k + 1].WorkDateStart.AddDays(-1);
                        }
                        catch (Exception xx)
                        { }
                    }
                }
            }

            List<ProductRecipes> TBP =    tdb.tBlockedProduct.Where(x=>x.Using==false)
                .Select(x=> new ProductRecipes
                {
                    ProductCode=x.ProductCode,
                    ProductName=x.ProductName
                })
                .ToList();

            for (int i = ProdRecipeList.Count-1; i>=0; i--)
            {
                foreach (var t in TBP)
                {
                    if (ProdRecipeList.Count > 0)
                    {
                        if (ProdRecipeList[i].ProductCode == t.ProductCode)
                        {
                            ProdRecipeList.RemoveAt(i);
                        }
                    }
                }
            }

            for (int i = SemiProdRecipeList.Count - 1; i >= 0; i--)
            {
                foreach (var t in TBP)
                {
                    if (SemiProdRecipeList.Count > 0)
                    {
                        if (SemiProdRecipeList[i].ProductCode == t.ProductCode)
                        {
                            SemiProdRecipeList.RemoveAt(i);
                        }
                    }
                }
            }

            ProdRecipeList.AddRange(SemiProdRecipeList);

            for (int i = 0; i<RDataList.Count; i++)
            {
                PR_GetRecipesDataSemiProductBis_Result sp = new PR_GetRecipesDataSemiProductBis_Result();

                sp.AlterMaterialCode = RDataList[i].AlterMaterialCode;
                sp.AlterProductCode = RDataList[i].AlterProductCode;
                sp.Bom = RDataList[i].Bom;
                sp.IsSemiProd= RDataList[i].IsSemiProd;
                sp.IPG = RDataList[i].IPG;
                sp.LineNumber = RDataList[i].LineNumber;
                sp.MaterialCode = RDataList[i].MaterialCode;
                sp.MaterialName = RDataList[i].MaterialName;
                sp.ProductCode = RDataList[i].ProductCode;
                sp.ProductName = RDataList[i].ProductName;
                sp.UnitMeasure = RDataList[i].UnitMeasure;
                sp.CreationDate = RDataList[i].CreationDate;

                SDataList.Add(sp);
            }

            DateTime FrontDate = DateTime.Today.Date.AddDays(-30);

            var DDT1 = PTList.Where(x => x.PropertyName == "DayIntervalToNewRecipies").ToList();

            if (DDT1.Count > 0)
            {
                try
                {
                    FrontDate = DateTime.Today.Date.AddDays(-1 * Convert.ToInt32(DDT1[0].PropertyValue));
                }
                catch (Exception xx)
                { }
            }

            NewRecList = SDataList.Where(x => x.CreationDate > FrontDate).ToList();

            //    ProdRecipeList = ProdRecipeList.Where(x => x.ProductCode == "1021000706").ToList();

            //Export
            /*    Excel1.Application xlApp = new Excel1.Application();
                xlApp.Visible = true;
                object misValue = System.Reflection.Missing.Value;

                Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

                Excel1.Range range1;



                for (int i = 0; i<ProdRecipeList.Count; i++)
                {
                    wSheet.Cells[i+ 1, 1] = ProdRecipeList[i].ProductCodeString;
                    wSheet.Cells[i + 1, 2] = ProdRecipeList[i].ProductName;
                }*/

            //spises

            // SpisesCodeList = new List<string>();

            // var sData = tdb.fn_select_SpicesData().ToList();
            // SpisesCodeList = sData.Select(x => x.ProductCode).Distinct().ToList();

           // var a1 = ProdRecipeList.Where(x => x.ProductCode == "1021001476").ToList();
           // var a2 = ProdRecipeList.Where(x => x.ProductCode == "4500000034").ToList();
           // var a3 = ProdRecipeList.Where(x => x.ProductCode == "4500000033").ToList();


            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            Cursor = Cursors.Default;
        }


        public DateTime ConvertDataToDate(string s)
        {
            DateTime Res = DateTime.Today.Date;
            string dd = "";
            string mm = "";
            string yy = "";

            /*try
            {
                Res = Convert.ToDateTime(s);
            }
            catch (Exception xx)
            {
                dd = s.Substring(0, 2);

                if (dd.Length == 1)
                {
                    dd = "0" + dd;
                }

                mm = s.Substring(3, 2);
                if (mm.Length == 1)
                {
                    mm = "0" + mm;
                }
                yy = s.Substring(6, 4);
                yy = yy + "-" + mm + "-" + dd;
                Res = Convert.ToDateTime(yy);
            }*/
            try
            {
                dd = s.Substring(0, 2);

                if (dd.Length == 1)
                {
                    dd = "0" + dd;
                }

                mm = s.Substring(3, 2);
                if (mm.Length == 1)
                {
                    mm = "0" + mm;
                }
                yy = s.Substring(6, 4);
                yy = yy + "-" + mm + "-" + dd;
                Res = Convert.ToDateTime(yy);
            }
            catch (Exception xx)
            {
                MessageBox.Show("Некорректная дата в файле данных" + System.Environment.NewLine + s, "Сообщение системы");
                this.Close();
            }

            return Res;
        }

        public double ConvertStringToDouble(string VV)
        {
            string s = VV;
            double res = 0;
            s = s.Replace(".", ",");
            try
            {
                res = Convert.ToDouble(s);
            }
            catch (Exception xx)
            {
                s = s.Replace(",", ".");
                try
                {
                    res = Convert.ToDouble(s);
                }
                catch (Exception xxx)
                {
                    res = 0;
                }
            }
            return res;
        }

        private void ExcelDataLoader()
        {
            LoadFlag = true;
            System.Threading.Thread threadWW = new System.Threading.Thread(WaitingShowFormWW);
            threadWW.Start();

            ErrorStr = "";
            ErrorMainStr = "";

            MainEGrid = new List<ErrorGrid>();
            CommonEGrid = new List<ErrorGrid>();

            Invoke(new Action(() =>
            {
                Cursor = Cursors.WaitCursor;

                progressBar1.Visible = true;
                progressBar1.Value = 1;

                label1.Visible = true;
                label1.Text = "Открытие Xls-файла ";
            }));

            Excel1.Application xlApp = new Excel1.Application();
            xlApp.Visible = false;
            Excel1.Workbook wBook = xlApp.Workbooks.Open(FileName);
            xlApp.DisplayAlerts = false;
            bool AllRight;
            string Err = "";
            Excel1.Worksheet wSheet = new Excel1.Worksheet();

            try
            {
                wSheet = (Excel1.Worksheet)wBook.Sheets["Лист1"];
                AllRight = true;
            }
            catch (Exception xx)
            {
                AllRight = false;
                Err = xx.Message;
            }

            Excel1.Range range1;

            if (AllRight == true)
            {
                Invoke(new Action(() =>
                {
                    label1.Text = "Открытие Xls-файла... ";
                    progressBar1.Value = 10;
                }));

                Int32 CurRec = 2;
                bool fl = false;
                string LineName = "";
                DateTime DateWork = DateTime.Today.Date;
                string ProductCode = "";
                string ProdName = "";
                double Quantity = 0;

                while (fl == false)
                {
                    range1 = (Excel1.Range)wSheet.Cells[CurRec, 1];
                    try
                    {
                        Invoke(new Action(() =>
                        {
                            LineName = range1.Value.ToString().Trim();
                            LineName = ReplaceLineName(LineName);
                        }));

                        range1 = (Excel1.Range)wSheet.Cells[CurRec, 2];
                        try
                        {
                            Invoke(new Action(() =>
                            {
                                DateWork = ConvertDataToDate(range1.Value.ToString().Trim());
                            }));
                            range1 = (Excel1.Range)wSheet.Cells[CurRec, 3];
                            try
                            {
                                Invoke(new Action(() =>
                                {
                                    if (!String.IsNullOrEmpty(range1.Value.ToString()))
                                    {
                                        ProductCode = range1.Value.ToString().Trim();
                                    }
                                    else
                                    {
                                        ProductCode = "0";
                                    }
                                }));

                                range1 = (Excel1.Range)wSheet.Cells[CurRec, 4];
                                try
                                {
                                    Invoke(new Action(() =>
                                    {
                                        if (!String.IsNullOrEmpty(range1.Value.ToString()))
                                        {
                                            ProdName = range1.Value.ToString().Trim();
                                        }
                                        else
                                        {
                                            ProdName = "";
                                        }
                                        ProdName = ProdName.Replace(ProductCode, "");
                                    }));

                                    range1 = (Excel1.Range)wSheet.Cells[CurRec, 5];

                                    try
                                    {
                                        Invoke(new Action(() =>
                                        {
                                            try
                                            {
                                                Quantity = ConvertStringToDouble(range1.Value.ToString().Trim());
                                            }
                                            catch (Exception xxx4)
                                            {
                                                Quantity = 0;
                                            }
                                        }));

                                        if ((ProductCode != "0000_00") && (ProductCode != "0") && (Quantity != 0))
                                        {
                                            lData ld = new lData
                                            {
                                                DateWork = DateWork,
                                                WorkDay = DateWork.Day,
                                                WorkMonth = DateWork.Month,
                                                WorkYear = DateWork.Year,
                                                ProdCode = ProductCode.Replace("_П/УП", ""),
                                                ProdCodeStr = ProductCode,
                                                Quantity = Quantity,
                                                ProdName = ProdName,
                                                AlterProdName = ProdName,
                                                RePack = ProductCode.Contains("_П/УП") ? true : false,
                                                Line = LineName,
                                                Spices = LineName.Contains("специй") ? true : false
                                            };

                                            Invoke(new Action(() =>
                                            {
                                                ExcelDataList.Add(ld);

                                                DateWork = new DateTime(DateWork.Year, DateWork.Month, 1);
                                                ExcelDatesList.Add(DateWork);
                                            }));
                                        }
                                        CurRec = CurRec + 1;

                                        Invoke(new Action(() =>
                                        {
                                            if (CurRec % 100 == 0)
                                            {
                                                label1.Text = "Загружено " + CurRec.ToString() + " записей";
                                                progressBar1.Value = progressBar1.Value + 1;
                                                if (progressBar1.Value >= 60)
                                                {
                                                    progressBar1.Value = 10;
                                                }
                                            }
                                        }));
                                    }
                                    catch (Exception xx3)
                                    {
                                        fl = true;
                                    }

                                }
                                catch (Exception xxx2)
                                {
                                    fl = true;
                                }
                            }
                            catch (Exception xx2)
                            {
                                fl = true;
                            }
                        }
                        catch (Exception xx1)
                        {
                            CurRec = CurRec + 1;
                            //fl = true;
                        }
                    }
                    catch (Exception xx)
                    {
                        fl = true;
                    }
                }
                //MOnthly graphic
                Invoke(new Action(() =>
                {
                    label1.Text = "Формирование списка дат";
                    progressBar1.Value = 65;

                    CorrectDataList();
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Коррекция списка дат";
                    progressBar1.Value = 70;

                    ExcelDatesList = ExcelDatesList.Distinct().ToList();
                    ExcelDataList = ExcelDataList.OrderBy(x => x.DateWork).ThenBy(x => x.Line).ToList();

                    for (int i = 0; i < ExcelDataList.Count; i++)
                    {
                        ExcelDataList[i].RowIndex = i + 1;
                    }
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Проверка рецептуры ";
                    progressBar1.Value = 75;
                }));

                Invoke(new Action(() =>
                {
                    ErrorStr = GetPecipesNotes();
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Наложение фильтров ";
                    progressBar1.Value = 80;
                }));

                Invoke(new Action(() =>
                {
                    var tdata = tdb.tComponentException.Where(x => x.IsActive == true).ToList();
                    List<string> ExcepNames = tdata.Select(x => x.ProductCode).ToList();
                    ExcepNames = ExcepNames.Distinct().ToList();

                    foreach (string en in ExcepNames)
                    {
                        var eee = ExcelDataList.Where(x => x.ProdCode.Contains(en)).ToList();
                        foreach (var e in eee)
                        {
                            e.Spices = true;
                        }
                    }
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Отрисовка данных. ждите";
                    progressBar1.Value = 85;
                }));

                Invoke(new Action(() =>
                {
                    GreateOperList();
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Коррекция наименований продуктов";
                    progressBar1.Value = 90;
                }));

                Invoke(new Action(() =>
                {
                    AlterProductNames();
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Запись в базу данных";
                    progressBar1.Value = 95;
                }));

                Invoke(new Action(() =>
                {
                //  SaveRecordToDB();
            }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Закрытие xls-файла";
                    progressBar1.Value = 97;
                }));

                Invoke(new Action(() =>
                 {
                     xlApp.Quit();
                     xlApp = null;
                     wBook = null;
                     wSheet = null;
                     range1 = null;
                 }));

                try
                {
                    LoadFlag = false;
                    System.Threading.Thread.Sleep(250);
                    threadWW.Abort();
                }
                catch (Exception xx)
                { }

                Invoke(new Action(() =>
                {
                    label1.Text = "Проверка на наличие ошибок";
                    progressBar1.Value = 98;
                    Cursor = Cursors.Default;

                    if ((ErrorMainStr.Length > 0) || (ErrorStr.Length > 0))
                    {
                    //  ErrorMainStr = ErrorMainStr + " " + RowSplitter + " " + RowSplitter + System.Environment.NewLine + " " + RowSplitter + " " + RowSplitter + System.Environment.NewLine + ErrorStr;
                    if (Application.OpenForms["Warning1"] == null)
                        {
                            Cursor = Cursors.WaitCursor;
                            w1 = new Warning1(this);
                            w1.Visible = true;
                        //   fs.MdiParent = this;
                        w1.WindowState = FormWindowState.Normal;
                            Cursor = Cursors.Default;
                        }
                        else
                        {
                            w1.BringToFront();
                            w1.Visible = true;
                        //   cf.UpdateData();
                    }
                        this.Enabled = false;
                    }
                    else
                    {

                        MessageBox.Show("Данные успешно приняты!", "Сообщение системы");
                        this.BringToFront();
                    }
                }));

                Invoke(new Action(() =>
                {
                    label1.Text = "Все почти готово";
                    progressBar1.Value = 99;

                    GC.Collect();
                    GC.WaitForPendingFinalizers();
                    GC.Collect();
                    Cursor = Cursors.Default;
                }));
            } //allRight
            else
            {
                Invoke(new Action(() =>
                {
                    xlApp.Quit();
                    xlApp = null;
                    wBook = null;
                    wSheet = null;
                    range1 = null;
                }));

                MessageBox.Show("Ошибка отерытия файла!" + System.Environment.NewLine + "Убедитесь, что вы открываете нужный файл.", "Сообщение системы");
            }

            Invoke(new Action(() =>
            {
                for (int i = 0; i<dataGridView1.ColumnCount; i++)
                {
                    dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
                }
            }));

            Invoke(new Action(() =>
            {
                label1.Text = "";
                label1.Visible = false;
                progressBar1.Visible = false;
                progressBar1.Value = 1;

                panel1.Enabled = true;
                panel3.Enabled = true;

            }));
        }

        private void CorrectDataList()
        {
            int CurrMonth = DateTime.Today.Month;
            int CurrYear = DateTime.Today.Year;
            int DayCount = System.DateTime.DaysInMonth(CurrYear, CurrMonth);
            DateTime DDT;
            DateTime DStart;
            DateTime DEnd;

            var CurrDataList = ExcelDataList.Where(x => x.WorkMonth == CurrMonth && x.WorkYear == CurrYear).ToList();
         //   CurrDataList = CurrDataList.OrderBy(x => x.DateWork).ToList();
            List<DateTime> CurrDList = CurrDataList.Select(x => x.DateWork).Distinct().ToList();
            CurrDList.Sort();

            if (CurrDList.Count > 0)
            {
                if (CurrDList.Count < DayCount)
                {
                    if (CurrDList[0].Day != 1)
                    {
                        DDT = CurrDList[0];

                        CurrDataList = ExcelDataList.Where(x => x.DateWork < DDT).ToList();
                        
                        if (CurrDataList.Count>0)
                        {
                            DDT = CurrDataList.Max(x => x.DateWork);
                            CurrDList.Add(DDT);
                            CurrDList.Sort();

                            DStart = CurrDList[0];
                            DEnd = CurrDList.Last();

                            CurrDataList = ExcelDataList.Where(x => x.DateWork >= DStart && x.DateWork <= DEnd).ToList();
                        }
                    }

                    if (CurrDList.Last().Day < DayCount)
                    {
                        DDT = new DateTime(CurrYear, CurrMonth, DayCount);
                        CurrDataList = ExcelDataList.Where(x => x.DateWork > DDT).ToList();

                        if (CurrDataList.Count > 0)
                        {
                            DDT = CurrDataList.Min(x => x.DateWork);

                            CurrDList.Add(DDT);

                            DStart = CurrDList[0];
                            DEnd = CurrDList.Last();

                            CurrDataList = ExcelDataList.Where(x => x.DateWork >= DStart && x.DateWork < DEnd).ToList();
                        }
                    }

                    for (int i = 0; i < CurrDList.Count - 1; i++)
                    {
                         DStart = CurrDList[i];
                         DEnd = CurrDList[i + 1];

                        var X = (DEnd - DStart).TotalDays;

                        if (X > 1)
                        {
                            while (DStart < DEnd)
                            {
                                if (DStart == CurrDList[i])
                                {
                                    var CDL = CurrDataList.Where(x => x.DateWork == DStart).ToList();

                                    for (int j = 0; j<CDL.Count; j++)
                                    {
                                        CDL[j].Quantity = Math.Round(CDL[j].Quantity / X, 3);
                                    }
                                }
                                else
                                {
                                    var CDL = CurrDataList.Where(x => x.DateWork == DStart).ToList();

                                    for (int j = 0; j < CDL.Count; j++)
                                    {
                                        lData LD = new lData();

                                        LD.AlterProdName = CDL[j].AlterProdName;
                                        LD.DateWork = DStart;
                                        LD.IPG = CDL[j].IPG;
                                        LD.BOM = CDL[j].BOM;
                                      
                                        LD.Line = CDL[j].Line;
                                        LD.ProdCode = CDL[j].ProdCode;
                                        LD.ProdCodeStr = CDL[j].ProdCodeStr;
                                        LD.ProdName = CDL[j].ProdName;
                                        LD.Quantity = CDL[j].Quantity;
                                        LD.RePack = CDL[j].RePack;
                                        LD.Spices = CDL[j].Spices;
                                        LD.WorkDay = DStart.Day;
                                        LD.WorkMonth = DStart.Month;
                                        LD.WorkYear = DStart.Year;

                                        ExcelDataList.Add(LD);
                                    }
                                }

                                DStart = DStart.AddDays(1);
                            }
                        }
                    }
                }
            }


            /*
            if (CurrDList.Count == 1)
            {
                if (CurrDList[0].Day == 1)
                {
                    DateTime DateStart = new DateTime(CurrYear, CurrMonth, 1, 0, 0, 0);
                    DateTime DateEnd = DateStart.Date.AddMonths(1).AddDays(-1);

                    double DayCount = (DateEnd - DateStart).TotalDays;
                    if (DayCount > 0)
                    {
                      //  DayCount = 1 / DayCount;

                        while (DateStart <= DateEnd)
                        {
                            int CurrDay = DateStart.Day;
                            
                            if (CurrDay == 1)
                            {
                                for (int i = 0; i<CurrDataList.Count; i++)
                                {
                                    CurrDataList[i].Quantity = Math.Round(CurrDataList[i].Quantity / DayCount, 3);
                                }
                            }
                            else
                            {
                                for (int i = 0; i<CurrDataList.Count; i++)
                                {
                                    lData LD = new lData();
                                                                        
                                    LD.AlterProdName = CurrDataList[i].AlterProdName;
                                    LD.DateWork = DateStart;
                                    LD.IPG = CurrDataList[i].IPG;
                                    LD.Line = CurrDataList[i].Line;
                                    LD.ProdCode = CurrDataList[i].ProdCode;
                                    LD.ProdCodeStr = CurrDataList[i].ProdCodeStr;
                                    LD.ProdName = CurrDataList[i].ProdName;
                                    LD.Quantity = CurrDataList[i].Quantity;
                                    LD.RePack = CurrDataList[i].RePack;
                                    LD.Spices = CurrDataList[i].Spices;
                                    LD.WorkDay = DateStart.Day;
                                    LD.WorkMonth = DateStart.Month;
                                    LD.WorkYear = DateStart.Year;

                                    ExcelDataList.Add(LD);
                                }                                
                            }

                            ExcelDatesList.Add(DateStart);
                            DateStart = DateStart.AddDays(1);
                        }
                    }
                }
            }
            */

            GC.Collect();
        }

        private void SaveRecordToDB()
        {
            tdb.Update_WorkPlanTable();

            for (int i=0; i<ExcelDataList.Count; i++)
            {
                if (ExcelDataList[i].Quantity>0)
                {
                    if (ExcelDataList[i].RePack==false)
                    {
                        string code = ExcelDataList[i].ProdCode;
                        string Group = "";
                        try
                        {
                            Group = tdb.RPSettings.Where(x => x.ProductCode == code).Select(x => x.ProductGroup).First();
                        }
                        catch (Exception xx)
                        {
                            Group = "NoData";
                        }

                        if ((ExcelDataList[i].DateWork > DateTime.Today.Date.AddDays(-20))&&(ExcelDataList[i].DateWork < DateTime.Today.Date.AddDays(20)))
                        {
                            tdb.Insert_WorkPlanTableRecord1(ExcelDataList[i].DateWork, ExcelDataList[i].Line, ExcelDataList[i].ProdCode, Group, ExcelDataList[i].Quantity);
                        }
                    }
                }
            }

          //  tdb.Modify_WorkPlanTable();
            tdb.SaveChanges();

            GC.Collect();
        }

        private void InsertFromExcel(bool flag)
        {
            DialogResult dr = MessageBox.Show("Подтверждаете загрузку данных из файла?", "Сообщение системы", MessageBoxButtons.OKCancel);

            /*bool Res = false;
            string ErrorStr = "";*/
    
            if (dr == DialogResult.OK)
            {
                if (openFileDialog1.ShowDialog() == DialogResult.OK)
                {
                    FileName = openFileDialog1.FileName;

                    if (flag == true)
                    {
                        ExcelDataList = new List<lData>();
                        ExcelDatesList = new List<DateTime>();
                    }

                    panel1.Enabled = false;
                    panel3.Enabled = false;
                    Thread thread = new Thread(ExcelDataLoader);
                    thread.Start();
                }
            }

            this.Text = "Планирование производства фабрика-кухня; Верс.1.0.2.15; Загружено " + DateTime.Now.ToString("dd.MM.yyyy HH:mm:ss") + " " + FileName;
        }

        public string GetPecipesNotes()
        {
            ProductRecipes pr;
            ProductLinesRecipes plr;
            Recipe rr;
            bool Res;
            DateTime ddt = DateTime.Today.Date;

            int alfa;
            int bravo;
            int charlie;

           ErrorStr = "";

            var RData = tdb.fn_select_RecipesBisView1().ToList();  //рецепты без линий!

            foreach (var xl in ExcelDataList) //.Where(x => (!x.ProdName.ToLower().Contains("тест")) && (x.DateWork >= DateTime.Today.Date)).ToList())
            {
                /*if ((xl.ProdCode == "1022001105") || (xl.ProdCode == "1022001106"))
                { }*/
                var PRList = ProdRecipeList.Where(x => ( x.ProductCode == xl.ProdCode)).ToList();

                if (PRList.Count == 0)
                {
                    Res = false;
                    var data = RData.Where(x => x.ProductCode == xl.ProdCode).ToList();     // tdb.fn_select_BasicRecipes(xl.ProdCode).ToList();

                    if (data.Count > 0)
                    {
                        ddt = (DateTime)data[0].WorkDate;     //StartingDate;

                        pr = new ProductRecipes();

                        pr.RecLineList = new List<ProductLinesRecipes>();
                        pr.ProductCode = data[0].ProductCode;
                        pr.ProductCodeString = pr.ProductCode;
                        pr.ProductName = data[0].ProductName;
                        pr.AlterProductCode= data[0].ProductCode;

                        ProdRecipeList.Add(pr);

                        plr = new ProductLinesRecipes();
                        plr.RecList = new List<Recipe>();
                        plr.LineNumber = "Линия 1";
                        alfa = ProdRecipeList.Count - 1;
                        ProdRecipeList[alfa].RecLineList.Add(plr);

                        rr = new Recipe();
                        rr.rList = new List<RecipeList>();
                        rr.WorkDateStart = ddt;
                        bravo = 0;
                        ProdRecipeList[alfa].RecLineList[bravo].RecList.Add(rr);

                        foreach (var dd in data.Where(x => x.WorkDate==ddt).ToList())   //(x.LineName.Trim() == "") && (x.StartingDate == ddt)
                        {
                            RecipeList rl = new RecipeList();
                            rl.IPG = dd.IPG;
                            rl.MaterialCode = dd.MaterialCode;
                            rl.MaterialName = dd.MaterialName;
                            rl.MaterialQuantity = (double)dd.Quantity;
                            rl.UMC = dd.UnitMeasure;
                            rl.BOM = dd.BOM;
                            charlie = ProdRecipeList[alfa].RecLineList[bravo].RecList.Count - 1;
                            ProdRecipeList[alfa].RecLineList[bravo].RecList[charlie].rList.Add(rl);

                            PR_GetRecipesDataMainProductBis_Result RVR = new PR_GetRecipesDataMainProductBis_Result
                            {
                                ProductCode= data[0].ProductCode,
                                AlterProductCode=data[0].ProductCode,
                                LineNumber= "Линия 1",
                                WorkDate=ddt,
                                MaterialCode= dd.MaterialCode,
                                MaterialName=dd.MaterialName,
                                UnitMeasure= dd.UnitMeasure,
                                Quantity= (decimal)dd.Quantity,
                                IPG=dd.IPG,
                                Bom = dd.BOM
                            };

                            lRec.Add(RVR);
                        }

                        if (ErrorStr.Length == 0)
                        {
                            ErrorStr = ErrorStr + " " + RowSplitter + "" + RowSplitter + "Отсутствует рецепт (версия рецепта) для линий " + RowSplitter + System.Environment.NewLine;
                        }
                        else
                        { }

                        // ErrorStr = ErrorStr + " " + xl.DateWork.ToString("dd.MM.yyyy") + " " + RowSplitter + xl.ProdCode + " " + RowSplitter + xl.ProdName + " " + RowSplitter + " использован базовый рецепт без номера версии и линии. Укажите версию спецификации в НАВ!" + RowSplitter + System.Environment.NewLine;

                        CommonEGrid.Add(new ErrorGrid
                        {
                            DateWork = xl.DateWork,
                            ProdCode = xl.ProdCode,
                            ProdName = xl.ProdName,
                            LineName = xl.Line,
                            Comment = "Использован базовый рецепт без номера версии и линии. Укажите версию спецификации в НАВ!"
                        });

                    }
                    else
                    {
                        Res = true;

                        if (ErrorMainStr.Length==0)
                        {
                            ErrorMainStr = ErrorMainStr + " " + RowSplitter + "" + RowSplitter + "Не указан КодВерсии рецепта либо отсутствует версия спецификации рецепта: " + RowSplitter + System.Environment.NewLine;
                        }
                        else
                        { }

                        //   ErrorMainStr = ErrorMainStr + " " + xl.DateWork.ToString("dd.MM.yyyy") + " " + RowSplitter + xl.ProdCode + " " + RowSplitter + xl.ProdName + " " + RowSplitter + xl.Line + RowSplitter + " В количестве: " + xl.Quantity.ToString("N1") + RowSplitter + System.Environment.NewLine;
                        MainEGrid.Add(new ErrorGrid
                        {
                            DateWork = xl.DateWork,
                            ProdCode = xl.ProdCode,
                            ProdName = xl.ProdName,
                            LineName = xl.Line,
                            Comment = " В количестве: " + xl.Quantity.ToString("N1")
                        });
                    }
                }
            }

            // if (Res==false)
            {
                foreach (var xl in ExcelDataList) //.Where(x => /*(!x.ProdName.ToLower().Contains("тест")) && */(x.DateWork >= DateTime.Today.Date)).ToList())
                {
                    var PRList = ProdRecipeList.Where(x => x.ProductCode == xl.ProdCode).ToList();
                    try
                    {
                        var LList = PRList[0].RecLineList.Where(x => x.LineNumber == xl.Line).ToList();
                        if (LList.Count == 0)
                        {
                            plr = new ProductLinesRecipes
                            {
                                LineNumber = xl.Line,
                                RecList = new List<Recipe>()
                            };

                            foreach (var RLL in PRList[0].RecLineList[0].RecList)
                            {
                                rr = RLL;
                                plr.RecList.Add(rr);


                                foreach (var RL in RLL.rList)
                                {

                                    PR_GetRecipesDataMainProductBis_Result RVR = new PR_GetRecipesDataMainProductBis_Result
                                    {
                                        ProductCode = xl.ProdCode,
                                        AlterProductCode = xl.ProdCode,
                                        LineNumber = xl.Line,
                                        WorkDate = RLL.WorkDateStart,
                                        MaterialCode = RL.MaterialCode,
                                        MaterialName = RL.MaterialName,
                                        UnitMeasure = RL.UMC,
                                        Quantity =(decimal) RL.MaterialQuantity,
                                        IPG = RL.IPG,
                                        Bom = RL.BOM
                                    };

                                    lRec.Add(RVR);
                                }
                            }

                            PRList[0].RecLineList.Add(plr);

                            if (ErrorStr.Length == 0)
                            {
                                ErrorStr = ErrorStr + " " + RowSplitter + "" + RowSplitter + "Отсутствует рецепт (версия рецепта) для линий " + RowSplitter + System.Environment.NewLine;
                            }
                            else
                            { }

                            //   ErrorStr = ErrorStr + " " + xl.DateWork.ToString("dd.MM.yyyy") + " " + RowSplitter + xl.ProdCode + " " + RowSplitter + xl.ProdName + " " + RowSplitter + xl.Line + RowSplitter + ". Использован  рецепт для другой линии. Добавьте рецепт для " + xl.Line + RowSplitter + System.Environment.NewLine;

                            CommonEGrid.Add(new ErrorGrid
                            {
                                DateWork = xl.DateWork,
                                ProdCode = xl.ProdCode,
                                ProdName = xl.ProdName,
                                LineName = xl.Line,
                                Comment = "Использован  рецепт для другой линии. Добавьте рецепт для " + xl.Line
                            });
                        }
                        else
                        {
                            // 
                        }
                    }
                    catch (Exception xx)
                    {
                        if (ErrorStr.Length == 0)
                        {
                            ErrorStr = ErrorStr + " " + RowSplitter + "" + RowSplitter + "Отсутствует рецепт (версия рецепта) для линий " + System.Environment.NewLine;
                        }
                        else
                        { }

                        // ErrorStr = ErrorStr + " " + xl.DateWork.ToString("dd.MM.yyyy") + " " + RowSplitter + xl.ProdCode + " " + RowSplitter + xl.ProdName + "   " + RowSplitter + xl.Line + ". Добавьте рецепт для " + xl.Line + RowSplitter + System.Environment.NewLine;
                        CommonEGrid.Add(new ErrorGrid
                        {
                            DateWork = xl.DateWork,
                            ProdCode = xl.ProdCode,
                            ProdName = xl.ProdName,
                            LineName = xl.Line,
                            Comment = "Добавьте рецепт для " + xl.Line
                        });
                    }
                }
            }

            //   if (Res==false)
            {
                foreach (var xl in ExcelDataList) //.Where(x => /*(!x.ProdName.ToLower().Contains("тест")) &&*/ (x.DateWork >= DateTime.Today.Date)).ToList())
                {
                    var PRList = ProdRecipeList.Where(x => x.ProductCode == xl.ProdCode).ToList();
                    try
                    {
                        var LList = PRList[0].RecLineList.Where(x => x.LineNumber == xl.Line).ToList();
                        var DDList = LList[0].RecList.Where(x => (x.WorkDateStart <= xl.DateWork) && (x.WorkDateEnd >= xl.DateWork)).ToList();

                        if (DDList.Count == 0)
                        {
                            LList[0].RecList[LList[0].RecList.Count - 1].WorkDateEnd = DateTime.Today.AddYears(25);

                            if (LList[0].RecList.Count == 1)
                            {
                                LList[0].RecList[0].WorkDateStart = xl.DateWork.AddDays(-2);
                            }
                            else if (LList[0].RecList[0].WorkDateStart > xl.DateWork)
                            {
                                LList[0].RecList[0].WorkDateStart = xl.DateWork.AddDays(-2);
                            }

                            if (ErrorStr.Length == 0)
                            {
                                ErrorStr = ErrorStr + " " + RowSplitter + " " + RowSplitter + "Отсутствует рецепт (версия рецепта) для линий " + System.Environment.NewLine;
                            }
                            else
                            { }

                            //   ErrorStr = ErrorStr + " " + xl.DateWork.ToString("dd.MM.yyyy") + " " + RowSplitter + xl.ProdCode + " " + RowSplitter + xl.ProdName + " " + RowSplitter + xl.Line + RowSplitter + " Дата начала/окончания увеличена принудительно. Укажите действующий рецепт на дату в НАВ" + RowSplitter + System.Environment.NewLine;
                            CommonEGrid.Add(new ErrorGrid
                            {
                                DateWork = xl.DateWork,
                                ProdCode = xl.ProdCode,
                                ProdName = xl.ProdName,
                                LineName = xl.Line,
                                Comment = " Дата начала/окончания увеличена принудительно. Укажите действующий рецепт на дату в НАВ"
                            });
                        }
                    }
                    catch (Exception xx)
                    {
                        if (ErrorStr.Length == 0)
                        {
                            ErrorStr = ErrorStr + " " + RowSplitter + " " + RowSplitter + "Отсутствует рецепт (версия рецепта) для линий " + System.Environment.NewLine;
                        }
                        else
                        { }

                        //   ErrorStr = ErrorStr + " " + xl.DateWork.ToString("dd.MM.yyyy") + " " + RowSplitter + xl.ProdCode + " " + RowSplitter + xl.ProdName + " " + RowSplitter + xl.Line + RowSplitter + " Отсутствует рецепт для линии! " + RowSplitter + System.Environment.NewLine;
                        CommonEGrid.Add(new ErrorGrid
                        {
                            DateWork = xl.DateWork,
                            ProdCode = xl.ProdCode,
                            ProdName = xl.ProdName,
                            LineName = xl.Line,
                            Comment =  " Отсутствует рецепт для линии!" 
                        });
                    }
                }
            }

            return ErrorStr;
        }

        public void AlterProductNames()
        {
            var RPS = tdb.tMaterial.ToList();   //eqp.RPSettings.ToList();

            foreach (var XL in ExcelDataList)
            {
                var r = RPS.Where(x => x.MaterialCode == XL.ProdCode).ToList();
                if (r.Count>0)
                {
                    if (r[0].Name.Length > 70)
                    {
                        XL.AlterProdName = r[0].Name.Substring(0, 70);
                    }
                    else
                    {
                        XL.AlterProdName = r[0].Name;
                    }
                }
            }

            GC.Collect();
        }

        private void принятьДанныеИзExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertFromExcel(true);
        }

        public void GreateOperList()
        {
            POList = new List<PlantOperList>();
            List<DateTime> DTList = ExcelDataList.Select(x=>x.DateWork).Distinct().ToList();

            foreach (DateTime dt in DTList)
            {
                var Edl = ExcelDataList.Where(x => x.DateWork == dt).ToList();
                PlantOperList pol = new PlantOperList
                {
                    DateWork = dt,
                    DL = new List<DataList>()
                };

                foreach (var e in Edl)
                {
                    DataList dl = new DataList
                    {
                        Line = e.Line,
                        ProdCode = e.ProdCode,
                        ProdCodeStr = e.ProdCodeStr,
                        ProdName = e.ProdName,
                        Quantity = e.Quantity,
                        RePack = e.RePack
                    };
                    pol.DL.Add(dl);
                }

                POList.Add(pol);
            }
            POList = POList.OrderBy(x => x.DateWork).ToList();

            GC.Collect();

            DrawDGV1();
        }

        public void DrawDGV1()
        {
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();
            dataGridView1.Columns.Clear();
            dataGridView1.DoubleBuffered = false;

            Int32 RowCount = 0;
            Int32 ColCount = POList.Count;
            dt1 = new DataTable();
            string[] NewString;
            string sVal;

            double TotQuant = 0;
            double RePackQuant = 0;
            double SpiceQuant = 0;

            if (ColCount > 0)
            {
                foreach (PlantOperList pol in POList)
                {
                    if (pol.DL.Count > RowCount)
                    {
                        RowCount = pol.DL.Count;
                      //  MessageBox.Show(RowCount.ToString());
                    }

                    TotQuant = 0;
                    RePackQuant = 0;
                    SpiceQuant = 0;

                    foreach (var dll in pol.DL)
                    {
                        if (dll.Line.Contains("специй"))
                        {
                            SpiceQuant = SpiceQuant + dll.Quantity;
                        }
                        else
                        {
                            if (dll.RePack == true)
                            {
                                RePackQuant = RePackQuant + dll.Quantity;
                            }
                            else
                            {
                                TotQuant = TotQuant + dll.Quantity;
                            }
                        }
                    }

                    DataList ddd = new DataList
                    {
                        ProdCodeStr = "Производство " + TotQuant.ToString("N0") + " кг, Переупаковка " + RePackQuant.ToString("N0") + " кг, Специи " + SpiceQuant.ToString("N0") + " кг",
                        Line = "",
                        ProdCode = "",
                        ProdName = "",
                        RePack = false,
                        Quantity = 0,
                        LastRec = true
                    };

                    pol.DL.Add(ddd);
                    dt1.Columns.Add(pol.DateWork.ToString("dd.MM.yyyy"));
                }
                dt1.Columns.Add("Zero");               

                RowCount = RowCount + 1;

                for (int i=0; i<RowCount; i++)
                {
                    sVal = "";
                    for (int j=0; j<ColCount; j++)
                    {
                        if (i < POList[j].DL.Count)
                        {
                            try
                            {
                                if (POList[j].DL[i].LastRec == false)
                                {
                                    sVal = sVal + POList[j].DL[i].Line + " (" + POList[j].DL[i].ProdCodeStr + ") " + POList[j].DL[i].ProdName;
                                }
                                else
                                {
                                    sVal = sVal + POList[j].DL[i].ProdCodeStr;
                                }
                                if (POList[j].DL[i].RePack == true)
                                {
                                    sVal = sVal + " (Переупаковка) [" + POList[j].DL[i].Quantity.ToString("N0") + " кг]";
                                }
                                else
                                {
                                    if (POList[j].DL[i].Line.Contains("специй"))
                                    {
                                        sVal = sVal + " (Специи) [" + POList[j].DL[i].Quantity.ToString("N0") + " кг]";
                                    }
                                    else
                                    {
                                        if (POList[j].DL[i].LastRec == false)
                                        {
                                            sVal = sVal + " [" + POList[j].DL[i].Quantity.ToString("N0") + " кг]";
                                        }
                                    }
                                }

                                sVal = sVal + RowSplitter;
                            }
                            catch (Exception xx)
                            {
                                sVal = sVal + " " + RowSplitter;
                            }
                        }
                        else
                        {
                            sVal = sVal + " " + RowSplitter;
                        }
                    }

                    NewString = sVal.Split(RowSplitter);
                    dt1.Rows.Add(NewString);
                }

                bs = new BindingSource();
                bs.DataSource = dt1;
                dataGridView1.DataSource = bs;

                Font f1 = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Bold);
            //    Font f2 = new Font(dataGridView1.DefaultCellStyle.Font, Color.LightSeaGreen);
                for (int i=0; i<POList.Count; i++)
                {
                    for (int j=0; j< POList[i].DL.Count-1; j++)
                    {
                        if (POList[i].DL[j].RePack == true)
                        {
                            dataGridView1.Rows[j].Cells[i].Style.ForeColor = Color.MediumSeaGreen;
                          //  dataGridView1.Rows[j].Cells[i].Style.BackColor = Color.DarkRed;
                        }
                        else if (POList[i].DL[j].Line.Contains("специй"))
                        {
                            dataGridView1.Rows[j].Cells[i].Style.ForeColor = Color.DodgerBlue;
                           // dataGridView1.Rows[j].Cells[i].Style.BackColor = Color.Gold;

                        }
                    }
                  
                    if (POList[i].DL[POList[i].DL.Count-1].LastRec==true)
                    {
                        dataGridView1.Rows[POList[i].DL.Count - 1].Cells[i].Style.BackColor = Color.LightYellow;
                        //  dataGridView1.Rows[POList[i].DL.Count - 1].Cells[i].Style.Font = new Font(dataGridView1.DefaultCellStyle.Font, FontStyle.Bold);    //f1;
                    } 
                }

                dataGridView1.Columns[dataGridView1.Columns.Count - 1].Visible = false;
                dataGridView1.ReadOnly = true;
                dataGridView1.AllowUserToAddRows = false;
                dataGridView1.AutoSizeColumnsMode = DataGridViewAutoSizeColumnsMode.AllCells;

                dataGridView1.EnableHeadersVisualStyles = false;
                dataGridView1.ColumnHeadersDefaultCellStyle.Font = new Font(dataGridView1.ColumnHeadersDefaultCellStyle.Font.FontFamily, 16f, FontStyle.Bold | FontStyle.Italic);
                dataGridView1.ColumnHeadersDefaultCellStyle.Alignment = DataGridViewContentAlignment.MiddleCenter;
                dataGridView1.DoubleBuffered = true;
            }

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void настройкиЗаявокToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormSettings"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fs = new FormSettings(this);
                fs.Visible = true;
             //   fs.MdiParent = this;
                fs.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fs.BringToFront();
                //   cf.UpdateData();
            }
        }

        private void потреблениеНаПроизводствеToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void показатьТекущиеОстаткиToolStripMenuItem_Click(object sender, EventArgs e)
        {

        }

        private void выходToolStripMenuItem_Click(object sender, EventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            this.Close();
        }

        private void выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (POList.Count > 0)
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
                wSheet.Cells[1, 1] = "Линия";
                wSheet.Cells[1, 2] = "Дата";
                wSheet.Cells[1, 3] = "Код ГП";
                wSheet.Cells[1, 4] = "Название ГП";
                wSheet.Cells[1, 5] = "Объем";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 5]];
                range1.Font.Bold = true;
                range1.HorizontalAlignment = HorizontalAlignment.Center;           

                foreach (var pol in POList.OrderBy(x=>x.DateWork).ToList())
                {
                    foreach (var ddl in pol.DL)
                    {
                        if (ddl.LastRec==false)
                        {
                            wSheet.Cells[CurRow, 1] = ddl.Line;
                            wSheet.Cells[CurRow, 2] = pol.DateWork.Date;
                            ((Excel1.Range)wSheet.Cells[CurRow, 3]).NumberFormat = "@";
                            wSheet.Cells[CurRow, 3] = ddl.ProdCodeStr;
                            wSheet.Cells[CurRow, 4] = ddl.ProdName;
                            ((Excel1.Range)wSheet.Cells[CurRow, 5]).NumberFormat = "### ### ### ###";
                            wSheet.Cells[CurRow, 5] = ddl.Quantity;

                            CurRow = CurRow + 1;
                        }
                    }
                }

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[CurRow, 5]];
                range1.EntireColumn.AutoFit();

                wSheet = null;
                wBook = null;
             //   xlApp.Quit();
                xlApp = null;          
            }

            Cursor = Cursors.Default;
            GC.Collect();
            MessageBox.Show("Данные успешно выгружены");
        }

        private void сохранитьКартуВБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormSaveCard"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fsc = new FormSaveCard(this);
                fsc.Visible = true;
                //   fs.MdiParent = this;
                fsc.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fsc.BringToFront();
                //   cf.UpdateData();
                fsc.Visible = true;
            }
            this.Enabled = false;
        }

        private void загрузитьКартуИзБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormLoadCard"] == null)
            {
                Cursor = Cursors.WaitCursor;
                flc = new FormLoadCard(this);
                flc.Visible = true;
                //   fs.MdiParent = this;
                flc.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                flc.BringToFront();
                //   cf.UpdateData();
                flc.Visible = true;
            }
            this.Enabled = false;
        }

        private void добавитьКартуИзБДToolStripMenuItem_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormAddCard"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fac = new FormAddCard(this);
                fac.Visible = true;
                //   fs.MdiParent = this;
                fac.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fac.BringToFront();
                fac.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void button6_Click(object sender, EventArgs e)
        {
            принятьДанныеИзExcelToolStripMenuItem_Click(sender, e);
        }

        private void button5_Click(object sender, EventArgs e)
        {
            выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem_Click(sender, e);
        }

        private void добавитьКорректирующиеДанныеИзExcelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            InsertFromExcel(false);
        }

        private void button4_Click(object sender, EventArgs e)
        {
            добавитьКорректирующиеДанныеИзExcelToolStripMenuItem_Click(sender, e);
        }

        private void button7_Click(object sender, EventArgs e)
        {
            загрузитьКартуИзБДToolStripMenuItem_Click(sender, e);
        }

        private void button16_Click(object sender, EventArgs e)
        {
            добавитьКартуИзБДToolStripMenuItem_Click(sender, e);
        }

        private void button8_Click(object sender, EventArgs e)
        {
            сохранитьКартуВБДToolStripMenuItem_Click(sender, e);
        }

        private void FormMain_Resize(object sender, EventArgs e)
        {
            if (this.Width<1700)
            {
                this.Width = 1700;
            }
            if (this.Height<900)
            {
                this.Height = 900;
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Подтверждаете удаление рабочего дня?", "Сообщение системы", MessageBoxButtons.OKCancel);
            if (dr == DialogResult.OK)
            {
                string s = dataGridView1.Columns[dataGridView1.CurrentCell.ColumnIndex].HeaderText;
                DateTime dt = ConvertDataToDate(s);

                ExcelDataList = ExcelDataList.Where(x => x.DateWork.Date != dt.Date).ToList();
                GreateOperList();
            }
        }

        private void button3_Click(object sender, EventArgs e)
        {
            DialogResult dr = MessageBox.Show("Подтверждаете очистку карты?", "Сообщение системы", MessageBoxButtons.OKCancel);
            if (dr==DialogResult.OK)
            {
                ExcelDataList = new List<lData>();
                POList.Clear();

                GC.Collect();

                GreateOperList();
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["RecipesForm"] == null)
            {
                Cursor = Cursors.WaitCursor;
                rf = new RecipesForm(this);
                rf.Visible = true;
                //   fs.MdiParent = this;
                rf.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                rf.BringToFront();
                rf.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void FormMain_FormClosed(object sender, FormClosedEventArgs e)
        {
            tdb.Dispose( );
        //    edb.Dispose();
         //   eqp.Dispose();
           
          //  conn1.Close();
          //  conn1 = null;

            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
        }

        private void button15_Click(object sender, EventArgs e)
        {
            /* var DataList = tdb.fn_select_RecipesList().ToList();

             foreach (var d in DataList)
             {
                 d.LineNumber = ReplaceLineName(d.LineNumber);
             }

             var Result = DataList.GroupBy(x => x).Where(g => g.Count() > 1).Select(g => g.Key).ToList();
             Result = Result.Distinct().ToList();
             if (Result.Count>0)*/

            bool Res = false;
             ErrorStr ="";
            ErrorMainStr = "";
            CommonEGrid = new List<ErrorGrid>();
            MainEGrid = new List<ErrorGrid>();

            Int32 Val1;
            Int32 Val2;

            List<string> ProdCode = new List<string>();
            List<string> LineList = new List<string>();
            List<DateTime> WorkDateList = new List<DateTime>();
            List<string> MaterialCode = new List<string>();

            ProdCode = lRec.Select(x => x.ProductCode).Distinct().ToList();

            foreach (string pr in ProdCode)
            {
                var lData = lRec.Where(x => x.ProductCode == pr).ToList();

                LineList = lData.Select(x => x.LineNumber).Distinct().ToList();

                foreach (string ll in LineList)
                {
                    var lData1 = lData.Where(x => x.LineNumber == ll).ToList();

                    WorkDateList = lData1.Select(x => (DateTime)x.WorkDate).Distinct().ToList();

                    foreach (DateTime dt in WorkDateList)
                    {
                        var lData2 = lData1.Where(x => (DateTime)x.WorkDate == dt).ToList();

                        MaterialCode = lData2.Select(x => x.MaterialCode).ToList();

                        Val1 = MaterialCode.Count();
                        MaterialCode = MaterialCode.Distinct().ToList();
                        Val2 = MaterialCode.Count();

                        if (Val1 != Val2)
                        {
                            //  ErrorString = ErrorString +" " +dt.ToString("dd.MM.yyyy")+RowSplitter  + " " + pr + RowSplitter + " " + lRec.Where(x=>x.ProductCode==pr).Select(x => x.ProductName).First() + RowSplitter + " " + ll + RowSplitter + " Дублирование компонентов в рецепте! Дата рецепта: " + dt.ToString("dd.MM.yyyy") + RowSplitter + System.Environment.NewLine;
                            CommonEGrid.Add(new ErrorGrid
                            {
                                DateWork = dt,
                                ProdCode = " " + pr,
                                ProdName = lRec.Where(x => x.ProductCode == pr).Select(x => x.ProductName).First(),
                                LineName = ll,
                                Comment = " Дублирование компонентов в рецепте! Дата рецепта: " + dt.ToString("dd.MM.yyyy")
                            });
                        }

                        if (Val1 == 0)
                        {
                            //  ErrorString = ErrorString + " " + dt.ToString("dd.MM.yyyy") + RowSplitter + " " + pr + RowSplitter + " " + lRec.Where(x => x.ProductCode == pr).Select(x => x.ProductName).First() + RowSplitter + " " + ll + RowSplitter + " Рецепт отсутствует! Дата запроса рецепта:" + dt.ToString("dd.MM.yyyy") + RowSplitter + System.Environment.NewLine;
                            CommonEGrid.Add(new ErrorGrid
                            {
                                DateWork = dt,
                                ProdCode = " " + pr,
                                ProdName = lRec.Where(x => x.ProductCode == pr).Select(x => x.ProductName).First(),
                                LineName = ll,
                                Comment = " Рецепт отсутствует! Дата запроса рецепта:" + dt.ToString("dd.MM.yyyy")
                            });
                        }
                    }
                }
            }

            /* foreach (var prl in ProdRecipeList)
             {
                 foreach (var rll in prl.RecLineList)
                 {
                     foreach (var rl in rll.RecList)
                     {
                        // foreach (var r in rl.rList)
                         {
                             List<string> SList = rl.rList.Select(x => x.MaterialCode).ToList();

                             Val1 = SList.Count;
                             SList = SList.Distinct().ToList();
                             Val2 = SList.Count;

                             if (Val1!=Val2)
                             {
                                 ErrorString = ErrorString  + rll.LineNumber + " "+prl.ProductCode + " " + prl.ProductName + " дублирование компонентов в рецепте! Дата рецепта: " + rl.WorkDateStart.ToString("dd.MM.yyyy") + System.Environment.NewLine;
                             }

                             if (Val1==0)
                             {
                                 ErrorString = ErrorString + rll.LineNumber + " " + prl.ProductCode + " " + prl.ProductName  + " рецепт отсутствует! Дата рецепта:" + rl.WorkDateStart.ToString("dd.MM.yyyy") + System.Environment.NewLine;
                             }
                         }
                     }
                 }
             }*/

            if  (CommonEGrid.Count>0)     //(ErrorStr.Length>0)
            {
                ErrorStr = " " + RowSplitter + " " + RowSplitter + "Задвоенные рецепты по состоянию на " + DateTime.Today.Date.ToString("dd.MM.yyyy") + RowSplitter + System.Environment.NewLine; // + " " + RowSplitter + " " + RowSplitter + System.Environment.NewLine;  + ErrorString;
                if (Application.OpenForms["Warning1"] == null)
                {
                    Cursor = Cursors.WaitCursor;
                    w1 = new Warning1(this);
                    w1.Visible = true;
                    //   fs.MdiParent = this;
                    w1.WindowState = FormWindowState.Normal;
                    Cursor = Cursors.Default;
                }
                else
                {
                    w1.BringToFront();
                    w1.Visible = true;
                    //   cf.UpdateData();
                }
                this.Enabled = false;

                /*foreach (var R in Result)
                {
                    ErrorString = ErrorString + R.ProductCode + " " + R.LineNumber + " " + R.ProductName + ((DateTime)R.WorkDate).ToString("dd.MM.yyyy") + System.Environment.NewLine;
                }*/

                //  NotepadHelper.ShowMessage(ErrorString, "Сообщение системы");
            }
            else
            {
                MessageBox.Show("Задвоенных норм нет!", "Сообщение системы");
            }
        }

        private void dataGridView1_RowPrePaint(object sender, DataGridViewRowPrePaintEventArgs e)
        {
           /* foreach (DataGridViewRow r in dataGridView1.Rows)
            {
                foreach (DataGridViewCell c in r.Cells)
                {
                     if (c.Value.ToString().Contains("специй"))
                    {
                        c.Style.ForeColor = Color.MediumSeaGreen;
                    }
                    else if (c.Value.ToString().Contains("П/УП"))
                    {
                        c.Style.ForeColor = Color.DodgerBlue;
                    }
                }
            }*/
        }

        private void button10_Click(object sender, EventArgs e)
        {
            if (ExcelDataList.Count > 0)
            {
                LoadFlag = true;
                System.Threading.Thread thread1 = new System.Threading.Thread(WaitingShowFormWW);
                thread1.Start();

                if (Application.OpenForms["ConsurmptionForm"] == null)
                {
                    Cursor = Cursors.WaitCursor;
                    cf = new ConsurmptionForm(this);
                    cf.Visible = true;
                    //   fs.MdiParent = this;
                    cf.WindowState = FormWindowState.Normal;
                    Cursor = Cursors.Default;
                }
                else
                {
                    cf.BringToFront();
                    cf.Visible = true;
                    //   cf.UpdateData();
                }

                try
                {
                    LoadFlag = false;
                    System.Threading.Thread.Sleep(250);
                    thread1.Abort();
                }
                catch (Exception xx)
                { }
            }
            else
            {
                MessageBox.Show("Отсутствуют данные для обработки!", "Сообщение системы");
            }
        }

        private void button9_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["StorageForm"] == null)
            {
                Cursor = Cursors.WaitCursor;
                sf = new StorageForm(this);
                sf.Visible = true;
                //   fs.MdiParent = this;
                sf.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                sf.BringToFront();
                sf.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void button13_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["Form_OPR"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fo = new Form_OPR(this);
                fo.Visible = true;
                //   fs.MdiParent = this;
                fo.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fo.BringToFront();
                fo.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        public void SetBorders(Microsoft.Office.Interop.Excel.Range r, int width = 1, bool insideVert = false)
        {
            if (width == 1)
            {
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;

                if (insideVert == true)
                {
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel1.XlLineStyle.xlContinuous;
                }
            }
            else
            {
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].Weight = width;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].Weight = width;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].Weight = width;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].Weight = width;
                if (insideVert == true)
                {
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].Weight = width;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].Weight = width;
                }
            }
        }

        private void ExportRecipes()
        {
            // список рецептов
            Invoke(new Action(() =>
            {
                Cursor = Cursors.WaitCursor;

                progressBar1.Visible = true;
                progressBar1.Value = 1;

                label1.Visible = true;
                label1.Text = "Формирование Xls-файла... ";
            }));

            Int32 RecCount = 0;

            Invoke(new Action(() =>
            {
                foreach (var prl in ProdRecipeList)
                {
                    foreach (var rll in prl.RecLineList)
                    {
                        foreach (var rl in rll.RecList)
                        {
                            foreach (var r in rl.rList)
                            {
                                RecCount = RecCount + 1;
                            }
                        }
                    }
                }
            }));

            Excel1.Application xlApp;
            object misValue = System.Reflection.Missing.Value;
            Excel1.Range range1;
            Excel1.Workbook wBook;
            Excel1.Worksheet wSheet;

            xlApp = new Excel1.Application();
            xlApp.Visible = false;
            xlApp.SheetsInNewWorkbook = 1;
            wBook = xlApp.Workbooks.Add(misValue);
            xlApp.DisplayAlerts = false;
            wSheet = (Excel1.Worksheet)xlApp.Worksheets.get_Item(1);

            Invoke(new Action(() =>
            {
                Cursor = Cursors.WaitCursor;

                progressBar1.Visible = true;
                progressBar1.Value = 10;

                label1.Text = "Начало выгрузки данных ";
            }));

            Invoke(new Action(() =>
            {
                wSheet.Cells[1, 1] = "Код Продукта";
                wSheet.Cells[1, 2] = "Название продукта";
                wSheet.Cells[1, 3] = "Код Материала";
                wSheet.Cells[1, 4] = "Название материала";
                wSheet.Cells[1, 5] = "Единица измерения";
                wSheet.Cells[1, 6] = "Количество";
                wSheet.Cells[1, 7] = "Линия    ";
                wSheet.Cells[1, 8] = "Дата начала";
                wSheet.Cells[1, 9] = "Дата окончания";
                label1.Text = "Начало выгрузки данных.. ";
            }));

            Invoke(new Action(() =>
            {
                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 9]];
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;
                range1.Font.Bold = true;
                range1.Columns.AutoFit();
                SetBorders(range1, 1, true);
                label1.Text = "Начало выгрузки данных.... ";
            }));

            Int32 CurRow = 2;

            foreach (var prl in ProdRecipeList)
            {
                foreach (var rll in prl.RecLineList)
                {
                    foreach (var rl in rll.RecList)
                    {
                        foreach (var r in rl.rList)
                        {
                            Invoke(new Action(() =>
                            {
                                wSheet.Cells[CurRow, 1] = prl.ProductCode;
                                wSheet.Cells[CurRow, 2] = prl.ProductName;
                                wSheet.Cells[CurRow, 7] = rll.LineNumber;
                                wSheet.Cells[CurRow, 3] = r.MaterialCode;
                                wSheet.Cells[CurRow, 4] = r.MaterialName;
                                wSheet.Cells[CurRow, 5] = r.UMC;
                                wSheet.Cells[CurRow, 6] = r.MaterialQuantity;
                                wSheet.Cells[CurRow, 8] = rl.WorkDateStart;
                                wSheet.Cells[CurRow, 9] = rl.WorkDateEnd;

                                CurRow = CurRow + 1;
                            }));

                            if (CurRow % 100 == 0)
                            {
                                Invoke(new Action(() =>
                                {
                                    label1.Text = "Выгружено " + CurRow.ToString() + " из " + RecCount.ToString();
                                    progressBar1.Value = 10 + (CurRow * 80 / RecCount);
                                }));
                            }
                        }
                    }
                }
            }
            
            CurRow = CurRow - 1;
            Invoke(new Action(() =>
            {
                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CurRow, 1]];
                range1.NumberFormat = "@";

                range1 = wSheet.Range[wSheet.Cells[2, 3], wSheet.Cells[CurRow, 3]];
                range1.NumberFormat = "@";

                range1 = wSheet.Range[wSheet.Cells[2, 2], wSheet.Cells[CurRow, 2]];
                range1.ColumnWidth = 100;
                range1.WrapText = true;
                label1.Text = "Расстановка переносов";
                progressBar1.Value = 99;
            }));

            Invoke(new Action(() =>
                {
                    range1 = wSheet.Range[wSheet.Cells[2, 7], wSheet.Cells[CurRow, 7]];
                    range1.ColumnWidth = 20;

                    range1 = wSheet.Range[wSheet.Cells[2, 4], wSheet.Cells[CurRow, 4]];
                    range1.ColumnWidth = 100;
                    range1.WrapText = true;

                    range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CurRow, 9]];
                    SetBorders(range1, 1, true);

                    label1.Text = "Bсе почти готово";
                    progressBar1.Value = 99;
                }));

            Invoke(new Action(() =>
            {
                xlApp.Visible = true;
                wSheet = null;
                range1 = null;
                wBook = null;
                //   xlApp.Quit();
                xlApp = null;
            }));

            Invoke(new Action(() =>
            {
                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
            }));

            Invoke(new Action(() =>
            {
                Cursor = Cursors.Default;
                progressBar1.Visible = false;
                progressBar1.Value = 1;
                label1.Visible = false;
                panel3.Enabled = true;
                panel1.Enabled = true;
            }));
        }

        private void button14_Click(object sender, EventArgs e)
        {
            panel1.Enabled = false;
            panel3.Enabled = false;
            Thread thread = new Thread(ExportRecipes);
            thread.Start();
        }

        private void button17_Click(object sender, EventArgs e)
        {
            LoadRecipes();
            string s = "";
            ErrorMainStr = "";

            if (ExcelDataList.Count>0)
            {
                 s= GetPecipesNotes();
            }

            if ((ErrorMainStr.Length > 0) || (s.Length > 0))
            {
                ErrorMainStr = ErrorMainStr + System.Environment.NewLine + System.Environment.NewLine + s;
                if (Application.OpenForms["Warning1"] == null)
                {
                    Cursor = Cursors.WaitCursor;
                    w1 = new Warning1(this);
                    w1.Visible = true;
                    //   fs.MdiParent = this;
                    w1.WindowState = FormWindowState.Normal;
                    Cursor = Cursors.Default;
                }
                else
                {
                    w1.BringToFront();
                    w1.Visible = true;
                    //   cf.UpdateData();
                }
                this.Enabled = false;
            }
        }

        public StockBalancesPlant NewStockBalancesPlant(Context.SelectStockBalancesPlantBis_Result ssbpr)
        {
            StockBalancesPlant sbp = new StockBalancesPlant();

            sbp.DayPeriod = ssbpr.DayPeriod;
            sbp.WorkDate = ssbpr.WorkDate;
            sbp.MaterialCode = ssbpr.MaterialCode;
            sbp.MaterialName = ssbpr.MaterialName;
         //   sbp.MaterialLotName = ssbpr.MaterialLotName;
            sbp.ValidFrom = ssbpr.ValidFrom;
            sbp.ValidTo = ssbpr.ValidTo;
            sbp.Quantity =(double) ssbpr.Quantity;
            sbp.AlterLotName = ssbpr.AlterLotName;
            sbp.TestQuality = ssbpr.TestQuality;
            sbp.TestQualityGr = ssbpr.TestQualityGr;

            if (ssbpr.AlterLineName.Trim() != "")
            {
                sbp.LineName = ssbpr.AlterLineName;
            }
            else
            {
                sbp.LineName = ssbpr.LineName;
            }

            if (ssbpr.AlterLotName.Trim()!="")
            {
                sbp.MaterialLotName = ssbpr.AlterLotName;
            }
            else
            {
                sbp.MaterialLotName = ssbpr.MaterialLotName;
            }

            return sbp;
        }

        public StockBalancesMes NewStockBalancesMes(Context.SelectStockBalancesMesBis_Result ssbmr)
        {
            StockBalancesMes sbm = new StockBalancesMes();

            sbm.DayPeriod = ssbmr.DayPeriod;
            sbm.WorkDate = ssbmr.WorkDate;
            sbm.MaterialCode = ssbmr.MaterialCode;
            sbm.MaterialName = ssbmr.MaterialName;
            sbm.MaterialLot = ssbmr.LotName;
            sbm.Quantity = (double)ssbmr.Quantity;
            sbm.Owner = ssbmr.MaterialOwner;
            sbm.Storage = ssbmr.Storage;
            sbm.ProdDate = ssbmr.ProdDate;
            sbm.BBFDate = ssbmr.BBFDate;
            sbm.TestQuality = ssbmr.TestQuality;
            sbm.TestQualityGr = ssbmr.TestQualityGr;

            return sbm;
        }

        public StockBalances NewStockBalance(Context.SelectStockBalances1_Result ssbr)
        {
            StockBalances sb = new StockBalances();

            sb.bbf = ssbr.bbf;
            sb.BBFDate = ssbr.BBFDate;
            sb.DayPeriod = ssbr.DayPeriod;
            sb.difcurdates = ssbr.difcurdates;
            sb.difdates = ssbr.difdates;
            sb.dpr = ssbr.dpr;
            sb.LotDescription = ssbr.LotDescription;
            sb.LotName = ssbr.LotName;
            sb.MaterialCode = ssbr.MaterialCode;
            sb.MaterialName = ssbr.MaterialName;
            sb.MaterialOwner = ssbr.MaterialOwner;
            sb.MatGroup = ssbr.MatGroup;
            sb.ProdDate = ssbr.ProdDate;
            sb.Quantity = ssbr.Quantity;
            sb.st = ssbr.st;
            sb.Storage = ssbr.Storage;
            sb.StorageType = ssbr.StorageType;
            sb.TestOnStorage = ssbr.TestOnStorage;
            sb.TestQuality = ssbr.TestQuality;
            sb.TestQualityGr = ssbr.TestQualityGr;
            sb.WorkDate = ssbr.WorkDate;
            sb.Storage = ssbr.Storage;
            sb.StorageType = ssbr.StorageType;

            return sb;
        }

        private void button11_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormNavData"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fnd = new FormNavData(this);
                fnd.Visible = true;
                //   fs.MdiParent = this;
                fnd.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fnd.BringToFront();
                fnd.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void button12_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["MesSettingsForm"] == null)
            {
                Cursor = Cursors.WaitCursor;
                msf = new MesSettingsForm(this);
                msf.Visible = true;
                //   fs.MdiParent = this;
                msf.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                msf.BringToFront();
                msf.Visible = true;
                //   cf.UpdateData();
            }
            this.Enabled = false;
        }

        private void button18_Click(object sender, EventArgs e)
        {
            /*DateTime dt1 = DateTime.Today.Date;
            DateTime dt2 = new DateTime(2022, 5, 1);
            MessageBox.Show(dt2.Subtract(dt1).Days.ToString());*/

            lData LD;
            List<lData> DataListTempNew = new List<lData>();
            List<lData> DataListTempOld = new List<lData>();

            string Code = "1010033633";

            var Recipe = ProdRecipeList.Where(x => x.ProductCode.Trim() == Code).ToList();
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
                            LD = new lData();

                            LD.ProdCode = RL.MaterialCode.Trim();
                           // LD.ParentCode = RL.p
                            LD.ProdCodeStr = RL.MaterialCode.Trim();
                            LD.DateWork = RecLineDate[0].WorkDateStart;
                            LD.ProdName = RL.MaterialName.Trim();
                            LD.Quantity = 1000 * RL.MaterialQuantity;
                            LD.Line = RecLine[0].LineNumber;
                            LD.IPG = RL.IPG;
                            LD.Spices = false;   //control: split 1-st level only!!!!!!!!!
                            LD.BOM = RL.BOM;

                          bool  fl = false;

                            for (int i = 0; i < DataListTempNew.Count; i++)
                            {
                                if (DataListTempNew[i].ProdCode.Trim() == LD.ProdCode.Trim())
                                {
                                    if (DataListTempNew[i].Line.Trim() == LD.Line.Trim())
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

           int  Iteration = 0;
            var XLDataList = new List<lData>();

            //create Excel

            Excel1.Application xlApp = new Excel1.Application();
            xlApp.Visible = true;
            object misValue = System.Reflection.Missing.Value;

            Excel1.Workbook wBook = xlApp.Workbooks.Add(misValue);
            xlApp.DisplayAlerts = false;
            Excel1.Worksheet wSheet = (Excel1.Worksheet)wBook.Sheets[1];

            wSheet.Cells[1, 1] = "Передел";
            wSheet.Cells[1, 2] = "Код ГП Родитель";
            wSheet.Cells[1, 3] = "Код материала/ПФ";
            wSheet.Cells[1, 4] = "Наименование";
            wSheet.Cells[1, 5] = "Количество родитель";
            wSheet.Cells[1, 6] = "Норма";
            wSheet.Cells[1, 7] = "Количество материал";


            int CurrRow = 2;

            while (DataListTempNew.Count > 0)
            {
                Iteration = Iteration + 1;

                
             
             //   CurrRow = CurrRow + 1;

                for (int i = DataListTempNew.Count; i > 0; i--)
                {
                    if (DataListTempNew[i - 1].ProdCode == "1031016659")
                    { }

                     bool  fl = false;
                    for (int j = 0; j < XLDataList.Count; j++)
                    {
                        if (XLDataList[j].ProdCode.Trim() == DataListTempNew[i - 1].ProdCode.Trim())
                        {
                            if (XLDataList[j].DateWork == DataListTempNew[i - 1].DateWork)
                            {
                                if (XLDataList[j].Line.Trim() == DataListTempNew[i - 1].Line.Trim())
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

                    wSheet.Cells[CurrRow, 1] = Iteration;
                    wSheet.Cells[CurrRow, 2] = DataListTempNew[i - 1].ProdCode;
                    wSheet.Cells[CurrRow, 3] = DataListTempNew[i - 1].ProdName;
                    wSheet.Cells[CurrRow, 4] = DataListTempNew[i - 1].Quantity;
                    wSheet.Cells[CurrRow, 5] = DataListTempNew[i - 1].IPG;
                    CurrRow = CurrRow + 1;

                    if (fl == false)
                    {
                        LD = new lData
                        {
                            ProdCode = DataListTempNew[i - 1].ProdCode.Trim(),
                            ProdName = DataListTempNew[i - 1].ProdName.Trim(),
                            DateWork = DataListTempNew[i - 1].DateWork,
                            Quantity = DataListTempNew[i - 1].Quantity,
                            IPG = DataListTempNew[i - 1].IPG,
                            Line = DataListTempNew[i - 1].Line.Trim(),
                            Spices = DataListTempNew[i - 1].Spices,
                            BOM = DataListTempNew[i - 1].BOM
                        };

                        if (LD.Quantity > 0)         // ((LD.IPG != "ГП") && (LD.Quantity > 0))
                        {
                            XLDataList.Add(LD);
                        }
                    }

                    if (!DataListTempNew[i - 1].IPG.ToLower().Contains(@"п/ф"))         //(DataListTempNew[i - 1].IPG.ToLower() != "полуфабрикат")
                    {
                        DataListTempNew.RemoveAt(i - 1);
                    }
                    else
                    {
                        /*if (DataListTempNew[i - 1].Quantity == 0)
                        {
                            DataListTempNew.RemoveAt(i - 1);
                        }*/
                    }
                }

                ///
                DataListTempOld.Clear();
                DataListTempOld.AddRange(DataListTempNew);


             /*   wSheet.Cells[CurrRow, 1] = Iteration;
                CurrRow = CurrRow + 1;
                //iteration export
                for (int i = 0; i< DataListTempNew.Count; i++)
                {
                    wSheet.Cells[CurrRow, 1] = DataListTempNew[i].ProdCode;
                    wSheet.Cells[CurrRow, 2] = DataListTempNew[i].ProdName;
                    wSheet.Cells[CurrRow, 3] = DataListTempNew[i].Quantity;

                    CurrRow = CurrRow + 1;
                }*/

                DataListTempNew.Clear();

                ////

                for (int i = 0; i < DataListTempOld.Count; i++)
                {
                    var RecList = ProdRecipeList.Where(x => x.ProductCode.Trim() == DataListTempOld[i].ProdCode.Trim()).ToList();

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
                                        LD = new lData
                                        {
                                            ProdCode = RR.MaterialCode.Trim(),
                                            ProdCodeStr = RR.MaterialCode.Trim(),
                                            DateWork = DataListTempOld[i].DateWork,
                                            ProdName = RR.MaterialName.Trim(),
                                            Quantity = RR.MaterialQuantity * DataListTempOld[i].Quantity,
                                            Line = RecLineList[0].LineNumber.Trim(),    //DataListTempOld[i].Line.Trim(),
                                            IPG = RR.IPG,
                                            BOM = RR.BOM
                                        };

                                      bool fl = false;
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
                }
            }
        }

        private void button19_Click(object sender, EventArgs e)
        {
            var RList = ProdRecipeList.Where(x => (x.AlterProductCode != x.ProductCode)).ToList();
            if (RList.Count > 0)
            {
                Cursor = Cursors.WaitCursor;
                Int32 CurrRow = 2;

                Excel1.Application xlApp;
                object misValue = System.Reflection.Missing.Value;
                Excel1.Range range1;
                Excel1.Workbook wBook;
                Excel1.Worksheet wSheet;

                xlApp = new Excel1.Application();
                xlApp.Visible = true;
                xlApp.SheetsInNewWorkbook = 1;
                wBook = xlApp.Workbooks.Add(misValue);
                xlApp.DisplayAlerts = false;
                wSheet = (Excel1.Worksheet)xlApp.Worksheets.get_Item(1);

                wSheet.Cells[1, 1] = "Код Продукта";
                wSheet.Cells[1, 2] = "Код спецификации";
                wSheet.Cells[1, 3] = "Название продукта";

                range1 = wSheet.Range[wSheet.Cells[1, 1], wSheet.Cells[1, 3]];
                range1.HorizontalAlignment = Excel1.XlHAlign.xlHAlignCenter;
                range1.VerticalAlignment = Excel1.XlVAlign.xlVAlignCenter;
                range1.Font.Bold = true;
                range1.Columns.AutoFit();
                SetBorders(range1, 1, true);

                foreach (var r in RList)
                {
                    wSheet.Cells[CurrRow, 1] = r.ProductCode;
                    wSheet.Cells[CurrRow, 2] = r.AlterProductCode;
                    wSheet.Cells[CurrRow, 3] = r.ProductName;

                    CurrRow = CurrRow + 1;
                }

                CurrRow = CurrRow - 1;
                range1 = wSheet.Range[wSheet.Cells[2, 1], wSheet.Cells[CurrRow, 3]];
                range1.NumberFormat = "@";
                range1.Columns.EntireColumn.AutoFit();

                Cursor = Cursors.Default;
            }
            else
            {
                MessageBox.Show("Расхождение между Кодом ГП и Кодом спецификации отсутствуют!", "Сообщение системы");
            }

            Cursor = Cursors.Default;
        }

        private void button20_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormUnliquid"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fu = new FormUnliquid(this);
                fu.Visible = true;
                //   fs.MdiParent = this;
                fu.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                fu.BringToFront();
                fu.Visible = true;
                //   cf.UpdateData();
            }

            this.Enabled = false;
        }

        private void button21_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["MercuryForm"] == null)
            {
                Cursor = Cursors.WaitCursor;
                mf = new MercuryForm(this);
                mf.Visible = true;
                //   fs.MdiParent = this;
                mf.WindowState = FormWindowState.Normal;
                Cursor = Cursors.Default;
            }
            else
            {
                mf.BringToFront();
                mf.Visible = true;
                //   cf.UpdateData();
            }

            this.Enabled = false;
        }

      /*  public void SetBorders(Microsoft.Office.Interop.Excel.Range r, int width = 1, bool insideVert = false)
        {
            if (width == 1)
            {
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;

                if (insideVert == true)
                {
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel1.XlLineStyle.xlContinuous;
                }
            }
            else
            {
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeTop].Weight = width;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeRight].Weight = width;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeLeft].Weight = width;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].LineStyle = Excel1.XlLineStyle.xlContinuous;
                r.Borders.Item[Excel1.XlBordersIndex.xlEdgeBottom].Weight = width;
                if (insideVert == true)
                {
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideVertical].Weight = width;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].LineStyle = Excel1.XlLineStyle.xlContinuous;
                    r.Borders.Item[Excel1.XlBordersIndex.xlInsideHorizontal].Weight = width;
                }
            }
        }*/

        private void button22_Click(object sender, EventArgs e)
        {
            if (Application.OpenForms["FormCommdet"] == null)
            {
                Cursor = Cursors.WaitCursor;
                fc = new FormCommdet(this);
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

            this.Enabled = false;
        }

        private void dataGridView1_KeyDown(object sender, KeyEventArgs e)
        {
          /*  if (e.KeyCode.ToString() == "Add" || e.KeyCode.ToString() == "Oemplus")
            {
                MessageBox.Show("let`s OK");
            }*/
        }

        private void button23_Click(object sender, EventArgs e)
        {
            this.Close();
        }
    }

    public static class NotepadHelper
    {
        [DllImport("user32.dll", EntryPoint = "SetWindowText")]
        private static extern int SetWindowText(IntPtr hWnd, string text);

        [DllImport("user32.dll", EntryPoint = "FindWindowEx")]
        private static extern IntPtr FindWindowEx(IntPtr hwndParent, IntPtr hwndChildAfter, string lpszClass, string lpszWindow);

        [DllImport("User32.dll", EntryPoint = "SendMessage")]
        private static extern int SendMessage(IntPtr hWnd, int uMsg, int wParam, string lParam);

        public static void ShowMessage(string message, string title)
        {
            var notepad = Process.Start("notepad.exe");
            if (notepad != null)
            {
                notepad.WaitForInputIdle();

                SetWindowText(notepad.MainWindowHandle, title);
                var child = FindWindowEx(notepad.MainWindowHandle, new IntPtr(0), "Edit", null);
                SendMessage(child, 0x000C, 0, message);
            }
        }
    }
}
