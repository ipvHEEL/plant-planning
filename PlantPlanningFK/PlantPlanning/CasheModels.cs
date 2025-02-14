using System;
using System.Data;
using System.Data.SqlClient;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using PlantPlanning.Context;

namespace PlantPlanning
{
    public class DBDataGridView : DataGridView
    {
        public new bool DoubleBuffered
        {
            get { return base.DoubleBuffered; }
            set { base.DoubleBuffered = value; }
        }

        public DBDataGridView()
        {
            DoubleBuffered = true;
        }
    }

    public class ErrorGrid
    {
        public DateTime DateWork { get; set; }
        public string ProdCode { get; set; }
        public string ProdName { get; set; }
        public string LineName { get; set; }
        public string Comment { get; set; }
    }

    public class AppData
    {
        public long id { get; set; }
        public System.DateTime DateWork { get; set; }
        public Nullable<System.DateTime> AlterDate { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public double Quantity { get; set; }
    }

    public class MesAreaLocation
    {
        public Int64 MA_ID { get; set; }
        public string MA_AreaName { get; set; }
        public bool MA_IsApproved { get; set; }

        public Int64 MS_ID { get; set; }
        public Int64 MS_AreaID { get; set; }
        public string MS_LocationName { get; set; }
        public bool MS_IsApproved { get; set; }
        public Int64 MS_MesLocationID { get; set; }
        public string MS_MesLocationName { get; set; }
        public string MS_LocationType { get; set; }
    }

    public class UserSelectedCell
    {
        public int RowIndex { get; set; }
        public int ColIndex { get; set; }
        public string UserSelectedCode { get; set; }
        public string UserSelectedProdName { get; set; }
        public DateTime UserSelectedDateTime { get; set; }
        public double Quantity { get; set; }
    }

    public class PurchTable
    {
        public string OrderNymber { get; set; }  //storage
        public string LostDescr { get; set; }  //LotName
        public DateTime PlanOperDate { get; set; }  //Proddate
        public DateTime BBFDate { get; set; }
        public double PlanQuantity { get; set; }  //Q-tyStorage
        public int Count { get; set; }
        public string Status { get; set; }
    }

    public class StockBalancesResult
    {
        public DateTime WorkDate { get; set; }
        public string Storage { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public string MatGroup { get; set; }
        public string LotName { get; set; }
        public string LotDescription { get; set; }
        public DateTime ProdDate { get; set; }
        public DateTime BBFDate { get; set; }
        public string Status { get; set; }
        public double QuantityStorageDay { get; set; } = 0; //QuantityStorageNight -QuantityMesDelta
        public double QuantityStorageNight { get; set; } = 0;
        public double QuantityMesDay { get; set; } = 0;
        public double QuantityMesNight { get; set; } = 0;
        public double QuantityMesDelta { get; set; } = 0;//== QuantityMesNight -QuantityMesDay
        public double QuantityPlantNight { get; set; } = 0;
        public double QuantityPlantDay { get; set; } = 0;
        public double QuantityDayItog { get; set; } = 0; //= QuantityStorageNight -QuantityMesDelta +QuantityPlantDay //QuantityStorageDay +QuantityPlantDay
        // public List <StockBalancesResultList> SBRL { get; set; }
    }

    public class StockBalancesShort
    {
        public Int32 Quantity { get; set; }
        public string Code { get; set; }
        public string Group { get; set; }
        public string Name1 { get; set; }
        public string MinStatus { get; set; }  //false
        public double MinStatusValue { get; set; }  //false
        public List<StockBalancesDetails> SBDList { get; set; }
    }

    public class StockBalancesDetails
    {
        public string Lot { get; set; }
        public Int32 Quantity { get; set; }
        public string Storage { get; set; }
        public string ProdDate { get; set; }
        public string Status { get; set; }
        public string TestQuality { get; set; }
        public string Comment { get; set; }
        public int difDates { get; set; }  //false
        public int difCurDates { get; set; }  //false
        public double Difference { get; set; }  //false
    }

    public class StockBalances
    {
        //  public long id { get; set; }
        public string Storage { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public string MatGroup { get; set; }
        public string LotName { get; set; }
        public string LotDescription { get; set; }
        public decimal Quantity { get; set; }
        public string MaterialOwner { get; set; }
        public string TestQuality { get; set; }
        public string StorageType { get; set; }
        public string TestQualityGr { get; set; }
        public string TestOnStorage { get; set; }
        public DateTime ProdDate { get; set; }
        public DateTime BBFDate { get; set; }
        public int dpr { get; set; }
        public int bbf { get; set; }
        public int difdates { get; set; }
        public int difcurdates { get; set; }
        public string st { get; set; }
        public DateTime WorkDate { get; set; }
        public string DayPeriod { get; set; }
    }

    public class StockBalancesMes
    {
        //   public long id { get; set; }
        public string Storage { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public string MaterialLot { get; set; }
        public double Quantity { get; set; }
        public string Owner { get; set; }
        public DateTime WorkDate { get; set; }
        public string DayPeriod { get; set; }
        public DateTime ProdDate { get; set; }
        public DateTime BBFDate { get; set; }
        public string TestQuality { get; set; }
    }

    public class StockBalancesPlant
    {
        //    public long id { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public string MaterialLotName { get; set; }
        public string LineName { get; set; }
        public DateTime ValidFrom { get; set; }
        public DateTime ValidTo { get; set; }
        public double Quantity { get; set; }
        public DateTime WorkDate { get; set; }
        public string DayPeriod { get; set; }
        public string AlterLotName { get; set; }
        public string Status { get; set; }
    }

    public class PlantOperList
    {
        public DateTime DateWork { get; set; }
        public List<DataList> DL { get; set; }
    }

    public class Colorres
    {
        public string ProdCode { get; set; }
        public List<int> ColNo { get; set; }
    }

    public class DataList
    {
        public string Line { get; set; }
        public string ProdCode { get; set; }
        public string ProdCodeStr { get; set; }
        public string ProdName { get; set; }
        public double Quantity { get; set; }
        public bool RePack { get; set; }
        public bool LastRec { get; set; } = false;
    }

    public class lData
    {
        public DateTime DateWork { get; set; }
        public string Line { get; set; }
        public string ProdCode { get; set; }
        public string ProdCodeStr { get; set; }
        public string ProdName { get; set; }
        public string AlterProdName { get; set; }
        public double Quantity { get; set; }
        public bool RePack { get; set; }
        public bool Spices { get; set; } = false;
        public string IPG { get; set; }
    }

    public class CommentData_Result
    {
        public Nullable<long> ID { get; set; }
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public Nullable<System.DateTime> DateRecord { get; set; }
        public Nullable<bool> IsActual { get; set; }
        public string UserName { get; set; }
        public string Comment { get; set; }
    }

    public class Material
    {
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public string MaterialGroup { get; set; } //сырье/упаковка/ГП
        public string MaterialCategory { get; set; }
        public Int32 MaterialMultiply { get; set; }
        public string MaterialMultiplyString { get; set; } = "";
        public Int32 WaitingDays { get; set; }
        public string WaitingDaysString { get; set; }
        //   public Int64 LineID { get; set; }
        //   public string LineName { get; set; }
        //   public Int64 AreaID { get; set; }
        //   public string AreaName { get; set; }
        public double StorageQuantity { get; set; }
        public string StorageQuantityString { get; set; }
        public string Responsible { get; set; }
        public string Sender { get; set; }
        public List<UsingByPlans> ListUBP { get; set; } //использоание в планах
        public List<UsingByDate> ListUBD { get; set; } //использование на дату (по рецептам)
        public List<StorageQuantByDate> ListSQBD { get; set; } //остатки на складах по датам

        public double RawStorage { get; set; } //остатки склады
        public double RawEnterprise { get; set; } //остатки производство
        public double RawNav1 { get; set; } //остатки внешние  
        public double RawNav2 { get; set; }//остатки  собств по НАВ
        public double RawMesOut1 { get; set; } //забл.внешние  
        public double RawMesOut2 { get; set; } //забл (МЕС)
        public double OldTaskNav { get; set; } //старые заявки НАВ
        public double CurrentConsumption { get; set; } = 0; //уже загружено в наборы

        public string RawStorageS { get; set; }
        public string RawEnterpriseS { get; set; }
        public string RawNav1S { get; set; }
        public string RawNav2S { get; set; }
        public string RawMesOut1S { get; set; }
        public string RawMesOut2S { get; set; }
        public string OldTaskNavS { get; set; }
        public string CurrentConsumptionS { get; set; }

        public double RawStorageUD { get; set; } //остатки склады UnDelayed не просроченные!!!!
        public double RawEnterpriseUD { get; set; } //остатки производство
        public double RawNav1UD { get; set; } //остатки внешние  
        public double RawNav2UD { get; set; }//остатки  собств по НАВ
        public double RawMesOut1UD { get; set; } //забл.внешние  
        public double RawMesOut2UD { get; set; } //забл (МЕС)
        public double OldTaskNavUD { get; set; } //старые заявки НАВ   не просроченные!!!!     

        public string RawStorageUDS { get; set; }
        public string RawEnterpriseUDS { get; set; }
        public string RawNav1UDS { get; set; }
        public string RawNav2UDS { get; set; }
        public string RawMesOut1UDS { get; set; }
        public string RawMesOut2UDS { get; set; }
        public string OldTaskNavUDS { get; set; }

        public string ShortString { get; set; }
        public string ShortStringUD { get; set; }
        public string ValueS { get; set; }
        public string ValueUDS { get; set; }
        public string AlterValueS { get; set; }
        public string AlterValueUDS { get; set; }

        public double Value_ { get; set; }
        public double ValueUD { get; set; }

        public string Color1 { get; set; } = "Black";   //остатки склады
        public string Color2 { get; set; } = "Black";//остатки производство

        public string Color1UD { get; set; } = "Black";   //остатки склады UnDelay
        public string Color2UD { get; set; } = "Black";//остатки производство

        public List<Raws> RawStorageList { get; set; }  //остатки склады
        public List<Raws> RawEnterpriseList { get; set; }  //остатки производство
        public List<Raws> RawNav1List { get; set; }
        public List<Raws> RawNav2List { get; set; }
        public List<Raws> RawMesOut1List { get; set; }
        public List<Raws> RawMesOut2List { get; set; }
        public List<Raws> OldTaskNavList { get; set; }//старые заявки НАВ
        public List<Raws> TaskNavList1 { get; set; }  //наст и будущее заявки НАВ //приход ТОЛЬКО заявки НАВ учесть собственные в коде! оформляемые
        public List<Raws> TaskNavList2 { get; set; }  //наст и будущее заявки НАВ //приход ТОЛЬКО заявки НАВ учесть собственные в коде! заказанные и в пути
        //  public List<Recipe> ListRec { get; set; }
        public List<Raws> OutList { get; set; } //расход сырья по плану
        public List<Raws> FactList { get; set; } //расход сырья по факту
        public List<Raws> StorageQuantList { get; set; }  //proddate & Q-ty остаток на дату с учетом прихода и расхода
        public List<Raws> ConsurmptionList { get; set; }  //уже потреблено
        public List<Raws> ListA { get; set; }
        public List<Raws> ListB { get; set; }
        public List<Raws> ListC { get; set; }
        public List<Raws> ListD { get; set; }

        public bool DontShow { get; set; } = false;

        public double Month1 { get; set; } //потребление за 3 прошлых месяца
        public double Month2 { get; set; }
        public double Month3 { get; set; }

        public bool Green { get; set; } = false;
        public bool Pink { get; set; } = false;

        public List<int> ListPlus { get; set; }
        public List<int> ListLine { get; set; }
    }

    //потр на дату. для сырья...
    public class UsingByDate
    {
        public DateTime WorkDate { get; set; }
        public string WorkedString { get; set; }
        public double TotalQuant { get; set; }
        public List<UsingByDateList> ListUBD { get; set; }
    }

    public class UsingByDateList
    {
        public string PlanNumber { get; set; }
        public double Quant { get; set; }
    }

    //потребление по планам
    public class UsingByPlans
    {
        public DateTime WorkDate { get; set; }
        public Int64 LineID { get; set; }
        public string LineName { get; set; }
        public Int64 AreaID { get; set; }
        public string AreaName { get; set; }
        public string ProductCode { get; set; }
        public string ProductName { get; set; }
        public double ProductQuant { get; set; }
        public double MaterialQuant { get; set; }
        public double NormQuant { get; set; }
    }

    //остатки на складах
    public class StorageQuantByDate
    {
        public DateTime WorkDate { get; set; }
        public double TotalQuant { get; set; }
        public double TotalIncomeQuant { get; set; }
        public List<StorageQuant> ListSQ { get; set; }
    }

    public class StorageQuant
    {
        public string StorageType { get; set; }
        public List<Storages> ListSt { get; set; }
    }

    public class Storages
    {
        public string StorageName { get; set; }
        public double Quantity { get; set; }
        public double IncomeQuant { get; set; }
    }

    //список остатков
    public class Raws
    {
        public string Storage { get; set; } //Склад/Линия/№ заявки
        public string LotName { get; set; } //Lot
        public string LotDescr { get; set; }  //in StorageQuantList - результирующая запись
        public string MaterialCode { get; set; } = "";
        public string MaterialName { get; set; } = "";
        public double Quantity { get; set; }
        public DateTime ProdDate { get; set; }
        public DateTime BBFDate { get; set; }  //PlanOperDate
        public DateTime AlterDate { get; set; }
        public DateTime AddDate { get; set; }
        public string Color { get; set; }

        public string Initiator { get; set; }
        public string StatusMZP { get; set; }
        public string MZP { get; set; }
    }

    public class DailyResult
    {
        public string LineName { get; set; }
        public string ProductCode { get; set; }
        public double Quantity { get; set; }
        public string MaterialType { get; set; }
        public int Location { get; set; }
    }

    //список рецептов. заголовки 
    public class ProductRecipes
    {
        public string ProductCode { get; set; }
        public string AlterProductCode { get; set; }
        public string ProductCodeString { get; set; }
        public string ProductName { get; set; }
        public bool RePack { get; set; } = false;
        public DateTime DStart { get; set; }
        public List<ProductLinesRecipes> RecLineList { get; set; }
    }

    public class ProductLinesRecipes  //рецепты на дату!!!!! 
    {
        public string LineNumber { get; set; }
        public List<Recipe> RecList { get; set; }
    }

    //потребность в сырье на тонну ГП. по линиям!!!!! для конкретного продукта
    public class Recipe
    {
        public DateTime WorkDateStart { get; set; }
        public DateTime WorkDateEnd { get; set; } = DateTime.Today.Date.AddYears(25);
        public List<RecipeList> rList { get; set; }
    }

    //список компонентов на дату. по линии
    public class RecipeList
    {
        public string MaterialCode { get; set; }
        public string MaterialName { get; set; }
        public double MaterialQuantity { get; set; }
        public string UMC { get; set; }
        public string IPG { get; set; }
    }





    public interface IDataPageRetriever
    {
        DataTable SupplyPageOfData(int lowerPageBoundary, int rowsPerPage , int ColumnIndex);
        int GetRowCount();
    }



    public class DataRetriever : IDataPageRetriever
    {
        private string tableName;
        private DataTable tTable;
       private FormMain fm;
       private ConsurmptionForm cf;


        private SqlCommand command;

        public DataRetriever (DataTable TableName, FormMain FM, ConsurmptionForm CF)     //(string connectionString, string tableName)
        {
            /* SqlConnection connection = new SqlConnection(connectionString);
             connection.Open();
             command = connection.CreateCommand();*/
            this.tTable = TableName; //.TableName;
            this.fm = FM;
            this.cf = CF;
        }

       

      /*  public GetDataTable (DataTable TableName)
        {
           this.tTable = TableName;
        }*/

        private int rowCountValue = -1;

        public int RowCount
        {
            get
            {
                // Return the existing value if it has already been determined.
                if (rowCountValue != -1)
                {
                    return rowCountValue;
                }
                else
                {
                    rowCountValue = tTable.Rows.Count;
                    return  rowCountValue;      //tTable.Rows.Count;
                }
            }
        }

        private int columnIndex = -1;

        public int ColumnIndex
        {
            get
            {
               /* if (columnIndex!= -1)
                {
                    return columnIndex;
                }*/

                return columnIndex;  //заменить на столбец, по которому идет сортировка
            }
            set
            {
                columnIndex = ColumnIndex;
            }
        }

        private DataColumnCollection columnsValue;

        public DataColumnCollection Columns
        {
            get
            {
                // Return the existing value if it has already been determined.
                if (columnsValue != null)
                {
                    return columnsValue;
                }

                // Retrieve the column information from the database.
                /*   command.CommandText = "SELECT * FROM " + tableName;
                   SqlDataAdapter adapter = new SqlDataAdapter();
                   adapter.SelectCommand = command;*/
                DataTable table = tTable;   //new DataTable();
                table.Locale = System.Globalization.CultureInfo.InvariantCulture;
              //  adapter.FillSchema(table, SchemaType.Source);
                columnsValue = table.Columns;
                return columnsValue;
            }
        }

        private string commaSeparatedListOfColumnNamesValue = null;

        private string CommaSeparatedListOfColumnNames
        {
            get
            {
                // Return the existing value if it has already been determined.
                if (commaSeparatedListOfColumnNamesValue != null)
                {
                    return commaSeparatedListOfColumnNamesValue;
                }

                // Store a list of column names for use in the
                // SupplyPageOfData method.
                System.Text.StringBuilder commaSeparatedColumnNames =
                    new System.Text.StringBuilder();
                bool firstColumn = true;
                foreach (DataColumn column in Columns)
                {
                    if (!firstColumn)
                    {
                        commaSeparatedColumnNames.Append(", ");
                    }
                    commaSeparatedColumnNames.Append(column.ColumnName);
                    firstColumn = false;
                }

                commaSeparatedListOfColumnNamesValue =
                    commaSeparatedColumnNames.ToString();
                return commaSeparatedListOfColumnNamesValue;
            }
        }

        // Declare variables to be reused by the SupplyPageOfData method.
        private string columnToSortBy;
        private SqlDataAdapter adapter = new SqlDataAdapter();

        public int GetRowCount()
        {
            return RowCount;
        }

        public DataTable SupplyPageOfData(int lowerPageBoundary, int rowsPerPage, int ColumnIndex)
        {
            // Store the name of the ID column. This column must contain unique
            // values so the SQL below will work properly.
            columnToSortBy = this.Columns[ColumnIndex].ColumnName;  //0

            DataTable t1 = tTable;
         //   t1.DefaultView.Sort = t1.Columns[columnToSortBy].ToString();
         //   t1 = t1.DefaultView.ToTable();

            DataRow[] rows = t1.Select();

            DataTable table = tTable.Clone();     //new DataTable();
            string[] NewString;
            char RowSplitter = '|';
            string sVal = "";
            int MonthCount = cf.WorkMonthList.Count + 9;
            double Value;
            int Cols;
            int TableRecCount = RowCount;

          /*  cf.CFMListGold = new List<string>();
            cf.CFMListGray = new List<string>();
            cf.CFMListLine = new List<Colorres>(); //[
            cf.CFMListPink = new List<string>();//[
            cf.CFMListPlus = new List<Colorres>(); //+
            cf.CFMListGreen = new List<string>();//+
            cf.CFMListMinus = new List<Colorres>();//-
            cf.CFMListRed = new List<string>(); //-
            cf.CFMListLinePlus = new List<Colorres>();  //[+
            cf.CFMListGoldens = new List<string>();  //[+*/

            for (int i = lowerPageBoundary; i < lowerPageBoundary + rowsPerPage; i++)
            {
                sVal = "";
                Value = 0;
                if (i <= TableRecCount)
                {
                    try
                    {
                        // object[] copyrow = rows[i].ItemArray;
                        //  Int32 j = 0;

                        //   foreach (var d in rows[i].ItemArray)
                        for (int j = 0; j < MonthCount; j++)
                        {
                            if (j < MonthCount)  //0-8
                            {
                                sVal = sVal + rows[i].ItemArray[j].ToString() + RowSplitter;
                            }
                        }
                        //9-13
                        if (cf.cbMaterialDelay.Checked == false)
                        {
                            if ((Math.Abs(cf.CFMList[i].RawStorage) <= 1) && (cf.CFMList[i].RawStorage != 0)) //9
                            {
                                sVal = sVal + cf.CFMList[i].RawStorage.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawStorage.ToString("N0") + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].RawEnterprise) <= 1) && (cf.CFMList[i].RawEnterprise != 0))  //10
                            {
                                sVal = sVal + cf.CFMList[i].RawEnterprise.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawEnterprise.ToString("N0") + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].RawNav1) <= 1) && (cf.CFMList[i].RawNav1 != 0)) //внешние
                            {
                                sVal = sVal + cf.CFMList[i].RawNav1.ToString("N1") + " (";
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawNav1.ToString("N0") + " (";
                            }

                            if ((Math.Abs(cf.CFMList[i].RawNav2) <= 1) && (cf.CFMList[i].RawNav2 != 0))  //собств по НАВ   11
                            {
                                sVal = sVal + cf.CFMList[i].RawNav2.ToString("N1") + ") " + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawNav2.ToString("N0") + ") " + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].RawMesOut1) <= 1) && (cf.CFMList[i].RawMesOut1 != 0)) //забл внешние
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut1.ToString("N1") + " (";
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut1.ToString("N0") + " (";
                            }

                            if ((Math.Abs(cf.CFMList[i].RawMesOut2) <= 1) && (cf.CFMList[i].RawMesOut2 != 0))   // забл  МЕС  12
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut2.ToString("N1") + ")" + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut2.ToString("N0") + ")" + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].OldTaskNav) <= 1) && (cf.CFMList[i].OldTaskNav != 0))  //13 старые заявки НАВ
                            {
                                sVal = sVal + cf.CFMList[i].OldTaskNav.ToString("N1") + RowSplitter;   //0-13
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].OldTaskNav.ToString("N0") + RowSplitter;   //0-13
                            }
                        }
                        else
                        {
                            if ((Math.Abs(cf.CFMList[i].RawStorageUD) <= 1) && (cf.CFMList[i].RawStorageUD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].RawStorageUD.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawStorageUD.ToString("N0") + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].RawEnterpriseUD) <= 1) && (cf.CFMList[i].RawEnterpriseUD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].RawEnterpriseUD.ToString("N1") + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawEnterpriseUD.ToString("N0") + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].RawNav1UD) <= 1) && (cf.CFMList[i].RawNav1UD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].RawNav1UD.ToString("N1") + " (";
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawNav1UD.ToString("N0") + " (";
                            }

                            if ((Math.Abs(cf.CFMList[i].RawNav2UD) <= 1) && (cf.CFMList[i].RawNav2UD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].RawNav2UD.ToString("N1") + ") " + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawNav2UD.ToString("N0") + ") " + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].RawMesOut2UD) <= 1) && (cf.CFMList[i].RawMesOut2UD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut2UD.ToString("N1") + " (";
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut2UD.ToString("N0") + " (";
                            }

                            if ((Math.Abs(cf.CFMList[i].RawMesOut1UD) <= 1) && (cf.CFMList[i].RawMesOut1UD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut1UD.ToString("N1") + ") " + RowSplitter;
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].RawMesOut1UD.ToString("N0") + ") " + RowSplitter;
                            }

                            if ((Math.Abs(cf.CFMList[i].OldTaskNavUD) <= 1) && (cf.CFMList[i].OldTaskNavUD != 0))
                            {
                                sVal = sVal + cf.CFMList[i].OldTaskNavUD.ToString("N1") + RowSplitter;   //0-12
                            }
                            else
                            {
                                sVal = sVal + cf.CFMList[i].OldTaskNavUD.ToString("N0") + RowSplitter;   //0-12
                            }
                        }//delayChecked==false

                        if (cf.cbViews.Text == "Остаток")
                        {
                            Value = cf.CFMList[i].CurrentConsumption + cf.CFMList[i].RawNav1;    //уже портеблено + внешние остатки по НАВ

                            if (cf.cbMaterialDelay.Checked == false)
                            {
                                foreach (var sl in cf.CFMList[i].RawStorageList)  //остатки СКЛАД, остатки ПР-ВО
                                {
                                    Value = Value + sl.Quantity;
                                }

                                foreach (var el in cf.CFMList[i].RawEnterpriseList)
                                {
                                    Value = Value + el.Quantity;
                                }
                            }
                            else
                            {
                                foreach (var sl in cf.CFMList[i].RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                                {
                                    Value = Value + sl.Quantity;
                                }

                                foreach (var el in cf.CFMList[i].RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                                {
                                    Value = Value + el.Quantity;
                                }
                            }

                            Cols = 14 + cf.WorkMonthList.Count;
                            foreach (DateTime dt in cf.DTList)
                            {
                                //   if (CFMList[i].MaterialCode == "1031008522")
                                //   { }

                                //income
                                double T2 = 0;
                                double T1 = 0;

                                var FDTList = cf.DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();

                                if (cf.cbUseMaterialPlanning.Checked == true)
                                {
                                    List<Raws> tData2 = new List<Raws>();

                                    if (cf.cbAddDay.Checked == false)
                                    {
                                        tData2 = cf.CFMList[i].TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                                    }
                                    else
                                    {
                                        tData2 = cf.CFMList[i].TaskNavList2.Where(x => x.AddDate == dt).ToList();
                                    }

                                    if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                                    {
                                        tData2 = cf.CFMList[i].TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                                    }

                                    foreach (var td2 in tData2)
                                    {
                                        T2 = T2 + td2.Quantity;
                                    }

                                    List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                                    if (cf.cbAddDay.Checked == false)
                                    {
                                        Cdata = fm.ContList.Where(x => (x.MaterialCode == cf.CFMList[i].MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                                    }
                                    else
                                    {
                                        Cdata = fm.ContList.Where(x => (x.MaterialCode == cf.CFMList[i].MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                                    }

                                    if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                                    {
                                        Cdata = fm.ContList.Where(x => (x.MaterialCode == cf.CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                                    }

                                    foreach (var cd in Cdata)
                                    {
                                        T2 = T2 + cd.Quantity;
                                    }
                                    Value = Value + T2;

                                    //ask in progress
                                    List<Raws> tData1 = new List<Raws>();

                                    if (cf.cbAddDay.Checked == false)
                                    {
                                        tData1 = cf.CFMList[i].TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                                    }
                                    else
                                    {
                                        tData1 = cf.CFMList[i].TaskNavList1.Where(x => x.AddDate == dt).ToList();
                                    }

                                    if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                                    {
                                        tData1 = cf.CFMList[i].TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                                    }

                                    foreach (var td1 in tData1)
                                    {
                                        T1 = T1 + td1.Quantity;
                                    }
                                    Value = Value + T1;
                                }

                                //credit
                                if (cf.cbUsingMaterial.Checked == true)
                                {
                                    var dll = cf.XLDataList.Where(x => (x.ProdCode == cf.CFMList[i].MaterialCode) && (x.DateWork.Date == dt)).ToList();
                                    foreach (var dl in dll)
                                    {
                                        Value = Value - dl.Quantity;
                                    }
                                }

                                //future planning Asks
                                double Q = 0;
                                var adata = cf.AppList;    // fm.tdb.tApplication.ToList();
                                                           //  if (adata.Count > 0)
                                {
                                    adata = adata.Where(x => x.MaterialCode == cf.CFMList[i].MaterialCode).ToList();
                                    //    if (adata.Count > 0)
                                    {
                                        adata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                        if ((adata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                        {
                                            adata = cf.AppList;
                                            adata = adata.Where(x => x.MaterialCode == cf.CFMList[i].MaterialCode).ToList();
                                            adata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                        }

                                        foreach (var ad in adata)
                                        {
                                            Q = Q + ad.Quantity;
                                        }

                                        cf.USCList.Add(new UserSelectedCell
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
                                    // if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                    {
                                        //  CFMBool[i] = false;
                                        if (!cf.CFMListGold.Contains(cf.CFMList[i].MaterialCode))  //-
                                        {
                                            cf.CFMListGold.Add(cf.CFMList[i].MaterialCode);
                                            cf.CFMListRed.Add(cf.CFMList[i].MaterialCode);
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cf.CFMList[i].MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            cf.CFMListMinus.Add(Col);
                                        }
                                        else
                                        {
                                            var Coll = cf.CFMListMinus.Where(x => x.ProdCode == cf.CFMList[i].MaterialCode).ToList();
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
                                                    ProdCode = cf.CFMList[i].MaterialCode,
                                                    ColNo = new List<int>()
                                                };
                                                Col.ColNo.Add(Cols);
                                                cf.CFMListMinus.Add(Col);
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
                                    if (!cf.CFMListGreen.Contains(cf.CFMList[i].MaterialCode))  //+
                                    {
                                        cf.CFMListGreen.Add(cf.CFMList[i].MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cf.CFMList[i].MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        cf.CFMListPlus.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = cf.CFMListPlus.Where(x => x.ProdCode == cf.CFMList[i].MaterialCode).ToList();
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
                                                ProdCode = cf.CFMList[i].MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            cf.CFMListPlus.Add(Col);
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

                                    if (!cf.CFMListPink.Contains(cf.CFMList[i].MaterialCode))  //[
                                    {
                                        cf.CFMListPink.Add(cf.CFMList[i].MaterialCode);
                                        Colorres Col = new Colorres
                                        {
                                            ProdCode = cf.CFMList[i].MaterialCode,
                                            ColNo = new List<int>()
                                        };
                                        Col.ColNo.Add(Cols);
                                        cf.CFMListLine.Add(Col);
                                    }
                                    else
                                    {
                                        var Coll = cf.CFMListLine.Where(x => x.ProdCode == cf.CFMList[i].MaterialCode).ToList();
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
                                                ProdCode = cf.CFMList[i].MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            cf.CFMListLine.Add(Col);
                                        }
                                    }

                                    if (cf.CFMListGreen.Contains(cf.CFMList[i].MaterialCode))  //[+
                                    {
                                        if (!cf.CFMListGoldens.Contains(cf.CFMList[i].MaterialCode))
                                        {
                                            cf.CFMListGoldens.Add(cf.CFMList[i].MaterialCode);
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cf.CFMList[i].MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            cf.CFMListLinePlus.Add(Col);
                                        }
                                        else
                                        {
                                            var Coll = cf.CFMListLinePlus.Where(x => x.ProdCode == cf.CFMList[i].MaterialCode).ToList();
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
                                                    ProdCode = cf.CFMList[i].MaterialCode,
                                                    ColNo = new List<int>()
                                                };
                                                Col.ColNo.Add(Cols);
                                                cf.CFMListLinePlus.Add(Col);
                                            }
                                        }
                                    }
                                }
                                if (Q != 0)
                                {
                                    if ((Math.Abs(Q) <= 1) && (Q != 0))
                                    {
                                        sVal = sVal + " {" + Q.ToString("N1") + "} ";
                                    }
                                    else
                                    {
                                        sVal = sVal + " {" + Q.ToString("N0") + "} ";
                                    }
                                }

                                sVal = sVal + "" + RowSplitter;
                                Cols = Cols + 1;
                            }
                        }  //остаток
                        else if (cf.cbViews.Text == "Потребление")
                        {
                            var xlData = cf.XLDataList.Where(x => x.ProdCode == cf.CFMList[i].MaterialCode).ToList();

                            foreach (DateTime dt in cf.DTList)
                            {
                                var xlD = xlData.Where(x => x.DateWork.Date == dt).ToList();

                                if (xlD.Count > 0)
                                {
                                    if ((Math.Abs(xlD[0].Quantity) <= 1) && (xlD[0].Quantity != 0))
                                    {
                                        sVal = sVal + xlD[0].Quantity.ToString("N1") + RowSplitter;
                                    }
                                    else
                                    {
                                        sVal = sVal + xlD[0].Quantity.ToString("N0") + RowSplitter;
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
                            Value = 0;

                            if (cf.cbMaterialDelay.Checked == false)
                            {
                                foreach (var sl in cf.CFMList[i].RawStorageList)
                                {
                                    Value = Value + sl.Quantity;
                                }

                                foreach (var el in cf.CFMList[i].RawEnterpriseList)
                                {
                                    Value = Value + el.Quantity;
                                }
                            }
                            else
                            {
                                foreach (var sl in cf.CFMList[i].RawStorageList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                                {
                                    Value = Value + sl.Quantity;
                                }

                                foreach (var el in cf.CFMList[i].RawEnterpriseList.Where(x => x.BBFDate >= DateTime.Today.Date).ToList())
                                {
                                    Value = Value + el.Quantity;
                                }
                            }

                            Cols = 14 + cf.WorkMonthList.Count;
                            foreach (DateTime dt in cf.DTList)
                            {
                                var FDTList = cf.DTList.Where(x => (x.Month == dt.Month) && (x.Year == dt.Year)).ToList();
                                //income
                                List<Raws> tData2 = new List<Raws>();
                                if (cf.cbAddDay.Checked == false)
                                {
                                    tData2 = cf.CFMList[i].TaskNavList2.Where(x => x.BBFDate == dt).ToList();
                                }
                                else
                                {
                                    tData2 = cf.CFMList[i].TaskNavList2.Where(x => x.AddDate == dt).ToList();
                                }

                                double T2 = 0;
                                if ((dt.Day == 1) && (FDTList.Count == 1) && (tData2.Count == 0))
                                {
                                    tData2 = cf.CFMList[i].TaskNavList2.Where(x => x.AlterDate == dt).ToList();
                                }

                                foreach (var td2 in tData2)
                                {
                                    T2 = T2 + td2.Quantity;
                                }

                                List<tPlannedMeatContainers> Cdata = new List<tPlannedMeatContainers>();

                                if (cf.cbAddDay.Checked == false)
                                {
                                    Cdata = fm.ContList.Where(x => (x.MaterialCode == cf.CFMList[i].MaterialCode) && (x.ExpDT.Date == dt.Date)).ToList();
                                }
                                else
                                {
                                    Cdata = fm.ContList.Where(x => (x.MaterialCode == cf.CFMList[i].MaterialCode) && (x.ExpDT.Date.AddDays(1) == dt.Date)).ToList();
                                }

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (Cdata.Count == 0))
                                {
                                    Cdata = fm.ContList.Where(x => (x.MaterialCode == cf.CFMList[i].MaterialCode) && (((DateTime)x.AlterDate).Date == dt.Date)).ToList();
                                }

                                foreach (var cd in Cdata)
                                {
                                    T2 = T2 + cd.Quantity;
                                }
                                Value = Value + T2;

                                //ask in progress
                                List<Raws> tData1 = new List<Raws>();

                                if (cf.cbAddDay.Checked == false)
                                {
                                    tData1 = cf.CFMList[i].TaskNavList1.Where(x => x.BBFDate == dt).ToList();
                                }
                                else
                                {
                                    tData1 = cf.CFMList[i].TaskNavList1.Where(x => x.AddDate == dt).ToList();
                                }
                                double T1 = 0;

                                if ((dt.Day == 1) && (FDTList.Count == 1) && (tData1.Count == 0))
                                {
                                    tData1 = cf.CFMList[i].TaskNavList1.Where(x => x.AlterDate == dt).ToList();
                                }

                                foreach (var td1 in tData1)
                                {
                                    T1 = T1 + td1.Quantity;
                                }
                                Value = Value + T1;

                                //credit
                                var dll = cf.XLDataList.Where(x => (x.ProdCode == cf.CFMList[i].MaterialCode) && (x.DateWork.Date == dt)).ToList();
                                foreach (var dl in dll)
                                {
                                    Value = Value - dl.Quantity;
                                }

                                //future planning Asks
                                double Q = 0;
                                var adata = cf.AppList;    // fm.tdb.tApplication.ToList();
                                                           //  if (adata.Count > 0)
                                {
                                    adata = adata.Where(x => x.MaterialCode == cf.CFMList[i].MaterialCode).ToList();
                                    //    if (adata.Count > 0)
                                    {
                                        adata = adata.Where(x => x.DateWork.Date == dt.Date).ToList();

                                        if ((adata.Count == 0) && (FDTList.Count == 1) && (dt.Day == 1))
                                        {
                                            adata = cf.AppList;
                                            adata = adata.Where(x => x.MaterialCode == cf.CFMList[i].MaterialCode).ToList();
                                            adata = adata.Where(x => (DateTime)x.AlterDate == dt.Date).ToList();
                                        }

                                        foreach (var ad in adata)
                                        {
                                            Q = Q + ad.Quantity;
                                        }

                                        cf.USCList.Add(new UserSelectedCell
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
                                    //  if (dt.Subtract(DateTime.Today.Date).Days < 60)
                                    {
                                        //  CFMBool[i] = false;
                                        if (!cf.CFMListGold.Contains(cf.CFMList[i].MaterialCode))  //-
                                        {
                                            cf.CFMListGold.Add(cf.CFMList[i].MaterialCode);
                                            cf.CFMListRed.Add(cf.CFMList[i].MaterialCode);
                                            Colorres Col = new Colorres
                                            {
                                                ProdCode = cf.CFMList[i].MaterialCode,
                                                ColNo = new List<int>()
                                            };
                                            Col.ColNo.Add(Cols);
                                            cf.CFMListMinus.Add(Col);
                                        }
                                        else
                                        {
                                            var Coll = cf.CFMListMinus.Where(x => x.ProdCode == cf.CFMList[i].MaterialCode).ToList();
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
                                                    ProdCode = cf.CFMList[i].MaterialCode,
                                                    ColNo = new List<int>()
                                                };
                                                Col.ColNo.Add(Cols);
                                                cf.CFMListMinus.Add(Col);
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

                        table.Rows.Add(NewString);
                    }
                    catch (Exception xx)
                    { }
                }
            }            
            table.Locale = System.Globalization.CultureInfo.InvariantCulture;
            //    adapter.Fill(table);

            GC.Collect();
         //   GC.WaitForPendingFinalizers();
         //   GC.Collect();

            return table;
        }
    }



    public class Cache
    {
        private static int RowsPerPage;
        private static int ColumnIndex;
        private static int RowsCount;
        private static int PagesCount;

        // Represents one page of data.
        public struct DataPage
        {
            public DataTable table;
            private int lowestIndexValue;
            private int highestIndexValue;
           // private int recordCount;
          

            public DataPage(DataTable table, int rowIndex)
            {
                this.table = table;
                lowestIndexValue = MapToLowerBoundary(rowIndex);
                highestIndexValue = MapToUpperBoundary(rowIndex);
              //  recordCount = table.Rows.Count;
                System.Diagnostics.Debug.Assert(lowestIndexValue >= 0);
                System.Diagnostics.Debug.Assert(highestIndexValue >= 0);
            }

            public int LowestIndex
            {
                get
                {
                    return lowestIndexValue;
                }
            }

            public int HighestIndex
            {
                get
                {
                    return highestIndexValue;
                }
            }

            public static int MapToLowerBoundary(int rowIndex)
            {
                // Return the lowest index of a page containing the given index.
                return (rowIndex / RowsPerPage) * RowsPerPage;
            }

            private static int MapToUpperBoundary(int rowIndex)
            {
                // Return the highest index of a page containing the given index.
                return MapToLowerBoundary(rowIndex) + RowsPerPage - 1;
            }
        }

        private DataPage[] cachePages;
        private IDataPageRetriever dataSupply;

        public int GetLowIndex()
        {
            try
            {
                return cachePages[0].LowestIndex / RowsPerPage;
            }
            catch (Exception xx)
            {
                return 0;
            }
        }

        public int GetHihgIndex()
        {
            try
            {
                return cachePages[cachePages.Length - 1].HighestIndex / RowsPerPage;
            }
            catch (Exception xx)
            {
                return 0;
            }
        }

        public Cache(IDataPageRetriever dataSupplier, int rowsPerPage)
        {
            dataSupply = dataSupplier;
            Cache.RowsPerPage = rowsPerPage;
            Cache.RowsCount = dataSupply.GetRowCount();
            if (RowsCount % RowsPerPage==0)
            {
                Cache.PagesCount = RowsCount / RowsPerPage;
            }
            else
            {
                Cache.PagesCount = RowsCount / RowsPerPage + 1;
            }
              LoadFirstTwoPages();
            //LoadPages(Cache.PagesCount);
        }

        // Sets the value of the element parameter if the value is in the cache.
        private bool IfPageCached_ThenSetElement(int rowIndex, int columnIndex, ref string element)
        {
            if (IsRowCachedInPage(0, rowIndex))
            {
                element = cachePages[0].table
                    .Rows[rowIndex % RowsPerPage][columnIndex].ToString();
                return true;
            }
            else if (IsRowCachedInPage(1, rowIndex))
            {
                element = cachePages[1].table
                    .Rows[rowIndex % RowsPerPage][columnIndex].ToString();
                return true;
            }

            return false;
        }

    //    public int LowerIndex()

        public string RetrieveElement(int rowIndex, int columnIndex)
        {
            string element = null;

            if (IfPageCached_ThenSetElement(rowIndex, columnIndex, ref element))
            {
                return element;
            }
            else
            {
                return RetrieveData_CacheIt_ThenReturnElement(
                    rowIndex, columnIndex);
            }
        }

        private void LoadPages(int PageCount)
        {
            cachePages = new DataPage[] { };

            for (int i=0; i<PageCount; i++)
            {
                DataPage dp = new DataPage(dataSupply.SupplyPageOfData(DataPage.MapToLowerBoundary(RowsPerPage * i), RowsPerPage, 0), RowsPerPage * i);
                Array.Resize(ref cachePages, cachePages.Length + 1);
                cachePages[cachePages.Length - 1] = dp;
            }
        }

        private void LoadFirstTwoPages()
        {
            cachePages = new DataPage[]{
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(0), RowsPerPage, 0), 0),
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(RowsPerPage),  RowsPerPage, 0), RowsPerPage),
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(2*RowsPerPage),  RowsPerPage, 0), 2*RowsPerPage),
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(3*RowsPerPage),  RowsPerPage, 0),3* RowsPerPage),
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(4*RowsPerPage),  RowsPerPage, 0), 4*RowsPerPage),
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(5*RowsPerPage),  RowsPerPage, 0), 5*RowsPerPage),
            new DataPage(dataSupply.SupplyPageOfData(  DataPage.MapToLowerBoundary(6*RowsPerPage),  RowsPerPage, 0), 6*RowsPerPage)
            };

            GC.Collect();
        }

        private string RetrieveData_CacheIt_ThenReturnElement(
            int rowIndex, int columnIndex)
        {
            // Retrieve a page worth of data containing the requested value.
            DataTable table = dataSupply.SupplyPageOfData(
                DataPage.MapToLowerBoundary(rowIndex), RowsPerPage, ColumnIndex);

            // Replace the cached page furthest from the requested cell
            // with a new page containing the newly retrieved data.
            cachePages[GetIndexToUnusedPage(rowIndex)] = new DataPage(table, rowIndex);

            return RetrieveElement(rowIndex, columnIndex);
        }

        // Returns the index of the cached page most distant from the given index
        // and therefore least likely to be reused.
        private int GetIndexToUnusedPage(int rowIndex)
        {
            if (rowIndex > cachePages[0].HighestIndex &&
                rowIndex > cachePages[1].HighestIndex)
            {
                int offsetFromPage0 = rowIndex - cachePages[0].HighestIndex;
                int offsetFromPage1 = rowIndex - cachePages[1].HighestIndex;
                if (offsetFromPage0 < offsetFromPage1)
                {
                    return 1;
                }
                return 0;
            }
            else
            {
                int offsetFromPage0 = cachePages[0].LowestIndex - rowIndex;
                int offsetFromPage1 = cachePages[1].LowestIndex - rowIndex;
                if (offsetFromPage0 < offsetFromPage1)
                {
                    return 1;
                }
                return 0;
            }
        }

        // Returns a value indicating whether the given row index is contained
        // in the given DataPage.
        private bool IsRowCachedInPage(int pageNumber, int rowIndex)
        {
            return rowIndex <= cachePages[pageNumber].HighestIndex &&
                rowIndex >= cachePages[pageNumber].LowestIndex;
        }
    }
}
