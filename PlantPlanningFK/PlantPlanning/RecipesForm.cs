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

namespace PlantPlanning
{
    public partial class RecipesForm : Form
    {
        FormMain fm;
        BindingSource bs;
        DataTable dt1;
        List<fn_select_RPSettings_Result> RP;
        List<ProductRecipes> PRList;
        List<Recipe> RecList;

        public RecipesForm(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            CreateRecipeList();
            LoadData();         
        }

        private void CreateRecipeList()
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

            Cursor = Cursors.WaitCursor;
            PRList = new List<ProductRecipes>();
            RecList = new List<Recipe>();

            for (int i=0; i<fm.ProdRecipeList.Count; i++)
            {
                pr = new ProductRecipes
                {
                    ProductCode = fm.ProdRecipeList[i].ProductCode,
                    ProductCodeString = fm.ProdRecipeList[i].ProductCodeString,
                    ProductName= fm.ProdRecipeList[i].ProductName,
                    RecLineList = new List<ProductLinesRecipes>()
                };

                for (int j=0; j < fm.ProdRecipeList[i].RecLineList.Count; j++)
                {
                    plr = new ProductLinesRecipes
                    {
                        LineNumber = fm.ProdRecipeList[i].RecLineList[j].LineNumber,
                        RecList = new List<Recipe>()
                    };

                    for (int k=0; k< fm.ProdRecipeList[i].RecLineList[j].RecList.Count; k++)
                    {
                        rr = new Recipe
                        {
                            WorkDateStart = fm.ProdRecipeList[i].RecLineList[j].RecList[k].WorkDateStart,
                            WorkDateEnd = fm.ProdRecipeList[i].RecLineList[j].RecList[k].WorkDateEnd,
                            rList = new List<RecipeList>()
                        };

                        for (int l=0; l< fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList.Count; l++)
                        {
                            rl = new RecipeList
                            {
                                MaterialCode = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].MaterialCode,
                                MaterialName = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].MaterialName,
                                MaterialQuantity = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].MaterialQuantity,
                                IPG = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].IPG
                            };
                            rr.rList.Add(rl);
                        }

                        plr.RecList.Add(rr);
                    }

                    pr.RecLineList.Add(plr);
                }

                PRList.Add(pr);

                pr = new ProductRecipes
                {
                    ProductCode = fm.ProdRecipeList[i].ProductCode,
                    ProductCodeString = fm.ProdRecipeList[i].ProductCodeString.Trim() + "_П/УП",
                    ProductName = fm.ProdRecipeList[i].ProductName+" (Переупаковка)",
                    RePack = true,
                    RecLineList = new List<ProductLinesRecipes>()
                };

                for (int j = 0; j < fm.ProdRecipeList[i].RecLineList.Count; j++)
                {
                    plr = new ProductLinesRecipes
                    {
                        LineNumber = fm.ProdRecipeList[i].RecLineList[j].LineNumber,
                        RecList = new List<Recipe>()
                    };

                    for (int k = 0; k < fm.ProdRecipeList[i].RecLineList[j].RecList.Count; k++)
                    {
                        rr = new Recipe
                        {
                            WorkDateStart = fm.ProdRecipeList[i].RecLineList[j].RecList[k].WorkDateStart,
                            WorkDateEnd = fm.ProdRecipeList[i].RecLineList[j].RecList[k].WorkDateEnd,
                            rList = new List<RecipeList>()
                        };

                        for (int l = 0; l < fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList.Count; l++)
                        {
                            if (fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].IPG.ToLower() == "тара")
                            {
                                rl = new RecipeList
                                {
                                    MaterialCode = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].MaterialCode,
                                    MaterialName = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].MaterialName,
                                    MaterialQuantity = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].MaterialQuantity,
                                    IPG = fm.ProdRecipeList[i].RecLineList[j].RecList[k].rList[l].IPG
                                };
                                rr.rList.Add(rl);
                            }
                        }

                        plr.RecList.Add(rr);
                    }

                    pr.RecLineList.Add(plr);
                }

                PRList.Add(pr);
            }

            GC.Collect();
            Cursor = Cursors.Default;
        }

        private void LoadData()
        {
            dataGridViewMain.DataSource = null;
            dataGridViewMain.Columns.Clear();

            dataGridViewSpec.Rows.Clear();
            dataGridViewDates.Rows.Clear();
            dataGridViewDetails.Rows.Clear();

            comboBoxLines.Items.Clear();
            dateTimePicker1.Value = DateTime.Today.Date.AddDays(1);
            textBoxPlan.Text = "";
            textBoxQuantity.Text = "0";
        //    comboBoxLines.Items.Clear();

            textBoxCode.Text = "";
            textBoxNaim.Text = "";

            string[] NewString;
            char RowSplitter = '|';
            string sVal = "";

          //  var data = fm.eqp.tMonitorLineName.ToList();

            foreach (var d in fm.LineList)
            {
                comboBoxLines.Items.Add(d.Trim());
            }

            if (comboBoxLines.Items.Count > 0)
            {
                comboBoxLines.Text = comboBoxLines.Items[0].ToString();
            }

          //  RP = new List<Context.fn_select_RPSettings_Result>();
          //  RP = fm.eqp.fn_select_RPSettings().ToList();

           /* foreach (var rp in RP)
            {
                if (rp.ProductionLine.Contains("Переупаковка"))
                {
                    rp.ProductionLine = rp.ProductionLine.Replace("Переупаковка", "").Trim();
                    rp.ProductCode = rp.ProductCode + "_П/УП";
                }    
            }*/

            var RPBase = PRList.Select(x => new { x.ProductCodeString, x.ProductName }).ToList();

            RPBase = RPBase.Distinct().OrderBy(x=>x.ProductCodeString).ToList();

            /* List<lData> RPItog = RP.GroupBy(x => new { x.ProductCode, x.ProductName })
                  .OrderBy(x => x.Key)
                  .Select(g => new lData
                  {
                     // DateWork=DateTime.Today.Date,
                      ProdCode=g.Key.ProductCode.Replace("_П/УП", ""),
                      ProdCodeStr= g.Key.ProductCode,
                      ProdName=g.Key.ProductName,                   
                      RePack = g.Key.ProductCode.Contains("_П/УП") ? true : false
                  }).ToList();

             dataGridViewMain.DataSource = RPItog;*/

            dt1 = new DataTable();
            dt1.Columns.Clear();
            dt1.Rows.Clear();

            dt1.Columns.Add("ProductCode");
            dt1.Columns.Add("ProductName");

            foreach (var rp in RPBase)
            {
                sVal = rp.ProductCodeString + RowSplitter + rp.ProductName;
                NewString = sVal.Split(RowSplitter);
                dt1.Rows.Add(NewString);
            }

            bs = new BindingSource();
            bs.DataSource = dt1;

            dataGridViewMain.DataSource = bs;
            dataGridViewMain.Columns[0].Width = 150;
            dataGridViewMain.Columns[1].Width = this.Width - 200;
            dataGridViewMain.ReadOnly = true;
            dataGridViewMain.AllowUserToAddRows = false;

            GC.Collect();
        }

        private void RecipesForm_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();

            fm.Enabled = true;
        }

        private void textBoxCode_TextChanged(object sender, EventArgs e)
        {
         //   bool fl = false;
            bs.Filter = "";
            dataGridViewSpec.Rows.Clear();
            dataGridViewDates.Rows.Clear();
            dataGridViewDetails.Rows.Clear();

            if (textBoxCode.Text.Trim()!="")
            {
                bs.Filter = " ProductCode like '%" + textBoxCode.Text.Trim() + "%' ";
               // fl = true;
            }

            if (textBoxNaim.Text.Trim() != "")
            {
                if (bs.Filter.Length > 0)
                {
                    bs.Filter = bs.Filter + " and ProductName like '%" + textBoxNaim.Text.Trim() + "%' ";
                }
                else
                {
                    bs.Filter = " ProductName like '%" + textBoxNaim.Text.Trim() + "%' ";
                }
            } 
        }

        private void dataGridViewMain_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewSpec.Rows.Clear();
            dataGridViewDates.Rows.Clear();
            dataGridViewDetails.Rows.Clear();

            if (dataGridViewMain.RowCount > 0)
            {
                string Code = dataGridViewMain.CurrentRow.Cells[0].Value.ToString();
                //  string Line = "";

                if (!String.IsNullOrEmpty(Code))
                {
                    var data = PRList.Where(x => x.ProductCodeString == Code).ToList();

                    if (data.Count > 0)
                    {
                        List<string> LList = data[0].RecLineList.Select(x => x.LineNumber).Distinct().ToList();

                        foreach (var ll in LList)
                        {
                            if (Code.Contains("П/УП"))
                            {
                                dataGridViewSpec.Rows.Add(ll.Trim() + " - П/УП");
                            }
                            else
                            {
                                dataGridViewSpec.Rows.Add(ll.Trim());
                            }
                        }
                    }
                }
                GC.Collect();
            }
         /*  
                var data = RP.Where(x => x.ProductCode == Code).ToList();
                List<string> LList = data.Select(x => x.ProductionLine).Distinct().ToList();

                foreach (var ll in LList)
                {
                    Line = ll;

                    if (Line.Contains("Переупаковка"))
                    {
                        Line = Line.Replace("Переупаковка", "");                        
                    }

                    dataGridViewSpec.Rows.Add(Line.Trim());
                }
            }*/
        }

        private void dataGridViewSpec_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewDates.Rows.Clear();
            dataGridViewDetails.Rows.Clear();

            if (dataGridViewMain.RowCount > 0)
            {
                if (dataGridViewSpec.RowCount > 0)
                {
                    string Code = dataGridViewMain.CurrentRow.Cells[0].Value.ToString();
                    string Line = dataGridViewSpec.CurrentRow.Cells[0].Value.ToString();
                    Line = Line.Replace(" - П/УП", "").Trim();

                    if ((!String.IsNullOrEmpty(Line)) && (!String.IsNullOrEmpty(Code)))
                    {
                        var data = PRList.Where(x => x.ProductCodeString == Code).ToList();
                        if (data.Count > 0)
                        {
                            var lData = data[0].RecLineList.Where(x => x.LineNumber == Line).ToList();
                            if (lData.Count > 0)
                            {
                                RecList = lData[0].RecList;

                                foreach (var rl in RecList)
                                {
                                    dataGridViewDates.Rows.Add("c " + rl.WorkDateStart.ToString("dd.MM.yyyy") + " по " + rl.WorkDateEnd.ToString("dd.MM.yyyy"));
                                }
                            }
                        }
                    }
                }
            }
            GC.Collect();
        }

        private void dataGridViewDates_SelectionChanged(object sender, EventArgs e)
        {
            dataGridViewDetails.Rows.Clear();
            Int32 Ind = 0;

            if (dataGridViewMain.RowCount > 0)
            {
                if (dataGridViewSpec.RowCount > 0)
                {
                    if (dataGridViewDates.RowCount > 0)
                    {
                        string Code = dataGridViewMain.CurrentRow.Cells[0].Value.ToString();
                        string Line = dataGridViewSpec.CurrentRow.Cells[0].Value.ToString();
                        Line = Line.Replace(" - П/УП", "").Trim();

                        if ((!String.IsNullOrEmpty(Line)) && (!String.IsNullOrEmpty(Code)))
                        {
                            var data = PRList.Where(x => x.ProductCodeString == Code).ToList();
                            if (data.Count > 0)
                            {
                                var lData = data[0].RecLineList.Where(x => x.LineNumber == Line).ToList();
                                if (lData.Count > 0)
                                {
                                    Ind = dataGridViewDates.CurrentRow.Index;

                                    Recipe rl = RecList[Ind];

                                    List<RecipeList> rList = rl.rList;

                                    foreach (var r in rList)
                                    {
                                        string[] NS = new string[] { r.MaterialCode,
                                                                    r.MaterialName,
                                                                    r.IPG,
                                                                    (r.MaterialQuantity*1000).ToString("N3")};
                                        dataGridViewDetails.Rows.Add(NS);
                                    }
                                }
                            }
                        }
                    }
                }
            }
            GC.Collect();
        }

        private void button4_Click(object sender, EventArgs e)
        {
            double Quantity = 0;

            try
            {
                Quantity = fm.ConvertStringToDouble(textBoxQuantity.Text.Trim());
            }
            catch (Exception xx)
            { }

            if (Quantity>0)
            {
                lData ld = new lData
                {
                    Quantity = Quantity,
                    Line = comboBoxLines.Text,
                    DateWork = dateTimePicker1.Value.Date,
                    ProdCode = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Replace(" - П/УП", "").Trim(),
                    ProdName= dataGridViewMain.CurrentRow.Cells[1].Value.ToString(),
                    ProdCodeStr = dataGridViewMain.CurrentRow.Cells[0].Value.ToString(),
                    RePack = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Contains("П/УП") ? true : false,
                    AlterProdName = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Replace(" - П/УП", "").Trim()
                };

                Cursor = Cursors.WaitCursor;
                fm.ExcelDataList.Add(ld);
                fm.GreateOperList();
                Cursor = Cursors.Default;
                MessageBox.Show("Данные успешно добавлены!", "Сообщение системы");
            }
            else
            {
                MessageBox.Show("Не указано количество!", "Сообщение системы");
            }
        }

        private void button1_Click(object sender, EventArgs e)
        {
            double Quantity = 0;

            try
            {
                Quantity = fm.ConvertStringToDouble(textBoxQuantity.Text.Trim());
            }
            catch (Exception xx)
            { }

            if (Quantity > 0)
            {
                lData ld = new lData
                {
                    Quantity = Quantity,
                    Line = comboBoxLines.Text,
                    DateWork = dateTimePicker1.Value.Date,
                    ProdName = dataGridViewMain.CurrentRow.Cells[1].Value.ToString(),
                    ProdCode = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Replace(" - П/УП", "").Trim(),
                    ProdCodeStr = dataGridViewMain.CurrentRow.Cells[0].Value.ToString(),
                    RePack = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Contains("П/УП") ? true : false,
                    AlterProdName = dataGridViewMain.CurrentRow.Cells[0].Value.ToString().Replace(" - П/УП", "").Trim()
                };

                Cursor = Cursors.WaitCursor;
                fm.ExcelDataList.Add(ld);
                fm.GreateOperList();
                Cursor = Cursors.Default;
                this.Close();
            }
            else
            {
                MessageBox.Show("Не указано количество!", "Сообщение системы");
            }
        }
    }
}
