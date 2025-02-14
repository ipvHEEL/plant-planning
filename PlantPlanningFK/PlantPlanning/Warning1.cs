using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace PlantPlanning
{
    public partial class Warning1 : Form
    {
        FormMain fm;
        string Rres = "";

        public Warning1(FormMain FM)
        {
            InitializeComponent();
            fm = FM;
            textBox1.Clear();
            LoadData2();
          //  textBox1.Text = fm.ErrorMainStr;
        }

        private void LoadData2()
        {
            string[] NewString;

            if (fm.ErrorMainStr.Length!=0)
            {
                NewString = fm.ErrorMainStr.Split(fm.RowSplitter);
                Rres = NewString + System.Environment.NewLine;
                dataGridView1.Rows.Add(NewString);

                foreach (var E in fm.MainEGrid.OrderBy(x=>x.DateWork).ToList())
                {
                    NewString = new string[] { E.DateWork.ToString("dd.MM.yyyy"),
                                                E.ProdCode,
                                                E.ProdName,
                                                E.LineName,
                                                E.Comment};
                    dataGridView1.Rows.Add(NewString);
                    Rres = Rres + NewString + System.Environment.NewLine;
                }

                NewString = new string[] { "", "", "" };
                dataGridView1.Rows.Add(NewString);
                NewString = new string[] { "", "", "" };
                dataGridView1.Rows.Add(NewString);
            }

            if (fm.ErrorStr.Length != 0)
            {
                NewString = fm.ErrorStr.Split(fm.RowSplitter);
                dataGridView1.Rows.Add(NewString);

                Rres=Rres+ NewString + System.Environment.NewLine;

                foreach (var E in fm.CommonEGrid.OrderBy(x => x.DateWork).ToList())
                {
                    NewString = new string[] { E.DateWork.ToString("dd.MM.yyyy"),
                                                E.ProdCode,
                                                E.ProdName,
                                                E.LineName,
                                                E.Comment};
                    dataGridView1.Rows.Add(NewString);

                    Rres = Rres + NewString + System.Environment.NewLine;
                }
            }
        }


        private void LoadData()
        {
            /*string[] NewString;
            string s = fm.ErrorMainStr;
            string s1 = "";
            Int32 Ind = 0;

            while (s.Length>0)
            {
                Ind = s.IndexOf(System.Environment.NewLine);
                    if (Ind > 0)
                {
                    s1 = s.Substring(0, Ind);
                    s = s.Substring(Ind + 2, s.Length - Ind - 2);
                }
                    else
                {
                    s1 = s;
                    s = "";
                }

                s.Replace(System.Environment.NewLine, "");
                NewString=s1.Split(fm.RowSplitter);
                dataGridView1.Rows.Add(NewString);         
            }

            for (int i = 0; i < dataGridView1.Columns.Count; i++)
            {
                // c.SortMode = DataGridViewColumnSortMode.NotSortable;
                dataGridView1.Columns[i].SortMode = DataGridViewColumnSortMode.NotSortable;
            }
            */
        }


        private void button2_Click(object sender, EventArgs e)
        {
            this.Close();
        }

        private void Warning1_FormClosed(object sender, FormClosedEventArgs e)
        {
            fm.Enabled = true;
        }

        private void button1_Click(object sender, EventArgs e)
        {
            // string res = fm.ErrorMainStr + System.Environment.NewLine + fm.ErrorStr;

            dataGridView1.SelectAll();
            dataGridView1.ClipboardCopyMode = DataGridViewClipboardCopyMode.EnableAlwaysIncludeHeaderText;

            DataObject data = dataGridView1.GetClipboardContent();
            if (data != null)
            {
                Clipboard.SetDataObject(data);
            }
            else
            {
                Clipboard.SetText(Rres.Replace(fm.RowSplitter, ' '));
            }
          //fm.ErrorMainStr.Replace(fm.RowSplitter,' '));  //    Clipboard.SetText( textBox1.Text )
            this.Close();
        }
    }
}
