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
    public partial class FormTask : Form
    {
        FormMain fm;
        ConsurmptionForm cf;

        SortableBindingList<TaskViewData> SortData = new SortableBindingList<TaskViewData>();
        List<TaskViewData> TVDList = new List<TaskViewData>();
        List<TaskViewData> TVDListBis = new List<TaskViewData>();

        public FormTask(FormMain FM, ConsurmptionForm CF)
        {
            InitializeComponent();

            fm = FM;
            cf = CF;

            dataGridView1.ReadOnly = true;
            dataGridView1.AllowUserToAddRows = false;

            LoadData();
        }

        private void LoadData()
        {
            Cursor = Cursors.WaitCursor;
            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            var Data = fm.tdb.GetNavPurchData().ToList();
            DateTime dt = DateTime.Today.Date;
            TVDList = new List<TaskViewData>();
            SortData = new SortableBindingList<TaskViewData>();

            if (Data.Count > 0)
            {
                Data = Data.Where(x => x.Datework.Date == dt).ToList();

                if (Data.Count > 0)
                {
                    dt = Data.Max(x => x.Datework);
                    dt = dt.AddMinutes(-30);

                    Data = Data.Where(x => x.Datework > dt).ToList();
                }
            }

            //comboBoxes

            List<string> Exec = Data.Select(x => x.Executor).Distinct().ToList();
         //   Exec = Exec.Distinct().ToList();
            Exec.Sort();
            cbExecutor.Items.Clear();

            for (int i = 0; i<Exec.Count; i++)
            {
                cbExecutor.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.Status).Distinct().ToList();
            Exec.Sort();
            cbStatus.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbStatus.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.NextStatus).Distinct().ToList();
            Exec.Sort();
            cbNextStatus.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbNextStatus.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.OrderNymber).Distinct().ToList();
            Exec.Sort();
            cbOrderNumber.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbOrderNumber.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.MaterialCode).Distinct().ToList();
            Exec.Sort();
            cbMaterialCode.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbMaterialCode.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.MaterialName).Distinct().ToList();
            Exec.Sort();
            cbMaterialName.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbMaterialName.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.Initiator).Distinct().ToList();
            Exec.Sort();
            cbInitiator.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbInitiator.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.MZP).Distinct().ToList();
            Exec.Sort();
            cbMZP.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbMZP.Items.Add(Exec[i]);
            }

            Exec = Data.Select(x => x.StatusMZP).Distinct().ToList();
            Exec.Sort();
            cbStatusMZP.Items.Clear();

            for (int i = 0; i < Exec.Count; i++)
            {
                cbStatusMZP.Items.Add(Exec[i]);
            }

            for (int i = 0; i<Data.Count; i++)
            {
                TaskViewData tvd = new TaskViewData();

                tvd.Address = Data[i].Address;
                tvd.ContrAgent = Data[i].ContrAgent;
                tvd.ContrAgentCode = Data[i].ContrAgentCode;
                tvd.CurrencyCode = Data[i].CurrencyCode;
                tvd.CurrPrice = Data[i].CurrPrice;
                tvd.CurrStatus = Data[i].CurrStatus;
                tvd.CurrStatusCode = Data[i].CurrStatusCode;
                tvd.DateCreateMZP = Data[i].DateCreateMZP;
                tvd.DateCreateZZ = Data[i].DateCreateZZ;
                tvd.DayToSend = Data[i].DayToSend;
                tvd.EdIzm = Data[i].EdIzm;
                tvd.ENSGroup = Data[i].ENSGroup;
                tvd.ENSGroupCode = Data[i].ENSGroupCode;
                tvd.Executor = Data[i].Executor;
                tvd.FactQuantity = Data[i].FactQuantity;
                tvd.FromBalanses = Data[i].FromBalanses;
                tvd.Initiator = Data[i].Initiator;
                tvd.LastStatusDate = Data[i].LastStatusDate;
                tvd.MaterialCode = Data[i].MaterialCode;
                tvd.MaterialName = Data[i].MaterialName;
                tvd.MZP = Data[i].MZP;
                tvd.MZPCode = Data[i].MZPCode;
                tvd.MZPEdIzm = Data[i].MZPEdIzm;
                tvd.MZPNaim = Data[i].MZPNaim;
                tvd.MZPQuantity = Data[i].MZPQuantity;
                tvd.NessDateOrder = Data[i].NessDateOrder;
                tvd.NextStatus = Data[i].NextStatus;
                tvd.NumberRowMZP = Data[i].NumberRowMZP;
                tvd.OrderNymber = Data[i].OrderNymber;
                tvd.PlanDate = Data[i].PlanDate;
                tvd.PlanOperDate = Data[i].PlanOperDate;
                tvd.PlanQuantity = Data[i].PlanQuantity;
                tvd.Quantity = Data[i].Quantity;
                tvd.QuantityKey = Data[i].QuantityKey;
                tvd.Status = Data[i].Status;
                tvd.StatusMZP = Data[i].StatusMZP;

                TVDList.Add(tvd);
            }

            ClearFilter();

            SortData = new SortableBindingList<TaskViewData>(TVDList);
            dataGridView1.DataSource = SortData;

            try
            {
                dataGridView1.Columns["Status"].Width = 200;
            }
            catch (Exception xx)
            { }

            try
            {
                dataGridView1.Columns["MaterialName"].Width = 350;
            }
            catch (Exception xx)
            { }

            Cursor = Cursors.Default;
            GC.Collect();
        }

        private void ClearFilter()
        {
            cbExecutor.Text = "";
            cbInitiator.Text = "";
            cbMaterialCode.Text = "";
            cbMaterialName.Text = "";
            cbMZP.Text = "";
            cbNextStatus.Text = "";
            cbOrderNumber.Text = "";
            cbStatus.Text = "";
            cbStatusMZP.Text = "";
        }

        private void FormTask_FormClosed(object sender, FormClosedEventArgs e)
        {
            GC.Collect();

            cf.Enabled = true;
        }

        private void button9_Click(object sender, EventArgs e)
        {
            LoadData();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            Cursor = Cursors.WaitCursor;

            TVDListBis = TVDList.ToList();

            dataGridView1.DataSource = null;
            dataGridView1.Rows.Clear();

            if (cbStatusMZP.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.StatusMZP.ToLower().Contains(cbStatusMZP.Text.ToLower())).ToList();
            }

            if (cbStatus.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.Status.ToLower().Contains(cbStatus.Text.ToLower())).ToList();
            }

            if (cbOrderNumber.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.OrderNymber.ToLower().Contains(cbOrderNumber.Text.ToLower())).ToList();
            }

            if (cbNextStatus.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.NextStatus.ToLower().Contains(cbNextStatus.Text.ToLower())).ToList();
            }

            if (cbMZP.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.MZP.ToLower().Contains(cbMZP.Text.ToLower())).ToList();
            }

            if (cbMaterialName.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.MaterialName.ToLower().Contains(cbMaterialName.Text.ToLower())).ToList();
            }

            if (cbMaterialCode.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.MaterialCode.ToLower().Contains(cbMaterialCode.Text.ToLower())).ToList();
            }

            if (cbInitiator.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.Initiator.ToLower().Contains(cbInitiator.Text.ToLower())).ToList();
            }

            if (cbExecutor.Text != "")
            {
                TVDListBis = TVDListBis.Where(x => x.Executor.ToLower().Contains(cbExecutor.Text.ToLower())).ToList();
            }

            SortData = new SortableBindingList<TaskViewData>(TVDListBis);
            dataGridView1.DataSource = SortData;

            try
            {
                dataGridView1.Columns["Status"].Width = 200;
            }
            catch (Exception xx)
            { }

            try
            {
                dataGridView1.Columns["MaterialName"].Width = 350;
            }
            catch (Exception xx)
            { }

            GC.Collect();
            Cursor = Cursors.Default;
        }
    }
}
