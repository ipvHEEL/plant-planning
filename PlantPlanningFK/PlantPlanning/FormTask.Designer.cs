namespace PlantPlanning
{
    partial class FormTask
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormTask));
            this.panel1 = new System.Windows.Forms.Panel();
            this.button9 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.label9 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label6 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.panel2 = new System.Windows.Forms.Panel();
            this.cbNextStatus = new System.Windows.Forms.ComboBox();
            this.cbStatus = new System.Windows.Forms.ComboBox();
            this.cbExecutor = new System.Windows.Forms.ComboBox();
            this.cbOrderNumber = new System.Windows.Forms.ComboBox();
            this.cbMaterialCode = new System.Windows.Forms.ComboBox();
            this.cbMaterialName = new System.Windows.Forms.ComboBox();
            this.cbInitiator = new System.Windows.Forms.ComboBox();
            this.cbMZP = new System.Windows.Forms.ComboBox();
            this.cbStatusMZP = new System.Windows.Forms.ComboBox();
            this.dataGridView1 = new PlantPlanning.DBDataGridView();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.cbStatusMZP);
            this.panel1.Controls.Add(this.cbMZP);
            this.panel1.Controls.Add(this.cbInitiator);
            this.panel1.Controls.Add(this.cbMaterialName);
            this.panel1.Controls.Add(this.cbMaterialCode);
            this.panel1.Controls.Add(this.cbOrderNumber);
            this.panel1.Controls.Add(this.cbExecutor);
            this.panel1.Controls.Add(this.cbStatus);
            this.panel1.Controls.Add(this.cbNextStatus);
            this.panel1.Controls.Add(this.button9);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.label9);
            this.panel1.Controls.Add(this.label8);
            this.panel1.Controls.Add(this.label7);
            this.panel1.Controls.Add(this.label6);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1834, 80);
            this.panel1.TabIndex = 0;
            // 
            // button9
            // 
            this.button9.Image = ((System.Drawing.Image)(resources.GetObject("button9.Image")));
            this.button9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button9.Location = new System.Drawing.Point(1272, 12);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(100, 53);
            this.button9.TabIndex = 19;
            this.button9.Text = "Очистить";
            this.button9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(1131, 12);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(105, 53);
            this.button1.TabIndex = 18;
            this.button1.Text = "Поиск";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(1596, 18);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(60, 13);
            this.label9.TabIndex = 16;
            this.label9.Text = "StatusMZP";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(1392, 18);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(41, 13);
            this.label8.TabIndex = 14;
            this.label8.Text = "Initiator";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(1494, 18);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(30, 13);
            this.label7.TabIndex = 12;
            this.label7.Text = "MZP";
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(746, 18);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(72, 13);
            this.label6.TabIndex = 10;
            this.label6.Text = "MaterialName";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(644, 18);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(69, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "MaterialCode";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(544, 18);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(69, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "OrderNymber";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(343, 18);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(59, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "NextStatus";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(143, 18);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Status";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(40, 18);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(49, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Executor";
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 80);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1834, 732);
            this.panel2.TabIndex = 1;
            // 
            // cbNextStatus
            // 
            this.cbNextStatus.FormattingEnabled = true;
            this.cbNextStatus.Location = new System.Drawing.Point(339, 34);
            this.cbNextStatus.Name = "cbNextStatus";
            this.cbNextStatus.Size = new System.Drawing.Size(100, 21);
            this.cbNextStatus.TabIndex = 20;
            // 
            // cbStatus
            // 
            this.cbStatus.FormattingEnabled = true;
            this.cbStatus.Location = new System.Drawing.Point(140, 34);
            this.cbStatus.Name = "cbStatus";
            this.cbStatus.Size = new System.Drawing.Size(197, 21);
            this.cbStatus.TabIndex = 21;
            // 
            // cbExecutor
            // 
            this.cbExecutor.FormattingEnabled = true;
            this.cbExecutor.Location = new System.Drawing.Point(38, 34);
            this.cbExecutor.Name = "cbExecutor";
            this.cbExecutor.Size = new System.Drawing.Size(100, 21);
            this.cbExecutor.TabIndex = 22;
            // 
            // cbOrderNumber
            // 
            this.cbOrderNumber.FormattingEnabled = true;
            this.cbOrderNumber.Location = new System.Drawing.Point(540, 34);
            this.cbOrderNumber.Name = "cbOrderNumber";
            this.cbOrderNumber.Size = new System.Drawing.Size(100, 21);
            this.cbOrderNumber.TabIndex = 23;
            // 
            // cbMaterialCode
            // 
            this.cbMaterialCode.FormattingEnabled = true;
            this.cbMaterialCode.Location = new System.Drawing.Point(642, 34);
            this.cbMaterialCode.Name = "cbMaterialCode";
            this.cbMaterialCode.Size = new System.Drawing.Size(100, 21);
            this.cbMaterialCode.TabIndex = 24;
            // 
            // cbMaterialName
            // 
            this.cbMaterialName.FormattingEnabled = true;
            this.cbMaterialName.Location = new System.Drawing.Point(744, 34);
            this.cbMaterialName.Name = "cbMaterialName";
            this.cbMaterialName.Size = new System.Drawing.Size(350, 21);
            this.cbMaterialName.TabIndex = 25;
            // 
            // cbInitiator
            // 
            this.cbInitiator.FormattingEnabled = true;
            this.cbInitiator.Location = new System.Drawing.Point(1388, 34);
            this.cbInitiator.Name = "cbInitiator";
            this.cbInitiator.Size = new System.Drawing.Size(100, 21);
            this.cbInitiator.TabIndex = 26;
            // 
            // cbMZP
            // 
            this.cbMZP.FormattingEnabled = true;
            this.cbMZP.Location = new System.Drawing.Point(1490, 34);
            this.cbMZP.Name = "cbMZP";
            this.cbMZP.Size = new System.Drawing.Size(100, 21);
            this.cbMZP.TabIndex = 27;
            // 
            // cbStatusMZP
            // 
            this.cbStatusMZP.FormattingEnabled = true;
            this.cbStatusMZP.Location = new System.Drawing.Point(1592, 34);
            this.cbStatusMZP.Name = "cbStatusMZP";
            this.cbStatusMZP.Size = new System.Drawing.Size(100, 21);
            this.cbStatusMZP.TabIndex = 28;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.DoubleBuffered = true;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1834, 732);
            this.dataGridView1.TabIndex = 0;
            // 
            // FormTask
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1834, 812);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormTask";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Существующие заявки";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormTask_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private  PlantPlanning.DBDataGridView dataGridView1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Label label9;
        public System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.ComboBox cbNextStatus;
        private System.Windows.Forms.ComboBox cbStatus;
        private System.Windows.Forms.ComboBox cbExecutor;
        private System.Windows.Forms.ComboBox cbOrderNumber;
        private System.Windows.Forms.ComboBox cbMaterialCode;
        private System.Windows.Forms.ComboBox cbMaterialName;
        private System.Windows.Forms.ComboBox cbInitiator;
        private System.Windows.Forms.ComboBox cbMZP;
        private System.Windows.Forms.ComboBox cbStatusMZP;
    }
}