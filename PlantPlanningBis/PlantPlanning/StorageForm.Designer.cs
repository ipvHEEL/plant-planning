namespace PlantPlanning
{
    partial class StorageForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(StorageForm));
            this.button4 = new System.Windows.Forms.Button();
            this.tbLotCode = new System.Windows.Forms.TextBox();
            this.tbProductName = new System.Windows.Forms.TextBox();
            this.tbProductCode = new System.Windows.Forms.TextBox();
            this.difcurdates = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.difdates = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.bbf = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dpr = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.BBFDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ProdDate = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TestOnStorage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.TestQualityGroup = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.tbLotName = new System.Windows.Forms.TextBox();
            this.StorageType = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Leader = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Quantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.LotDescription = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MaterialLot = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.SIG = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ProductName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.ProductCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.Storage = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.TestQuality = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // button4
            // 
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.Location = new System.Drawing.Point(1363, 3);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(90, 40);
            this.button4.TabIndex = 17;
            this.button4.Text = "Сброс";
            this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // tbLotCode
            // 
            this.tbLotCode.Location = new System.Drawing.Point(824, 12);
            this.tbLotCode.Name = "tbLotCode";
            this.tbLotCode.Size = new System.Drawing.Size(141, 20);
            this.tbLotCode.TabIndex = 15;
            this.tbLotCode.TextChanged += new System.EventHandler(this.tbProductCode_TextChanged);
            // 
            // tbProductName
            // 
            this.tbProductName.Location = new System.Drawing.Point(378, 12);
            this.tbProductName.Name = "tbProductName";
            this.tbProductName.Size = new System.Drawing.Size(341, 20);
            this.tbProductName.TabIndex = 14;
            this.tbProductName.TextChanged += new System.EventHandler(this.tbProductCode_TextChanged);
            // 
            // tbProductCode
            // 
            this.tbProductCode.Location = new System.Drawing.Point(257, 12);
            this.tbProductCode.Name = "tbProductCode";
            this.tbProductCode.Size = new System.Drawing.Size(115, 20);
            this.tbProductCode.TabIndex = 13;
            this.tbProductCode.TextChanged += new System.EventHandler(this.tbProductCode_TextChanged);
            // 
            // difcurdates
            // 
            this.difcurdates.HeaderText = "difcurdates";
            this.difcurdates.Name = "difcurdates";
            this.difcurdates.ReadOnly = true;
            // 
            // difdates
            // 
            this.difdates.HeaderText = "difdates";
            this.difdates.Name = "difdates";
            this.difdates.ReadOnly = true;
            // 
            // bbf
            // 
            this.bbf.HeaderText = "bbf";
            this.bbf.Name = "bbf";
            this.bbf.ReadOnly = true;
            // 
            // dpr
            // 
            this.dpr.HeaderText = "dpr";
            this.dpr.Name = "dpr";
            this.dpr.ReadOnly = true;
            // 
            // BBFDate
            // 
            this.BBFDate.HeaderText = "Срок годности";
            this.BBFDate.Name = "BBFDate";
            this.BBFDate.ReadOnly = true;
            // 
            // ProdDate
            // 
            this.ProdDate.HeaderText = "Дата производства";
            this.ProdDate.Name = "ProdDate";
            this.ProdDate.ReadOnly = true;
            // 
            // TestOnStorage
            // 
            this.TestOnStorage.HeaderText = "Тест на складе";
            this.TestOnStorage.Name = "TestOnStorage";
            this.TestOnStorage.ReadOnly = true;
            // 
            // TestQualityGroup
            // 
            this.TestQualityGroup.HeaderText = "Тест качества группа";
            this.TestQualityGroup.Name = "TestQualityGroup";
            this.TestQualityGroup.ReadOnly = true;
            // 
            // tbLotName
            // 
            this.tbLotName.Location = new System.Drawing.Point(971, 12);
            this.tbLotName.Name = "tbLotName";
            this.tbLotName.Size = new System.Drawing.Size(350, 20);
            this.tbLotName.TabIndex = 16;
            this.tbLotName.TextChanged += new System.EventHandler(this.tbProductCode_TextChanged);
            // 
            // StorageType
            // 
            this.StorageType.HeaderText = "Тип склада";
            this.StorageType.Name = "StorageType";
            this.StorageType.ReadOnly = true;
            // 
            // Leader
            // 
            this.Leader.HeaderText = "Владелец";
            this.Leader.Name = "Leader";
            this.Leader.ReadOnly = true;
            this.Leader.Width = 150;
            // 
            // Quantity
            // 
            this.Quantity.HeaderText = "Количество";
            this.Quantity.Name = "Quantity";
            this.Quantity.ReadOnly = true;
            // 
            // LotDescription
            // 
            this.LotDescription.HeaderText = "Описание лота";
            this.LotDescription.Name = "LotDescription";
            this.LotDescription.ReadOnly = true;
            this.LotDescription.Width = 350;
            // 
            // MaterialLot
            // 
            this.MaterialLot.HeaderText = "Лот";
            this.MaterialLot.Name = "MaterialLot";
            this.MaterialLot.ReadOnly = true;
            this.MaterialLot.Width = 150;
            // 
            // SIG
            // 
            this.SIG.HeaderText = "Мат. группа";
            this.SIG.Name = "SIG";
            this.SIG.ReadOnly = true;
            // 
            // ProductName
            // 
            this.ProductName.HeaderText = "Наименование";
            this.ProductName.Name = "ProductName";
            this.ProductName.ReadOnly = true;
            this.ProductName.Width = 350;
            // 
            // ProductCode
            // 
            this.ProductCode.HeaderText = "Код продукта";
            this.ProductCode.Name = "ProductCode";
            this.ProductCode.ReadOnly = true;
            this.ProductCode.Width = 120;
            // 
            // Storage
            // 
            this.Storage.HeaderText = "Склад";
            this.Storage.Name = "Storage";
            this.Storage.ReadOnly = true;
            this.Storage.Width = 200;
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Storage,
            this.ProductCode,
            this.ProductName,
            this.SIG,
            this.MaterialLot,
            this.LotDescription,
            this.Quantity,
            this.Leader,
            this.TestQuality,
            this.StorageType,
            this.TestQualityGroup,
            this.TestOnStorage,
            this.ProdDate,
            this.BBFDate,
            this.dpr,
            this.bbf,
            this.difdates,
            this.difcurdates});
            this.dataGridView1.Location = new System.Drawing.Point(11, 47);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.ReadOnly = true;
            this.dataGridView1.Size = new System.Drawing.Size(1726, 637);
            this.dataGridView1.TabIndex = 12;
            this.dataGridView1.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridView1_RowPrePaint);
            // 
            // TestQuality
            // 
            this.TestQuality.HeaderText = "Тест качества";
            this.TestQuality.Name = "TestQuality";
            this.TestQuality.ReadOnly = true;
            // 
            // button3
            // 
            this.button3.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Right)));
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(1582, 690);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(155, 53);
            this.button3.TabIndex = 11;
            this.button3.Text = "          Выход";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Anchor = System.Windows.Forms.AnchorStyles.Bottom;
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(789, 690);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(155, 53);
            this.button2.TabIndex = 10;
            this.button2.Text = "Экспорт в Excel";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(12, 690);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(155, 53);
            this.button1.TabIndex = 9;
            this.button1.Text = "           Обновить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // StorageForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1749, 746);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.tbLotCode);
            this.Controls.Add(this.tbProductName);
            this.Controls.Add(this.tbProductCode);
            this.Controls.Add(this.tbLotName);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "StorageForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Текущие остатки";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.StorageForm_FormClosed);
            this.Resize += new System.EventHandler(this.StorageForm_Resize);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.TextBox tbLotCode;
        private System.Windows.Forms.TextBox tbProductName;
        private System.Windows.Forms.TextBox tbProductCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn difcurdates;
        private System.Windows.Forms.DataGridViewTextBoxColumn difdates;
        private System.Windows.Forms.DataGridViewTextBoxColumn bbf;
        private System.Windows.Forms.DataGridViewTextBoxColumn dpr;
        private System.Windows.Forms.DataGridViewTextBoxColumn BBFDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn ProdDate;
        private System.Windows.Forms.DataGridViewTextBoxColumn TestOnStorage;
        private System.Windows.Forms.DataGridViewTextBoxColumn TestQualityGroup;
        private System.Windows.Forms.TextBox tbLotName;
        private System.Windows.Forms.DataGridViewTextBoxColumn StorageType;
        private System.Windows.Forms.DataGridViewTextBoxColumn Leader;
        private System.Windows.Forms.DataGridViewTextBoxColumn Quantity;
        private System.Windows.Forms.DataGridViewTextBoxColumn LotDescription;
        private System.Windows.Forms.DataGridViewTextBoxColumn MaterialLot;
        private System.Windows.Forms.DataGridViewTextBoxColumn SIG;
        private System.Windows.Forms.DataGridViewTextBoxColumn ProductName;
        private System.Windows.Forms.DataGridViewTextBoxColumn ProductCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn Storage;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridViewTextBoxColumn TestQuality;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
    }
}