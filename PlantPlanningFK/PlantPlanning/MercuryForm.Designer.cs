namespace PlantPlanning
{
    partial class MercuryForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MercuryForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.panel3 = new System.Windows.Forms.Panel();
            this.dataGridView1 = new PlantPlanning.DBDataGridView();
            this.label1 = new System.Windows.Forms.Label();
            this.textBoxProdCode = new System.Windows.Forms.TextBox();
            this.textBoxProdName = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBoxLotNo = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.textBoxMatCode = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.textBoxMatName = new System.Windows.Forms.TextBox();
            this.label5 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.panel1.SuspendLayout();
            this.panel3.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button9);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button5);
            this.panel1.Controls.Add(this.textBoxMatName);
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.textBoxMatCode);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.textBoxLotNo);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.textBoxProdName);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.textBoxProdCode);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Top;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1834, 94);
            this.panel1.TabIndex = 0;
            // 
            // panel2
            // 
            this.panel2.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel2.Location = new System.Drawing.Point(0, 793);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1834, 19);
            this.panel2.TabIndex = 1;
            // 
            // panel3
            // 
            this.panel3.Controls.Add(this.dataGridView1);
            this.panel3.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel3.Location = new System.Drawing.Point(0, 94);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(1834, 699);
            this.panel3.TabIndex = 2;
            // 
            // dataGridView1
            // 
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridView1.DoubleBuffered = true;
            this.dataGridView1.Location = new System.Drawing.Point(0, 0);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1834, 699);
            this.dataGridView1.TabIndex = 0;
            this.dataGridView1.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridView1_RowPrePaint);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(45, 21);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(75, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Код продукта";
            // 
            // textBoxProdCode
            // 
            this.textBoxProdCode.Location = new System.Drawing.Point(35, 46);
            this.textBoxProdCode.Name = "textBoxProdCode";
            this.textBoxProdCode.Size = new System.Drawing.Size(100, 20);
            this.textBoxProdCode.TabIndex = 1;
            this.textBoxProdCode.TextChanged += new System.EventHandler(this.textBoxProdCode_TextChanged);
            // 
            // textBoxProdName
            // 
            this.textBoxProdName.Location = new System.Drawing.Point(141, 46);
            this.textBoxProdName.Name = "textBoxProdName";
            this.textBoxProdName.Size = new System.Drawing.Size(290, 20);
            this.textBoxProdName.TabIndex = 3;
            this.textBoxProdName.TextChanged += new System.EventHandler(this.textBoxProdCode_TextChanged);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(151, 21);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(83, 13);
            this.label2.TabIndex = 2;
            this.label2.Text = "Наименование";
            // 
            // textBoxLotNo
            // 
            this.textBoxLotNo.Location = new System.Drawing.Point(437, 46);
            this.textBoxLotNo.Name = "textBoxLotNo";
            this.textBoxLotNo.Size = new System.Drawing.Size(100, 20);
            this.textBoxLotNo.TabIndex = 5;
            this.textBoxLotNo.TextChanged += new System.EventHandler(this.textBoxProdCode_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(447, 21);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(44, 13);
            this.label3.TabIndex = 4;
            this.label3.Text = "№ лота";
            // 
            // textBoxMatCode
            // 
            this.textBoxMatCode.Location = new System.Drawing.Point(843, 46);
            this.textBoxMatCode.Name = "textBoxMatCode";
            this.textBoxMatCode.Size = new System.Drawing.Size(100, 20);
            this.textBoxMatCode.TabIndex = 7;
            this.textBoxMatCode.TextChanged += new System.EventHandler(this.textBoxProdCode_TextChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(853, 21);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(84, 13);
            this.label4.TabIndex = 6;
            this.label4.Text = "Код материала";
            // 
            // textBoxMatName
            // 
            this.textBoxMatName.Location = new System.Drawing.Point(949, 46);
            this.textBoxMatName.Name = "textBoxMatName";
            this.textBoxMatName.Size = new System.Drawing.Size(295, 20);
            this.textBoxMatName.TabIndex = 9;
            this.textBoxMatName.TextChanged += new System.EventHandler(this.textBoxProdCode_TextChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(959, 21);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(83, 13);
            this.label5.TabIndex = 8;
            this.label5.Text = "Наименование";
            // 
            // button5
            // 
            this.button5.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button5.Location = new System.Drawing.Point(1713, 21);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(109, 53);
            this.button5.TabIndex = 10;
            this.button5.Text = " в Excel";
            this.button5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(650, 21);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(109, 53);
            this.button1.TabIndex = 11;
            this.button1.Text = "Искать";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button9
            // 
            this.button9.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.button9.Image = ((System.Drawing.Image)(resources.GetObject("button9.Image")));
            this.button9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button9.Location = new System.Drawing.Point(1481, 21);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(109, 53);
            this.button9.TabIndex = 15;
            this.button9.Text = "       Сброс";
            this.button9.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // MercuryForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1834, 812);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MercuryForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Срок годности по Меркурию";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MercuryForm_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel3.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        private DBDataGridView dataGridView1;
        private System.Windows.Forms.TextBox textBoxMatName;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox textBoxMatCode;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox textBoxLotNo;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBoxProdName;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxProdCode;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button1;
        public System.Windows.Forms.Button button9;
    }
}