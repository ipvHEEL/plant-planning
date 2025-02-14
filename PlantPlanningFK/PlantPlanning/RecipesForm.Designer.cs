namespace PlantPlanning
{
    partial class RecipesForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(RecipesForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.textBoxQuantity = new System.Windows.Forms.TextBox();
            this.dateTimePicker1 = new System.Windows.Forms.DateTimePicker();
            this.comboBoxLines = new System.Windows.Forms.ComboBox();
            this.textBoxPlan = new System.Windows.Forms.TextBox();
            this.label1 = new System.Windows.Forms.Label();
            this.dataGridViewDetails = new System.Windows.Forms.DataGridView();
            this.MatCode = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MatName = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MatClass = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.MatQuantity = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewDates = new System.Windows.Forms.DataGridView();
            this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewSpec = new System.Windows.Forms.DataGridView();
            this.Specify = new System.Windows.Forms.DataGridViewTextBoxColumn();
            this.dataGridViewMain = new System.Windows.Forms.DataGridView();
            this.textBoxCode = new System.Windows.Forms.TextBox();
            this.textBoxNaim = new System.Windows.Forms.TextBox();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDetails)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDates)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSpec)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMain)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.label5);
            this.panel1.Controls.Add(this.label4);
            this.panel1.Controls.Add(this.label3);
            this.panel1.Controls.Add(this.label2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.textBoxQuantity);
            this.panel1.Controls.Add(this.dateTimePicker1);
            this.panel1.Controls.Add(this.comboBoxLines);
            this.panel1.Controls.Add(this.textBoxPlan);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.dataGridViewDetails);
            this.panel1.Controls.Add(this.dataGridViewDates);
            this.panel1.Controls.Add(this.dataGridViewSpec);
            this.panel1.Location = new System.Drawing.Point(0, 437);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1685, 324);
            this.panel1.TabIndex = 0;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(1499, 178);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(128, 15);
            this.label5.TabIndex = 13;
            this.label5.Text = "Объем заявки (кг)";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(1499, 120);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(139, 15);
            this.label4.TabIndex = 12;
            this.label4.Text = "Дата производства";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(1476, 65);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(179, 15);
            this.label3.TabIndex = 11;
            this.label3.Text = "Производственная линия";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(1529, 6);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(64, 15);
            this.label2.TabIndex = 10;
            this.label2.Text = "№ плана";
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(1552, 260);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(120, 53);
            this.button1.TabIndex = 9;
            this.button1.Text = "                  Добавить    и закрыть";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button4
            // 
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.Location = new System.Drawing.Point(1415, 260);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(120, 53);
            this.button4.TabIndex = 8;
            this.button4.Text = "                  Добавить  в план";
            this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // textBoxQuantity
            // 
            this.textBoxQuantity.Location = new System.Drawing.Point(1475, 196);
            this.textBoxQuantity.Name = "textBoxQuantity";
            this.textBoxQuantity.Size = new System.Drawing.Size(183, 20);
            this.textBoxQuantity.TabIndex = 7;
            // 
            // dateTimePicker1
            // 
            this.dateTimePicker1.Location = new System.Drawing.Point(1475, 138);
            this.dateTimePicker1.Name = "dateTimePicker1";
            this.dateTimePicker1.Size = new System.Drawing.Size(183, 20);
            this.dateTimePicker1.TabIndex = 6;
            // 
            // comboBoxLines
            // 
            this.comboBoxLines.FormattingEnabled = true;
            this.comboBoxLines.Location = new System.Drawing.Point(1475, 83);
            this.comboBoxLines.Name = "comboBoxLines";
            this.comboBoxLines.Size = new System.Drawing.Size(183, 21);
            this.comboBoxLines.TabIndex = 5;
            // 
            // textBoxPlan
            // 
            this.textBoxPlan.Location = new System.Drawing.Point(1475, 27);
            this.textBoxPlan.Name = "textBoxPlan";
            this.textBoxPlan.Size = new System.Drawing.Size(183, 20);
            this.textBoxPlan.TabIndex = 4;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(285, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(198, 15);
            this.label1.TabIndex = 3;
            this.label1.Text = "Подробности спецификации";
            // 
            // dataGridViewDetails
            // 
            this.dataGridViewDetails.AllowUserToAddRows = false;
            this.dataGridViewDetails.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDetails.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.MatCode,
            this.MatName,
            this.MatClass,
            this.MatQuantity});
            this.dataGridViewDetails.Location = new System.Drawing.Point(249, 27);
            this.dataGridViewDetails.Name = "dataGridViewDetails";
            this.dataGridViewDetails.ReadOnly = true;
            this.dataGridViewDetails.RowHeadersWidth = 20;
            this.dataGridViewDetails.Size = new System.Drawing.Size(1149, 294);
            this.dataGridViewDetails.TabIndex = 2;
            // 
            // MatCode
            // 
            this.MatCode.HeaderText = "Код материала";
            this.MatCode.Name = "MatCode";
            this.MatCode.ReadOnly = true;
            this.MatCode.Width = 120;
            // 
            // MatName
            // 
            this.MatName.HeaderText = "Наименование";
            this.MatName.Name = "MatName";
            this.MatName.ReadOnly = true;
            this.MatName.Width = 750;
            // 
            // MatClass
            // 
            this.MatClass.HeaderText = "Класс материала";
            this.MatClass.Name = "MatClass";
            this.MatClass.ReadOnly = true;
            this.MatClass.Width = 120;
            // 
            // MatQuantity
            // 
            this.MatQuantity.HeaderText = "К-во (на тонну)";
            this.MatQuantity.Name = "MatQuantity";
            this.MatQuantity.ReadOnly = true;
            this.MatQuantity.Width = 120;
            // 
            // dataGridViewDates
            // 
            this.dataGridViewDates.AllowUserToAddRows = false;
            this.dataGridViewDates.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewDates.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1});
            this.dataGridViewDates.Location = new System.Drawing.Point(3, 149);
            this.dataGridViewDates.Name = "dataGridViewDates";
            this.dataGridViewDates.ReadOnly = true;
            this.dataGridViewDates.RowHeadersWidth = 25;
            this.dataGridViewDates.Size = new System.Drawing.Size(240, 172);
            this.dataGridViewDates.TabIndex = 1;
            this.dataGridViewDates.SelectionChanged += new System.EventHandler(this.dataGridViewDates_SelectionChanged);
            // 
            // dataGridViewTextBoxColumn1
            // 
            this.dataGridViewTextBoxColumn1.HeaderText = "Периоды";
            this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
            this.dataGridViewTextBoxColumn1.ReadOnly = true;
            this.dataGridViewTextBoxColumn1.Width = 200;
            // 
            // dataGridViewSpec
            // 
            this.dataGridViewSpec.AllowUserToAddRows = false;
            this.dataGridViewSpec.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewSpec.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Specify});
            this.dataGridViewSpec.Location = new System.Drawing.Point(3, 3);
            this.dataGridViewSpec.Name = "dataGridViewSpec";
            this.dataGridViewSpec.ReadOnly = true;
            this.dataGridViewSpec.RowHeadersWidth = 25;
            this.dataGridViewSpec.Size = new System.Drawing.Size(240, 133);
            this.dataGridViewSpec.TabIndex = 0;
            this.dataGridViewSpec.SelectionChanged += new System.EventHandler(this.dataGridViewSpec_SelectionChanged);
            // 
            // Specify
            // 
            this.Specify.HeaderText = "Спецификации";
            this.Specify.Name = "Specify";
            this.Specify.ReadOnly = true;
            this.Specify.Width = 200;
            // 
            // dataGridViewMain
            // 
            this.dataGridViewMain.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridViewMain.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewMain.Location = new System.Drawing.Point(0, 41);
            this.dataGridViewMain.Name = "dataGridViewMain";
            this.dataGridViewMain.RowHeadersWidth = 4;
            this.dataGridViewMain.Size = new System.Drawing.Size(1685, 390);
            this.dataGridViewMain.TabIndex = 1;
            this.dataGridViewMain.SelectionChanged += new System.EventHandler(this.dataGridViewMain_SelectionChanged);
            // 
            // textBoxCode
            // 
            this.textBoxCode.Location = new System.Drawing.Point(12, 11);
            this.textBoxCode.Name = "textBoxCode";
            this.textBoxCode.Size = new System.Drawing.Size(123, 20);
            this.textBoxCode.TabIndex = 7;
            this.textBoxCode.TextChanged += new System.EventHandler(this.textBoxCode_TextChanged);
            // 
            // textBoxNaim
            // 
            this.textBoxNaim.Location = new System.Drawing.Point(141, 11);
            this.textBoxNaim.Name = "textBoxNaim";
            this.textBoxNaim.Size = new System.Drawing.Size(1514, 20);
            this.textBoxNaim.TabIndex = 8;
            this.textBoxNaim.TextChanged += new System.EventHandler(this.textBoxCode_TextChanged);
            // 
            // RecipesForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1684, 762);
            this.Controls.Add(this.textBoxNaim);
            this.Controls.Add(this.textBoxCode);
            this.Controls.Add(this.dataGridViewMain);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "RecipesForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Рецепты Мираторг ЗАПАД";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.RecipesForm_FormClosed);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDetails)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewDates)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewSpec)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewMain)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.DataGridView dataGridViewMain;
        private System.Windows.Forms.DataGridView dataGridViewSpec;
        private System.Windows.Forms.DataGridViewTextBoxColumn Specify;
        private System.Windows.Forms.DataGridView dataGridViewDates;
        private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.DataGridView dataGridViewDetails;
        private System.Windows.Forms.TextBox textBoxQuantity;
        private System.Windows.Forms.DateTimePicker dateTimePicker1;
        private System.Windows.Forms.ComboBox comboBoxLines;
        private System.Windows.Forms.TextBox textBoxPlan;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBoxCode;
        private System.Windows.Forms.TextBox textBoxNaim;
        private System.Windows.Forms.DataGridViewTextBoxColumn MatCode;
        private System.Windows.Forms.DataGridViewTextBoxColumn MatName;
        private System.Windows.Forms.DataGridViewTextBoxColumn MatClass;
        private System.Windows.Forms.DataGridViewTextBoxColumn MatQuantity;
    }
}