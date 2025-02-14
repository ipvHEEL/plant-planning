namespace PlantPlanning
{
    partial class MesSettingsForm
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MesSettingsForm));
            this.panel1 = new System.Windows.Forms.Panel();
            this.buttonChildOK = new System.Windows.Forms.Button();
            this.buttonParentOK = new System.Windows.Forms.Button();
            this.buttonChildCancel = new System.Windows.Forms.Button();
            this.buttonParentCancel = new System.Windows.Forms.Button();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.dataGridViewParent = new System.Windows.Forms.DataGridView();
            this.dataGridViewChild = new System.Windows.Forms.DataGridView();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewParent)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewChild)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.buttonChildOK);
            this.panel1.Controls.Add(this.buttonParentOK);
            this.panel1.Controls.Add(this.buttonChildCancel);
            this.panel1.Controls.Add(this.buttonParentCancel);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.panel1.Location = new System.Drawing.Point(0, 622);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1047, 100);
            this.panel1.TabIndex = 0;
            // 
            // buttonChildOK
            // 
            this.buttonChildOK.Image = ((System.Drawing.Image)(resources.GetObject("buttonChildOK.Image")));
            this.buttonChildOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonChildOK.Location = new System.Drawing.Point(551, 24);
            this.buttonChildOK.Name = "buttonChildOK";
            this.buttonChildOK.Size = new System.Drawing.Size(161, 53);
            this.buttonChildOK.TabIndex = 6;
            this.buttonChildOK.Text = "Использовать     ";
            this.buttonChildOK.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonChildOK.UseVisualStyleBackColor = true;
            this.buttonChildOK.Click += new System.EventHandler(this.buttonChildOK_Click);
            // 
            // buttonParentOK
            // 
            this.buttonParentOK.Image = ((System.Drawing.Image)(resources.GetObject("buttonParentOK.Image")));
            this.buttonParentOK.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonParentOK.Location = new System.Drawing.Point(12, 24);
            this.buttonParentOK.Name = "buttonParentOK";
            this.buttonParentOK.Size = new System.Drawing.Size(161, 53);
            this.buttonParentOK.TabIndex = 5;
            this.buttonParentOK.Text = "Использовать     ";
            this.buttonParentOK.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonParentOK.UseVisualStyleBackColor = true;
            this.buttonParentOK.Click += new System.EventHandler(this.button2_Click);
            // 
            // buttonChildCancel
            // 
            this.buttonChildCancel.Image = ((System.Drawing.Image)(resources.GetObject("buttonChildCancel.Image")));
            this.buttonChildCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonChildCancel.Location = new System.Drawing.Point(728, 24);
            this.buttonChildCancel.Name = "buttonChildCancel";
            this.buttonChildCancel.Size = new System.Drawing.Size(161, 53);
            this.buttonChildCancel.TabIndex = 4;
            this.buttonChildCancel.Text = "Не использовать";
            this.buttonChildCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonChildCancel.UseVisualStyleBackColor = true;
            this.buttonChildCancel.Click += new System.EventHandler(this.buttonChildCancel_Click);
            // 
            // buttonParentCancel
            // 
            this.buttonParentCancel.Image = ((System.Drawing.Image)(resources.GetObject("buttonParentCancel.Image")));
            this.buttonParentCancel.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.buttonParentCancel.Location = new System.Drawing.Point(190, 24);
            this.buttonParentCancel.Name = "buttonParentCancel";
            this.buttonParentCancel.Size = new System.Drawing.Size(161, 53);
            this.buttonParentCancel.TabIndex = 3;
            this.buttonParentCancel.Text = "Не использовать";
            this.buttonParentCancel.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.buttonParentCancel.UseVisualStyleBackColor = true;
            this.buttonParentCancel.Click += new System.EventHandler(this.buttonParentCancel_Click);
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.dataGridViewParent);
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.dataGridViewChild);
            this.splitContainer1.Size = new System.Drawing.Size(1047, 622);
            this.splitContainer1.SplitterDistance = 530;
            this.splitContainer1.TabIndex = 1;
            // 
            // dataGridViewParent
            // 
            this.dataGridViewParent.AllowUserToAddRows = false;
            this.dataGridViewParent.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewParent.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewParent.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewParent.Name = "dataGridViewParent";
            this.dataGridViewParent.ReadOnly = true;
            this.dataGridViewParent.RowHeadersWidth = 20;
            this.dataGridViewParent.Size = new System.Drawing.Size(530, 622);
            this.dataGridViewParent.TabIndex = 0;
            this.dataGridViewParent.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridViewParent_RowPrePaint);
            this.dataGridViewParent.SelectionChanged += new System.EventHandler(this.dataGridViewParent_SelectionChanged);
            // 
            // dataGridViewChild
            // 
            this.dataGridViewChild.AllowUserToAddRows = false;
            this.dataGridViewChild.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridViewChild.Dock = System.Windows.Forms.DockStyle.Fill;
            this.dataGridViewChild.Location = new System.Drawing.Point(0, 0);
            this.dataGridViewChild.Name = "dataGridViewChild";
            this.dataGridViewChild.ReadOnly = true;
            this.dataGridViewChild.RowHeadersWidth = 20;
            this.dataGridViewChild.Size = new System.Drawing.Size(513, 622);
            this.dataGridViewChild.TabIndex = 1;
            this.dataGridViewChild.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridViewChild_RowPrePaint);
            this.dataGridViewChild.SelectionChanged += new System.EventHandler(this.dataGridViewChild_SelectionChanged);
            // 
            // MesSettingsForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1047, 722);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.panel1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "MesSettingsForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Поиск остатков в МЕС";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.MesSettingsForm_FormClosed);
            this.panel1.ResumeLayout(false);
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewParent)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridViewChild)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.DataGridView dataGridViewParent;
        private System.Windows.Forms.DataGridView dataGridViewChild;
        private System.Windows.Forms.Button buttonChildCancel;
        private System.Windows.Forms.Button buttonParentCancel;
        private System.Windows.Forms.Button buttonChildOK;
        private System.Windows.Forms.Button buttonParentOK;
    }
}