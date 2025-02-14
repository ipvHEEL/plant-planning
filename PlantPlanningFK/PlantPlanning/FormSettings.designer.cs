namespace PlantPlanning
{
    partial class FormSettings
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormSettings));
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label5 = new System.Windows.Forms.Label();
            this.tbUrLic = new System.Windows.Forms.TextBox();
            this.tbObject = new System.Windows.Forms.TextBox();
            this.tbCFOCode = new System.Windows.Forms.TextBox();
            this.tbCFOSogl = new System.Windows.Forms.TextBox();
            this.tbAddress = new System.Windows.Forms.TextBox();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label1.Location = new System.Drawing.Point(25, 30);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(56, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Юр.лицо";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label2.Location = new System.Drawing.Point(25, 72);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(84, 13);
            this.label2.TabIndex = 1;
            this.label2.Text = "Для объекта";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label3.Location = new System.Drawing.Point(25, 114);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(63, 13);
            this.label3.TabIndex = 2;
            this.label3.Text = "ЦФО Код";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label4.Location = new System.Drawing.Point(25, 156);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(124, 13);
            this.label4.TabIndex = 3;
            this.label4.Text = "ЦФО согласования";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label5.Location = new System.Drawing.Point(25, 198);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(105, 13);
            this.label5.TabIndex = 4;
            this.label5.Text = "Адрес Доставки";
            // 
            // tbUrLic
            // 
            this.tbUrLic.Location = new System.Drawing.Point(203, 23);
            this.tbUrLic.Name = "tbUrLic";
            this.tbUrLic.Size = new System.Drawing.Size(585, 20);
            this.tbUrLic.TabIndex = 5;
            // 
            // tbObject
            // 
            this.tbObject.Location = new System.Drawing.Point(203, 65);
            this.tbObject.Name = "tbObject";
            this.tbObject.Size = new System.Drawing.Size(585, 20);
            this.tbObject.TabIndex = 6;
            // 
            // tbCFOCode
            // 
            this.tbCFOCode.Location = new System.Drawing.Point(203, 111);
            this.tbCFOCode.Name = "tbCFOCode";
            this.tbCFOCode.Size = new System.Drawing.Size(585, 20);
            this.tbCFOCode.TabIndex = 7;
            // 
            // tbCFOSogl
            // 
            this.tbCFOSogl.Location = new System.Drawing.Point(203, 153);
            this.tbCFOSogl.Name = "tbCFOSogl";
            this.tbCFOSogl.Size = new System.Drawing.Size(585, 20);
            this.tbCFOSogl.TabIndex = 8;
            // 
            // tbAddress
            // 
            this.tbAddress.Location = new System.Drawing.Point(203, 191);
            this.tbAddress.Multiline = true;
            this.tbAddress.Name = "tbAddress";
            this.tbAddress.Size = new System.Drawing.Size(585, 107);
            this.tbAddress.TabIndex = 9;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(28, 415);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 10;
            this.button1.Text = "Сохранить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(713, 415);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 11;
            this.button2.Text = "Отмена";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // FormSettings
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(800, 450);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.tbAddress);
            this.Controls.Add(this.tbCFOSogl);
            this.Controls.Add(this.tbCFOCode);
            this.Controls.Add(this.tbObject);
            this.Controls.Add(this.tbUrLic);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FormSettings";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройки";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormSettings_FormClosed);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.TextBox tbUrLic;
        private System.Windows.Forms.TextBox tbObject;
        private System.Windows.Forms.TextBox tbCFOCode;
        private System.Windows.Forms.TextBox tbCFOSogl;
        private System.Windows.Forms.TextBox tbAddress;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button2;
    }
}