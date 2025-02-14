namespace PlantPlanning
{
    partial class FormMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FormMain));
            this.menuStrip1 = new System.Windows.Forms.MenuStrip();
            this.файлToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.принятьДанныеИзExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.загрузитьКартуИзБДToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.добавитьКартуИзБДToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.сохранитьКартуВБДToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.настройкиЗаявокToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.операцииToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.потреблениеНаПроизводствеToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.показатьТекущиеОстаткиToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.выходToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.splitContainer1 = new System.Windows.Forms.SplitContainer();
            this.panel2 = new System.Windows.Forms.Panel();
            this.trackBar1 = new System.Windows.Forms.TrackBar();
            this.button20 = new System.Windows.Forms.Button();
            this.panel3 = new System.Windows.Forms.Panel();
            this.button22 = new System.Windows.Forms.Button();
            this.button21 = new System.Windows.Forms.Button();
            this.button19 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.button17 = new System.Windows.Forms.Button();
            this.button15 = new System.Windows.Forms.Button();
            this.button14 = new System.Windows.Forms.Button();
            this.button13 = new System.Windows.Forms.Button();
            this.button12 = new System.Windows.Forms.Button();
            this.button11 = new System.Windows.Forms.Button();
            this.button10 = new System.Windows.Forms.Button();
            this.button9 = new System.Windows.Forms.Button();
            this.label8 = new System.Windows.Forms.Label();
            this.panel1 = new System.Windows.Forms.Panel();
            this.button16 = new System.Windows.Forms.Button();
            this.button8 = new System.Windows.Forms.Button();
            this.button18 = new System.Windows.Forms.Button();
            this.button7 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button1 = new System.Windows.Forms.Button();
            this.dataGridView1 = new PlantPlanning.DBDataGridView();
            this.menuStrip1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).BeginInit();
            this.splitContainer1.Panel1.SuspendLayout();
            this.splitContainer1.Panel2.SuspendLayout();
            this.splitContainer1.SuspendLayout();
            this.panel2.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).BeginInit();
            this.panel3.SuspendLayout();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            this.SuspendLayout();
            // 
            // menuStrip1
            // 
            this.menuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.файлToolStripMenuItem,
            this.операцииToolStripMenuItem,
            this.выходToolStripMenuItem});
            this.menuStrip1.Location = new System.Drawing.Point(0, 0);
            this.menuStrip1.Name = "menuStrip1";
            this.menuStrip1.Size = new System.Drawing.Size(1884, 24);
            this.menuStrip1.TabIndex = 2;
            this.menuStrip1.Text = "menuStrip1";
            this.menuStrip1.Visible = false;
            // 
            // файлToolStripMenuItem
            // 
            this.файлToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.принятьДанныеИзExcelToolStripMenuItem,
            this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem,
            this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem,
            this.загрузитьКартуИзБДToolStripMenuItem,
            this.добавитьКартуИзБДToolStripMenuItem,
            this.сохранитьКартуВБДToolStripMenuItem,
            this.настройкиЗаявокToolStripMenuItem});
            this.файлToolStripMenuItem.Name = "файлToolStripMenuItem";
            this.файлToolStripMenuItem.Size = new System.Drawing.Size(62, 20);
            this.файлToolStripMenuItem.Text = "Данные";
            // 
            // принятьДанныеИзExcelToolStripMenuItem
            // 
            this.принятьДанныеИзExcelToolStripMenuItem.Name = "принятьДанныеИзExcelToolStripMenuItem";
            this.принятьДанныеИзExcelToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.принятьДанныеИзExcelToolStripMenuItem.Text = "Загрузить корректирующие данные из Excel";
            this.принятьДанныеИзExcelToolStripMenuItem.Click += new System.EventHandler(this.принятьДанныеИзExcelToolStripMenuItem_Click);
            // 
            // добавитьКорректирующиеДанныеИзExcelToolStripMenuItem
            // 
            this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem.Name = "добавитьКорректирующиеДанныеИзExcelToolStripMenuItem";
            this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem.Text = "Добавить корректирующие данные из Excel";
            this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem.Click += new System.EventHandler(this.добавитьКорректирующиеДанныеИзExcelToolStripMenuItem_Click);
            // 
            // выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem
            // 
            this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem.Name = "выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem";
            this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem.Text = "Выгрузить корректирующие данные в Excel";
            this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem.Click += new System.EventHandler(this.выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem_Click);
            // 
            // загрузитьКартуИзБДToolStripMenuItem
            // 
            this.загрузитьКартуИзБДToolStripMenuItem.Name = "загрузитьКартуИзБДToolStripMenuItem";
            this.загрузитьКартуИзБДToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.загрузитьКартуИзБДToolStripMenuItem.Text = "Загрузить карту из БД";
            this.загрузитьКартуИзБДToolStripMenuItem.Click += new System.EventHandler(this.загрузитьКартуИзБДToolStripMenuItem_Click);
            // 
            // добавитьКартуИзБДToolStripMenuItem
            // 
            this.добавитьКартуИзБДToolStripMenuItem.Name = "добавитьКартуИзБДToolStripMenuItem";
            this.добавитьКартуИзБДToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.добавитьКартуИзБДToolStripMenuItem.Text = "Добавить карту из БД";
            this.добавитьКартуИзБДToolStripMenuItem.Click += new System.EventHandler(this.добавитьКартуИзБДToolStripMenuItem_Click);
            // 
            // сохранитьКартуВБДToolStripMenuItem
            // 
            this.сохранитьКартуВБДToolStripMenuItem.Name = "сохранитьКартуВБДToolStripMenuItem";
            this.сохранитьКартуВБДToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.сохранитьКартуВБДToolStripMenuItem.Text = "Сохранить карту в БД";
            this.сохранитьКартуВБДToolStripMenuItem.Click += new System.EventHandler(this.сохранитьКартуВБДToolStripMenuItem_Click);
            // 
            // настройкиЗаявокToolStripMenuItem
            // 
            this.настройкиЗаявокToolStripMenuItem.Name = "настройкиЗаявокToolStripMenuItem";
            this.настройкиЗаявокToolStripMenuItem.Size = new System.Drawing.Size(318, 22);
            this.настройкиЗаявокToolStripMenuItem.Text = "Настройки заявок";
            this.настройкиЗаявокToolStripMenuItem.Click += new System.EventHandler(this.настройкиЗаявокToolStripMenuItem_Click);
            // 
            // операцииToolStripMenuItem
            // 
            this.операцииToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.потреблениеНаПроизводствеToolStripMenuItem,
            this.показатьТекущиеОстаткиToolStripMenuItem});
            this.операцииToolStripMenuItem.Name = "операцииToolStripMenuItem";
            this.операцииToolStripMenuItem.Size = new System.Drawing.Size(75, 20);
            this.операцииToolStripMenuItem.Text = "Операции";
            // 
            // потреблениеНаПроизводствеToolStripMenuItem
            // 
            this.потреблениеНаПроизводствеToolStripMenuItem.Name = "потреблениеНаПроизводствеToolStripMenuItem";
            this.потреблениеНаПроизводствеToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.потреблениеНаПроизводствеToolStripMenuItem.Text = "Потребление на производстве";
            this.потреблениеНаПроизводствеToolStripMenuItem.Click += new System.EventHandler(this.потреблениеНаПроизводствеToolStripMenuItem_Click);
            // 
            // показатьТекущиеОстаткиToolStripMenuItem
            // 
            this.показатьТекущиеОстаткиToolStripMenuItem.Name = "показатьТекущиеОстаткиToolStripMenuItem";
            this.показатьТекущиеОстаткиToolStripMenuItem.Size = new System.Drawing.Size(242, 22);
            this.показатьТекущиеОстаткиToolStripMenuItem.Text = "Показать текущие остатки";
            this.показатьТекущиеОстаткиToolStripMenuItem.Click += new System.EventHandler(this.показатьТекущиеОстаткиToolStripMenuItem_Click);
            // 
            // выходToolStripMenuItem
            // 
            this.выходToolStripMenuItem.Name = "выходToolStripMenuItem";
            this.выходToolStripMenuItem.Size = new System.Drawing.Size(54, 20);
            this.выходToolStripMenuItem.Text = "Выход";
            this.выходToolStripMenuItem.Click += new System.EventHandler(this.выходToolStripMenuItem_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            this.openFileDialog1.Filter = "Файлы Эксель |*.xls*| Xlsx files (*.xlsx)|*.xlsx";
            // 
            // splitContainer1
            // 
            this.splitContainer1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.splitContainer1.Location = new System.Drawing.Point(0, 0);
            this.splitContainer1.Name = "splitContainer1";
            this.splitContainer1.Orientation = System.Windows.Forms.Orientation.Horizontal;
            // 
            // splitContainer1.Panel1
            // 
            this.splitContainer1.Panel1.Controls.Add(this.panel2);
            this.splitContainer1.Panel1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            // 
            // splitContainer1.Panel2
            // 
            this.splitContainer1.Panel2.Controls.Add(this.panel1);
            this.splitContainer1.Panel2.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.splitContainer1.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.splitContainer1.Size = new System.Drawing.Size(1884, 881);
            this.splitContainer1.SplitterDistance = 795;
            this.splitContainer1.TabIndex = 6;
            // 
            // panel2
            // 
            this.panel2.Controls.Add(this.trackBar1);
            this.panel2.Controls.Add(this.button20);
            this.panel2.Controls.Add(this.panel3);
            this.panel2.Controls.Add(this.label8);
            this.panel2.Controls.Add(this.dataGridView1);
            this.panel2.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel2.Location = new System.Drawing.Point(0, 0);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(1884, 795);
            this.panel2.TabIndex = 7;
            // 
            // trackBar1
            // 
            this.trackBar1.Location = new System.Drawing.Point(1040, 683);
            this.trackBar1.Maximum = 100;
            this.trackBar1.Minimum = 40;
            this.trackBar1.Name = "trackBar1";
            this.trackBar1.Size = new System.Drawing.Size(180, 45);
            this.trackBar1.TabIndex = 17;
            this.trackBar1.Value = 50;
            this.trackBar1.Visible = false;
            // 
            // button20
            // 
            this.button20.Image = ((System.Drawing.Image)(resources.GetObject("button20.Image")));
            this.button20.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button20.Location = new System.Drawing.Point(1247, 667);
            this.button20.Name = "button20";
            this.button20.Size = new System.Drawing.Size(195, 56);
            this.button20.TabIndex = 20;
            this.button20.Text = "                        Неликвид";
            this.button20.UseVisualStyleBackColor = true;
            this.button20.Visible = false;
            this.button20.Click += new System.EventHandler(this.button20_Click);
            // 
            // panel3
            // 
            this.panel3.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel3.Controls.Add(this.button22);
            this.panel3.Controls.Add(this.button21);
            this.panel3.Controls.Add(this.button19);
            this.panel3.Controls.Add(this.label1);
            this.panel3.Controls.Add(this.progressBar1);
            this.panel3.Controls.Add(this.button17);
            this.panel3.Controls.Add(this.button15);
            this.panel3.Controls.Add(this.button14);
            this.panel3.Controls.Add(this.button13);
            this.panel3.Controls.Add(this.button12);
            this.panel3.Controls.Add(this.button11);
            this.panel3.Controls.Add(this.button10);
            this.panel3.Controls.Add(this.button9);
            this.panel3.Location = new System.Drawing.Point(1677, 3);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(207, 789);
            this.panel3.TabIndex = 5;
            // 
            // button22
            // 
            this.button22.Image = ((System.Drawing.Image)(resources.GetObject("button22.Image")));
            this.button22.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button22.Location = new System.Drawing.Point(7, 632);
            this.button22.Name = "button22";
            this.button22.Size = new System.Drawing.Size(195, 56);
            this.button22.TabIndex = 21;
            this.button22.Text = "                             Комментарии по                         компонентам";
            this.button22.UseVisualStyleBackColor = true;
            this.button22.Click += new System.EventHandler(this.button22_Click);
            // 
            // button21
            // 
            this.button21.Image = ((System.Drawing.Image)(resources.GetObject("button21.Image")));
            this.button21.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button21.Location = new System.Drawing.Point(7, 570);
            this.button21.Name = "button21";
            this.button21.Size = new System.Drawing.Size(195, 56);
            this.button21.TabIndex = 20;
            this.button21.Text = "                             Просрочка по                         Меркурию";
            this.button21.UseVisualStyleBackColor = true;
            this.button21.Click += new System.EventHandler(this.button21_Click);
            // 
            // button19
            // 
            this.button19.Image = ((System.Drawing.Image)(resources.GetObject("button19.Image")));
            this.button19.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button19.Location = new System.Drawing.Point(7, 384);
            this.button19.Name = "button19";
            this.button19.Size = new System.Drawing.Size(195, 56);
            this.button19.TabIndex = 19;
            this.button19.Text = "                                 Глючные                                    специ" +
    "фикации";
            this.button19.UseVisualStyleBackColor = true;
            this.button19.Click += new System.EventHandler(this.button19_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(18, 723);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(64, 13);
            this.label1.TabIndex = 18;
            this.label1.Text = "Выгружено";
            this.label1.Visible = false;
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(3, 748);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(199, 23);
            this.progressBar1.TabIndex = 17;
            this.progressBar1.Visible = false;
            // 
            // button17
            // 
            this.button17.Image = ((System.Drawing.Image)(resources.GetObject("button17.Image")));
            this.button17.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button17.Location = new System.Drawing.Point(7, 446);
            this.button17.Name = "button17";
            this.button17.Size = new System.Drawing.Size(195, 56);
            this.button17.TabIndex = 15;
            this.button17.Text = "                            Обновление                              спецификаций";
            this.button17.UseVisualStyleBackColor = true;
            this.button17.Click += new System.EventHandler(this.button17_Click);
            // 
            // button15
            // 
            this.button15.Image = ((System.Drawing.Image)(resources.GetObject("button15.Image")));
            this.button15.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button15.Location = new System.Drawing.Point(7, 508);
            this.button15.Name = "button15";
            this.button15.Size = new System.Drawing.Size(195, 56);
            this.button15.TabIndex = 14;
            this.button15.Text = "                        Поиск задвоенных                        норм";
            this.button15.UseVisualStyleBackColor = true;
            this.button15.Click += new System.EventHandler(this.button15_Click);
            // 
            // button14
            // 
            this.button14.Image = ((System.Drawing.Image)(resources.GetObject("button14.Image")));
            this.button14.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button14.Location = new System.Drawing.Point(7, 322);
            this.button14.Name = "button14";
            this.button14.Size = new System.Drawing.Size(195, 56);
            this.button14.TabIndex = 13;
            this.button14.Text = "                                 Экспорт                                    специ" +
    "фикаций";
            this.button14.UseVisualStyleBackColor = true;
            this.button14.Click += new System.EventHandler(this.button14_Click);
            // 
            // button13
            // 
            this.button13.Image = ((System.Drawing.Image)(resources.GetObject("button13.Image")));
            this.button13.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button13.Location = new System.Drawing.Point(7, 260);
            this.button13.Name = "button13";
            this.button13.Size = new System.Drawing.Size(195, 56);
            this.button13.TabIndex = 12;
            this.button13.Text = "ОПР";
            this.button13.UseVisualStyleBackColor = true;
            this.button13.Click += new System.EventHandler(this.button13_Click);
            // 
            // button12
            // 
            this.button12.Image = ((System.Drawing.Image)(resources.GetObject("button12.Image")));
            this.button12.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button12.Location = new System.Drawing.Point(7, 198);
            this.button12.Name = "button12";
            this.button12.Size = new System.Drawing.Size(195, 56);
            this.button12.TabIndex = 11;
            this.button12.Text = "      Настройки поиска в МЕС";
            this.button12.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button12.UseVisualStyleBackColor = true;
            this.button12.Click += new System.EventHandler(this.button12_Click);
            // 
            // button11
            // 
            this.button11.Image = ((System.Drawing.Image)(resources.GetObject("button11.Image")));
            this.button11.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button11.Location = new System.Drawing.Point(7, 136);
            this.button11.Name = "button11";
            this.button11.Size = new System.Drawing.Size(195, 56);
            this.button11.TabIndex = 10;
            this.button11.Text = "      Сырье из Навижн";
            this.button11.UseVisualStyleBackColor = true;
            this.button11.Click += new System.EventHandler(this.button11_Click);
            // 
            // button10
            // 
            this.button10.Image = ((System.Drawing.Image)(resources.GetObject("button10.Image")));
            this.button10.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button10.Location = new System.Drawing.Point(7, 12);
            this.button10.Name = "button10";
            this.button10.Size = new System.Drawing.Size(195, 56);
            this.button10.TabIndex = 9;
            this.button10.Text = "Рассчитать потребность";
            this.button10.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button10.UseVisualStyleBackColor = true;
            this.button10.Click += new System.EventHandler(this.button10_Click);
            // 
            // button9
            // 
            this.button9.Image = ((System.Drawing.Image)(resources.GetObject("button9.Image")));
            this.button9.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button9.Location = new System.Drawing.Point(7, 74);
            this.button9.Name = "button9";
            this.button9.Size = new System.Drawing.Size(195, 56);
            this.button9.TabIndex = 8;
            this.button9.Text = "Остатки";
            this.button9.UseVisualStyleBackColor = true;
            this.button9.Click += new System.EventHandler(this.button9_Click);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(1056, 667);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(157, 13);
            this.label8.TabIndex = 18;
            this.label8.Text = "Записей на странице (40-100)";
            this.label8.Visible = false;
            // 
            // panel1
            // 
            this.panel1.Controls.Add(this.button16);
            this.panel1.Controls.Add(this.button8);
            this.panel1.Controls.Add(this.button18);
            this.panel1.Controls.Add(this.button7);
            this.panel1.Controls.Add(this.button6);
            this.panel1.Controls.Add(this.button5);
            this.panel1.Controls.Add(this.button4);
            this.panel1.Controls.Add(this.button3);
            this.panel1.Controls.Add(this.button2);
            this.panel1.Controls.Add(this.button1);
            this.panel1.Dock = System.Windows.Forms.DockStyle.Fill;
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(1884, 82);
            this.panel1.TabIndex = 6;
            // 
            // button16
            // 
            this.button16.Enabled = false;
            this.button16.Image = ((System.Drawing.Image)(resources.GetObject("button16.Image")));
            this.button16.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button16.Location = new System.Drawing.Point(1007, 7);
            this.button16.Name = "button16";
            this.button16.Size = new System.Drawing.Size(154, 53);
            this.button16.TabIndex = 8;
            this.button16.Text = "Добавить из БД";
            this.button16.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button16.UseVisualStyleBackColor = true;
            this.button16.Visible = false;
            this.button16.Click += new System.EventHandler(this.button16_Click);
            // 
            // button8
            // 
            this.button8.Enabled = false;
            this.button8.Image = ((System.Drawing.Image)(resources.GetObject("button8.Image")));
            this.button8.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button8.Location = new System.Drawing.Point(1508, 7);
            this.button8.Name = "button8";
            this.button8.Size = new System.Drawing.Size(157, 53);
            this.button8.TabIndex = 7;
            this.button8.Text = "Сохранить в БД";
            this.button8.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button8.UseVisualStyleBackColor = true;
            this.button8.Visible = false;
            this.button8.Click += new System.EventHandler(this.button8_Click);
            // 
            // button18
            // 
            this.button18.Location = new System.Drawing.Point(1642, 59);
            this.button18.Name = "button18";
            this.button18.Size = new System.Drawing.Size(75, 23);
            this.button18.TabIndex = 16;
            this.button18.Text = "button18";
            this.button18.UseVisualStyleBackColor = true;
            this.button18.Visible = false;
            this.button18.Click += new System.EventHandler(this.button18_Click);
            // 
            // button7
            // 
            this.button7.Enabled = false;
            this.button7.Image = ((System.Drawing.Image)(resources.GetObject("button7.Image")));
            this.button7.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button7.Location = new System.Drawing.Point(847, 7);
            this.button7.Name = "button7";
            this.button7.Size = new System.Drawing.Size(154, 53);
            this.button7.TabIndex = 6;
            this.button7.Text = "Загрузить из БД";
            this.button7.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button7.UseVisualStyleBackColor = true;
            this.button7.Visible = false;
            this.button7.Click += new System.EventHandler(this.button7_Click);
            // 
            // button6
            // 
            this.button6.Image = ((System.Drawing.Image)(resources.GetObject("button6.Image")));
            this.button6.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button6.Location = new System.Drawing.Point(699, 7);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(142, 53);
            this.button6.TabIndex = 5;
            this.button6.Text = "Загрузить карту";
            this.button6.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button6.UseVisualStyleBackColor = true;
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // button5
            // 
            this.button5.Image = ((System.Drawing.Image)(resources.GetObject("button5.Image")));
            this.button5.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button5.Location = new System.Drawing.Point(1334, 7);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(157, 53);
            this.button5.TabIndex = 4;
            this.button5.Text = "Выгрузить в Excel";
            this.button5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button4
            // 
            this.button4.Enabled = false;
            this.button4.Image = ((System.Drawing.Image)(resources.GetObject("button4.Image")));
            this.button4.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button4.Location = new System.Drawing.Point(551, 7);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(142, 53);
            this.button4.TabIndex = 3;
            this.button4.Text = "Добавить карту";
            this.button4.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button3
            // 
            this.button3.Image = ((System.Drawing.Image)(resources.GetObject("button3.Image")));
            this.button3.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button3.Location = new System.Drawing.Point(411, 7);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(134, 53);
            this.button3.TabIndex = 2;
            this.button3.Text = "Убрать все";
            this.button3.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button2
            // 
            this.button2.Image = ((System.Drawing.Image)(resources.GetObject("button2.Image")));
            this.button2.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button2.Location = new System.Drawing.Point(271, 7);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(134, 53);
            this.button2.TabIndex = 1;
            this.button2.Text = "Убрать день";
            this.button2.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button1
            // 
            this.button1.Image = ((System.Drawing.Image)(resources.GetObject("button1.Image")));
            this.button1.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.button1.Location = new System.Drawing.Point(12, 7);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(167, 53);
            this.button1.TabIndex = 0;
            this.button1.Text = "Рецепты продуктов";
            this.button1.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // dataGridView1
            // 
            this.dataGridView1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.DoubleBuffered = true;
            this.dataGridView1.Location = new System.Drawing.Point(3, 3);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(1675, 789);
            this.dataGridView1.TabIndex = 4;
            this.dataGridView1.RowPrePaint += new System.Windows.Forms.DataGridViewRowPrePaintEventHandler(this.dataGridView1_RowPrePaint);
            // 
            // FormMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1884, 881);
            this.Controls.Add(this.splitContainer1);
            this.Controls.Add(this.menuStrip1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "FormMain";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Планирование производства фабрика-кухня";
            this.FormClosed += new System.Windows.Forms.FormClosedEventHandler(this.FormMain_FormClosed);
            this.Resize += new System.EventHandler(this.FormMain_Resize);
            this.menuStrip1.ResumeLayout(false);
            this.menuStrip1.PerformLayout();
            this.splitContainer1.Panel1.ResumeLayout(false);
            this.splitContainer1.Panel2.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.splitContainer1)).EndInit();
            this.splitContainer1.ResumeLayout(false);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.trackBar1)).EndInit();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel1.ResumeLayout(false);
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        public System.Windows.Forms.MenuStrip menuStrip1;
        private System.Windows.Forms.ToolStripMenuItem файлToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem принятьДанныеИзExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem настройкиЗаявокToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem операцииToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem потреблениеНаПроизводствеToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem показатьТекущиеОстаткиToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem выходToolStripMenuItem;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.ToolStripMenuItem выгрузитьКорректирующиеДанныеИзExcelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem сохранитьКартуВБДToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem загрузитьКартуИзБДToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem добавитьКартуИзБДToolStripMenuItem;
        private System.Windows.Forms.SplitContainer splitContainer1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Panel panel3;
        //  private System.Windows.Forms.DataGridView dataGridView1;
        private  DBDataGridView dataGridView1;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Button button7;
        private System.Windows.Forms.Button button6;
        private System.Windows.Forms.Button button8;
        private System.Windows.Forms.Button button10;
        private System.Windows.Forms.Button button9;
        private System.Windows.Forms.Button button11;
        private System.Windows.Forms.Button button12;
        private System.Windows.Forms.Button button13;
        private System.Windows.Forms.Button button15;
        private System.Windows.Forms.Button button14;
        private System.Windows.Forms.ToolStripMenuItem добавитьКорректирующиеДанныеИзExcelToolStripMenuItem;
        private System.Windows.Forms.Button button16;
        private System.Windows.Forms.Button button17;
        private System.Windows.Forms.Button button18;
        public System.Windows.Forms.ProgressBar progressBar1;
        public System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button button19;
        private System.Windows.Forms.Button button20;
        public System.Windows.Forms.TrackBar trackBar1;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.Button button21;
        private System.Windows.Forms.Button button22;
    }
}

