namespace ZapolnenieDOC
{
    partial class Form1
    {
        /// <summary>
        /// Обязательная переменная конструктора.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Освободить все используемые ресурсы.
        /// </summary>
        /// <param name="disposing">истинно, если управляемый ресурс должен быть удален; иначе ложно.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Код, автоматически созданный конструктором форм Windows

        /// <summary>
        /// Требуемый метод для поддержки конструктора — не изменяйте 
        /// содержимое этого метода с помощью редактора кода.
        /// </summary>
        private void InitializeComponent()
        {
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.button1 = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.listBox1 = new System.Windows.Forms.ListBox();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label4 = new System.Windows.Forms.Label();
            this.button5 = new System.Windows.Forms.Button();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.progressBar2 = new System.Windows.Forms.ProgressBar();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox3 = new System.Windows.Forms.TextBox();
            this.listBox2 = new System.Windows.Forms.ListBox();
            this.label5 = new System.Windows.Forms.Label();
            this.dataGridView1 = new System.Windows.Forms.DataGridView();
            this.dataGridView2 = new System.Windows.Forms.DataGridView();
            this.dataGridView3 = new System.Windows.Forms.DataGridView();
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.label8 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).BeginInit();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Font = new System.Drawing.Font("Rage Italic", 36F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.button1.Location = new System.Drawing.Point(1413, 722);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(180, 88);
            this.button1.TabIndex = 0;
            this.button1.Text = "START";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(687, 316);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(383, 13);
            this.label1.TabIndex = 3;
            this.label1.Text = "Список студентов у которых была произведена вставка данных паспорта";
            // 
            // listBox1
            // 
            this.listBox1.FormattingEnabled = true;
            this.listBox1.Location = new System.Drawing.Point(690, 333);
            this.listBox1.Name = "listBox1";
            this.listBox1.Size = new System.Drawing.Size(413, 225);
            this.listBox1.TabIndex = 4;
            // 
            // button3
            // 
            this.button3.Location = new System.Drawing.Point(521, 216);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(141, 54);
            this.button3.TabIndex = 5;
            this.button3.Text = "Сохранить Результат Заполнения в ексель";
            this.button3.UseVisualStyleBackColor = true;
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Location = new System.Drawing.Point(521, 27);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(141, 52);
            this.button4.TabIndex = 6;
            this.button4.Text = "Поместить данные в таблицу сверху";
            this.button4.UseVisualStyleBackColor = true;
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(702, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(164, 13);
            this.label2.TabIndex = 7;
            this.label2.Text = "Путь к файлу с данными (docx)";
            // 
            // textBox1
            // 
            this.textBox1.Location = new System.Drawing.Point(705, 28);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(286, 20);
            this.textBox1.TabIndex = 8;
            this.textBox1.Click += new System.EventHandler(this.textBox1_Click);
            // 
            // textBox2
            // 
            this.textBox2.Location = new System.Drawing.Point(705, 144);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(286, 20);
            this.textBox2.TabIndex = 10;
            this.textBox2.Click += new System.EventHandler(this.textBox2_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(702, 127);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(164, 13);
            this.label4.TabIndex = 11;
            this.label4.Text = "Путь к файлу с данными (docx)";
            // 
            // button5
            // 
            this.button5.Location = new System.Drawing.Point(521, 127);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(141, 52);
            this.button5.TabIndex = 13;
            this.button5.Text = "Поместить данные в таблицу снизу";
            this.button5.UseVisualStyleBackColor = true;
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(705, 54);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(286, 15);
            this.progressBar1.TabIndex = 15;
            // 
            // progressBar2
            // 
            this.progressBar2.Location = new System.Drawing.Point(705, 170);
            this.progressBar2.Name = "progressBar2";
            this.progressBar2.Size = new System.Drawing.Size(286, 15);
            this.progressBar2.TabIndex = 16;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(716, 217);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(174, 13);
            this.label3.TabIndex = 17;
            this.label3.Text = "Папка куда сохранить файл(exel)";
            // 
            // textBox3
            // 
            this.textBox3.Location = new System.Drawing.Point(705, 250);
            this.textBox3.Name = "textBox3";
            this.textBox3.Size = new System.Drawing.Size(286, 20);
            this.textBox3.TabIndex = 18;
            this.textBox3.Click += new System.EventHandler(this.textBox3_Click);
            // 
            // listBox2
            // 
            this.listBox2.FormattingEnabled = true;
            this.listBox2.Location = new System.Drawing.Point(690, 599);
            this.listBox2.Name = "listBox2";
            this.listBox2.Size = new System.Drawing.Size(413, 199);
            this.listBox2.TabIndex = 20;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(687, 567);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(312, 26);
            this.label5.TabIndex = 19;
            this.label5.Text = "Список студентов у которых были некоректные данные или\r\n их не было в таблице сни" +
    "зу";
            // 
            // dataGridView1
            // 
            this.dataGridView1.AllowUserToAddRows = false;
            this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView1.Location = new System.Drawing.Point(12, 12);
            this.dataGridView1.Name = "dataGridView1";
            this.dataGridView1.Size = new System.Drawing.Size(503, 385);
            this.dataGridView1.TabIndex = 21;
            // 
            // dataGridView2
            // 
            this.dataGridView2.AllowUserToAddRows = false;
            this.dataGridView2.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView2.Location = new System.Drawing.Point(12, 403);
            this.dataGridView2.Name = "dataGridView2";
            this.dataGridView2.Size = new System.Drawing.Size(503, 407);
            this.dataGridView2.TabIndex = 22;
            // 
            // dataGridView3
            // 
            this.dataGridView3.AllowUserToAddRows = false;
            this.dataGridView3.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
            this.dataGridView3.Location = new System.Drawing.Point(1109, 98);
            this.dataGridView3.Name = "dataGridView3";
            this.dataGridView3.Size = new System.Drawing.Size(483, 618);
            this.dataGridView3.TabIndex = 23;
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.label6.Location = new System.Drawing.Point(1103, 54);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(302, 31);
            this.label6.TabIndex = 24;
            this.label6.Text = "Результат Заполнения";
            this.label6.Click += new System.EventHandler(this.label6_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(997, 144);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(100, 13);
            this.label7.TabIndex = 25;
            this.label7.Text = "(Файл с данными)";
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(997, 27);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(171, 13);
            this.label8.TabIndex = 26;
            this.label8.Text = "(Файл который надо заполнить)";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1605, 822);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.dataGridView3);
            this.Controls.Add(this.dataGridView2);
            this.Controls.Add(this.dataGridView1);
            this.Controls.Add(this.listBox2);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.textBox3);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.progressBar2);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.textBox2);
            this.Controls.Add(this.textBox1);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.listBox1);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.button1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Form1";
            this.Text = "Программа для заполнения данных паспорта";
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView2)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.dataGridView3)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.ListBox listBox1;
        private System.Windows.Forms.Button button3;
        private System.Windows.Forms.Button button4;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Button button5;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.Windows.Forms.ProgressBar progressBar2;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox3;
        private System.Windows.Forms.ListBox listBox2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.DataGridView dataGridView1;
        private System.Windows.Forms.DataGridView dataGridView2;
        private System.Windows.Forms.DataGridView dataGridView3;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Label label8;
    }
}

