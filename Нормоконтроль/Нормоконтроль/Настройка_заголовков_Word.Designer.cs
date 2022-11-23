
namespace Нормоконтроль
{
    partial class Настройка_заголовков_Word
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
            this.llllll = new System.Windows.Forms.Label();
            this.Save = new System.Windows.Forms.Button();
            this.Close = new System.Windows.Forms.Button();
            this.addTitle = new System.Windows.Forms.Button();
            this.remobeTitle = new System.Windows.Forms.Button();
            this.ll = new System.Windows.Forms.Label();
            this.lll = new System.Windows.Forms.Label();
            this.llll = new System.Windows.Forms.Label();
            this.lllll = new System.Windows.Forms.Label();
            this.l = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // llllll
            // 
            this.llllll.AutoSize = true;
            this.llllll.Location = new System.Drawing.Point(12, 9);
            this.llllll.Name = "llllll";
            this.llllll.Size = new System.Drawing.Size(498, 52);
            this.llllll.TabIndex = 1;
            this.llllll.Text = "Добавить уровни заголовков, которые не будут проверяться\r\nна форматирование прави" +
    "лами основного текста\r\n\r\nДля каждого заголовка необходимо определить атрибуты ег" +
    "о форматирования (см. ниже)\r\n";
            this.llllll.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // Save
            // 
            this.Save.Location = new System.Drawing.Point(508, 380);
            this.Save.Name = "Save";
            this.Save.Size = new System.Drawing.Size(75, 23);
            this.Save.TabIndex = 2;
            this.Save.Text = "Cохранить";
            this.Save.UseVisualStyleBackColor = true;
            this.Save.Click += new System.EventHandler(this.Save_Click);
            // 
            // Close
            // 
            this.Close.Location = new System.Drawing.Point(18, 380);
            this.Close.Name = "Close";
            this.Close.Size = new System.Drawing.Size(75, 23);
            this.Close.TabIndex = 3;
            this.Close.Text = "Закрыть";
            this.Close.UseVisualStyleBackColor = true;
            this.Close.Click += new System.EventHandler(this.Close_Click);
            // 
            // addTitle
            // 
            this.addTitle.Location = new System.Drawing.Point(561, 92);
            this.addTitle.Name = "addTitle";
            this.addTitle.Size = new System.Drawing.Size(22, 23);
            this.addTitle.TabIndex = 4;
            this.addTitle.Text = "+";
            this.addTitle.UseVisualStyleBackColor = true;
            this.addTitle.Click += new System.EventHandler(this.addTitle_Click);
            // 
            // remobeTitle
            // 
            this.remobeTitle.Location = new System.Drawing.Point(561, 121);
            this.remobeTitle.Name = "remobeTitle";
            this.remobeTitle.Size = new System.Drawing.Size(22, 23);
            this.remobeTitle.TabIndex = 5;
            this.remobeTitle.Text = "-";
            this.remobeTitle.UseVisualStyleBackColor = true;
            this.remobeTitle.Click += new System.EventHandler(this.remobeTitle_Click);
            // 
            // ll
            // 
            this.ll.AutoSize = true;
            this.ll.Location = new System.Drawing.Point(117, 90);
            this.ll.Name = "ll";
            this.ll.Size = new System.Drawing.Size(90, 13);
            this.ll.TabIndex = 6;
            this.ll.Text = "Размер шрифта";
            // 
            // lll
            // 
            this.lll.AutoSize = true;
            this.lll.Location = new System.Drawing.Point(235, 90);
            this.lll.Name = "lll";
            this.lll.Size = new System.Drawing.Size(53, 13);
            this.lll.TabIndex = 7;
            this.lll.Text = "Жирный";
            // 
            // llll
            // 
            this.llll.AutoSize = true;
            this.llll.Location = new System.Drawing.Point(334, 90);
            this.llll.Name = "llll";
            this.llll.Size = new System.Drawing.Size(45, 13);
            this.llll.TabIndex = 8;
            this.llll.Text = "Курсив";
            // 
            // lllll
            // 
            this.lllll.AutoSize = true;
            this.lllll.Location = new System.Drawing.Point(404, 90);
            this.lllll.Name = "lllll";
            this.lllll.Size = new System.Drawing.Size(87, 13);
            this.lllll.TabIndex = 9;
            this.lllll.Text = "Выравнивание";
            // 
            // l
            // 
            this.l.AutoSize = true;
            this.l.Location = new System.Drawing.Point(15, 92);
            this.l.Name = "l";
            this.l.Size = new System.Drawing.Size(52, 13);
            this.l.TabIndex = 17;
            this.l.Text = "Уровень";
            // 
            // Настройка_заголовков_Word
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(598, 411);
            this.Controls.Add(this.l);
            this.Controls.Add(this.llll);
            this.Controls.Add(this.lllll);
            this.Controls.Add(this.lll);
            this.Controls.Add(this.ll);
            this.Controls.Add(this.remobeTitle);
            this.Controls.Add(this.addTitle);
            this.Controls.Add(this.Close);
            this.Controls.Add(this.Save);
            this.Controls.Add(this.llllll);
            this.Name = "Настройка_заголовков_Word";
            this.ShowIcon = false;
            this.Text = "Настройка заголовков Word";
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Label llllll;
        private System.Windows.Forms.Button Save;
        private System.Windows.Forms.Button Close;
        private System.Windows.Forms.Button addTitle;
        private System.Windows.Forms.Button remobeTitle;
        private System.Windows.Forms.Label ll;
        private System.Windows.Forms.Label lll;
        private System.Windows.Forms.Label llll;
        private System.Windows.Forms.Label lllll;
        private System.Windows.Forms.Label l;
    }
}