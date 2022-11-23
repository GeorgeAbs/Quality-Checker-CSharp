
namespace Нормоконтроль
{
    partial class AddWordRule
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(AddWordRule));
            this.nameOfRule = new System.Windows.Forms.TextBox();
            this.l1 = new System.Windows.Forms.Label();
            this.l2 = new System.Windows.Forms.Label();
            this.save = new System.Windows.Forms.Button();
            this.B = new System.Windows.Forms.RichTextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.label1 = new System.Windows.Forms.Label();
            this.A = new System.Windows.Forms.TextBox();
            this.набор_правил = new System.Windows.Forms.ListBox();
            this.checkRunningTitle = new System.Windows.Forms.CheckBox();
            this.label3 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // nameOfRule
            // 
            this.nameOfRule.Location = new System.Drawing.Point(16, 28);
            this.nameOfRule.Name = "nameOfRule";
            this.nameOfRule.Size = new System.Drawing.Size(275, 22);
            this.nameOfRule.TabIndex = 7;
            this.nameOfRule.Text = "Мое правило";
            // 
            // l1
            // 
            this.l1.AutoSize = true;
            this.l1.Location = new System.Drawing.Point(100, 9);
            this.l1.Name = "l1";
            this.l1.Size = new System.Drawing.Size(106, 13);
            this.l1.TabIndex = 8;
            this.l1.Text = "Название правила";
            // 
            // l2
            // 
            this.l2.AutoSize = true;
            this.l2.Location = new System.Drawing.Point(83, 94);
            this.l2.Name = "l2";
            this.l2.Size = new System.Drawing.Size(123, 13);
            this.l2.TabIndex = 9;
            this.l2.Text = "Выбрать тип правила";
            // 
            // save
            // 
            this.save.Location = new System.Drawing.Point(929, 638);
            this.save.Name = "save";
            this.save.Size = new System.Drawing.Size(75, 23);
            this.save.TabIndex = 23;
            this.save.Text = "Добавить";
            this.save.UseVisualStyleBackColor = true;
            this.save.Click += new System.EventHandler(this.save_Click);
            // 
            // B
            // 
            this.B.Location = new System.Drawing.Point(634, 110);
            this.B.Name = "B";
            this.B.Size = new System.Drawing.Size(370, 484);
            this.B.TabIndex = 33;
            this.B.Text = resources.GetString("B.Text");
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(823, 94);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(14, 13);
            this.label2.TabIndex = 32;
            this.label2.Text = "Б";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(467, 94);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(14, 13);
            this.label1.TabIndex = 31;
            this.label1.Text = "А";
            // 
            // A
            // 
            this.A.Location = new System.Drawing.Point(297, 110);
            this.A.Name = "A";
            this.A.Size = new System.Drawing.Size(330, 22);
            this.A.TabIndex = 30;
            // 
            // набор_правил
            // 
            this.набор_правил.FormattingEnabled = true;
            this.набор_правил.HorizontalScrollbar = true;
            this.набор_правил.Items.AddRange(new object[] {
            "1. Документ должен включать А",
            "2. Документ не должен включать А",
            "3. Если есть А, то должно быть Б",
            "4. Если есть А, то не должно быть Б",
            "5. Если есть А, то А должно быть в Б",
            "6. Если есть А, то должно быть Б (А - многострочное)"});
            this.набор_правил.Location = new System.Drawing.Point(12, 110);
            this.набор_правил.Name = "набор_правил";
            this.набор_правил.Size = new System.Drawing.Size(279, 498);
            this.набор_правил.TabIndex = 34;
            this.набор_правил.SelectedValueChanged += new System.EventHandler(this.набор_правил_SelectedValueChanged);
            // 
            // checkRunningTitle
            // 
            this.checkRunningTitle.AutoSize = true;
            this.checkRunningTitle.Font = new System.Drawing.Font("Segoe UI Semibold", 10F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkRunningTitle.Location = new System.Drawing.Point(634, 600);
            this.checkRunningTitle.Name = "checkRunningTitle";
            this.checkRunningTitle.Size = new System.Drawing.Size(200, 23);
            this.checkRunningTitle.TabIndex = 35;
            this.checkRunningTitle.Text = "Проверять в колонтитулах";
            this.checkRunningTitle.UseVisualStyleBackColor = true;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(502, 625);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(430, 39);
            this.label3.TabIndex = 36;
            this.label3.Text = "Внимание! \r\nПроверка в колонтитулах может значительно замедлить проверку правилом" +
    ". \r\nРекомендуется использовать только в случае необходимости.";
            this.label3.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // AddWordRule
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1018, 673);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.checkRunningTitle);
            this.Controls.Add(this.набор_правил);
            this.Controls.Add(this.B);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.A);
            this.Controls.Add(this.save);
            this.Controls.Add(this.l2);
            this.Controls.Add(this.l1);
            this.Controls.Add(this.nameOfRule);
            this.Name = "AddWordRule";
            this.ShowIcon = false;
            this.Text = "Новое правило Word";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.AddWordRule_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.TextBox nameOfRule;
        private System.Windows.Forms.Label l1;
        private System.Windows.Forms.Label l2;
        private System.Windows.Forms.Button save;
        private System.Windows.Forms.RichTextBox B;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.TextBox A;
        private System.Windows.Forms.ListBox набор_правил;
        private System.Windows.Forms.CheckBox checkRunningTitle;
        private System.Windows.Forms.Label label3;
    }
}