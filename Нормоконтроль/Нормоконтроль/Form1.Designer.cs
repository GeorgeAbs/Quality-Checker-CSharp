
namespace Нормоконтроль
{
    partial class Form1
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
            this.choseForChech = new System.Windows.Forms.Button();
            this.choseTemplate = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.chosenDocLabel = new System.Windows.Forms.Label();
            this.chosenTemplateLabel = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.startCheck = new System.Windows.Forms.Button();
            this.label7 = new System.Windows.Forms.Label();
            this.word_adjust = new System.Windows.Forms.Button();
            this.chosenRulesFileLabel = new System.Windows.Forms.Label();
            this.label11 = new System.Windows.Forms.Label();
            this.choseRules = new System.Windows.Forms.Button();
            this.panel1 = new System.Windows.Forms.Panel();
            this.panel2 = new System.Windows.Forms.Panel();
            this.label5 = new System.Windows.Forms.Label();
            this.panel3 = new System.Windows.Forms.Panel();
            this.label12 = new System.Windows.Forms.Label();
            this.panel4 = new System.Windows.Forms.Panel();
            this.label3 = new System.Windows.Forms.Label();
            this.panel1.SuspendLayout();
            this.panel2.SuspendLayout();
            this.panel3.SuspendLayout();
            this.panel4.SuspendLayout();
            this.SuspendLayout();
            // 
            // choseForChech
            // 
            this.choseForChech.Location = new System.Drawing.Point(20, 28);
            this.choseForChech.Name = "choseForChech";
            this.choseForChech.Size = new System.Drawing.Size(85, 23);
            this.choseForChech.TabIndex = 0;
            this.choseForChech.Text = "Выбрать";
            this.choseForChech.UseVisualStyleBackColor = true;
            this.choseForChech.Click += new System.EventHandler(this.choseForChech_Click);
            // 
            // choseTemplate
            // 
            this.choseTemplate.Location = new System.Drawing.Point(17, 29);
            this.choseTemplate.Name = "choseTemplate";
            this.choseTemplate.Size = new System.Drawing.Size(85, 23);
            this.choseTemplate.TabIndex = 1;
            this.choseTemplate.Text = "Выбрать";
            this.choseTemplate.UseVisualStyleBackColor = true;
            this.choseTemplate.Click += new System.EventHandler(this.choseTemplate_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(17, 12);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(163, 13);
            this.label1.TabIndex = 2;
            this.label1.Text = "Выбрать файл для проверки:";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(14, 13);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(153, 13);
            this.label2.TabIndex = 3;
            this.label2.Text = "Выбрать документ-шаблон";
            // 
            // chosenDocLabel
            // 
            this.chosenDocLabel.AutoSize = true;
            this.chosenDocLabel.Location = new System.Drawing.Point(136, 33);
            this.chosenDocLabel.Name = "chosenDocLabel";
            this.chosenDocLabel.Size = new System.Drawing.Size(128, 13);
            this.chosenDocLabel.TabIndex = 4;
            this.chosenDocLabel.Text = "Выбранный документ:";
            // 
            // chosenTemplateLabel
            // 
            this.chosenTemplateLabel.AutoSize = true;
            this.chosenTemplateLabel.Location = new System.Drawing.Point(133, 34);
            this.chosenTemplateLabel.Name = "chosenTemplateLabel";
            this.chosenTemplateLabel.Size = new System.Drawing.Size(128, 13);
            this.chosenTemplateLabel.TabIndex = 5;
            this.chosenTemplateLabel.Text = "Выбранный документ:";
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // startCheck
            // 
            this.startCheck.Location = new System.Drawing.Point(20, 19);
            this.startCheck.Name = "startCheck";
            this.startCheck.Size = new System.Drawing.Size(82, 24);
            this.startCheck.TabIndex = 11;
            this.startCheck.Text = "Проверить";
            this.startCheck.UseVisualStyleBackColor = true;
            this.startCheck.Click += new System.EventHandler(this.startCheck_Click);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(122, 19);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(292, 26);
            this.label7.TabIndex = 13;
            this.label7.Text = "После проверки документ сохраняется \r\nв той же директории с припиской \"после пров" +
    "ерки\"";
            // 
            // word_adjust
            // 
            this.word_adjust.Location = new System.Drawing.Point(331, 99);
            this.word_adjust.Name = "word_adjust";
            this.word_adjust.Size = new System.Drawing.Size(125, 39);
            this.word_adjust.TabIndex = 18;
            this.word_adjust.Text = "Настройка правил проверки Word";
            this.word_adjust.UseVisualStyleBackColor = true;
            this.word_adjust.Click += new System.EventHandler(this.word_adjust_Click);
            // 
            // chosenRulesFileLabel
            // 
            this.chosenRulesFileLabel.AutoSize = true;
            this.chosenRulesFileLabel.Location = new System.Drawing.Point(136, 27);
            this.chosenRulesFileLabel.Name = "chosenRulesFileLabel";
            this.chosenRulesFileLabel.Size = new System.Drawing.Size(128, 13);
            this.chosenRulesFileLabel.TabIndex = 25;
            this.chosenRulesFileLabel.Text = "Выбранный документ:";
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(17, 6);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(210, 13);
            this.label11.TabIndex = 24;
            this.label11.Text = "Выбрать файл с правилами проверки";
            // 
            // choseRules
            // 
            this.choseRules.Enabled = false;
            this.choseRules.Location = new System.Drawing.Point(20, 22);
            this.choseRules.Name = "choseRules";
            this.choseRules.Size = new System.Drawing.Size(85, 23);
            this.choseRules.TabIndex = 23;
            this.choseRules.Text = "Выбрать";
            this.choseRules.UseVisualStyleBackColor = true;
            this.choseRules.Click += new System.EventHandler(this.choseRules_Click);
            // 
            // panel1
            // 
            this.panel1.BackColor = System.Drawing.Color.OldLace;
            this.panel1.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel1.Controls.Add(this.choseForChech);
            this.panel1.Controls.Add(this.label1);
            this.panel1.Controls.Add(this.chosenDocLabel);
            this.panel1.Location = new System.Drawing.Point(9, 12);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(586, 74);
            this.panel1.TabIndex = 27;
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.Beige;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label5);
            this.panel2.Controls.Add(this.choseRules);
            this.panel2.Controls.Add(this.label11);
            this.panel2.Controls.Add(this.word_adjust);
            this.panel2.Controls.Add(this.chosenRulesFileLabel);
            this.panel2.Location = new System.Drawing.Point(9, 121);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(586, 151);
            this.panel2.TabIndex = 28;
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(221, 70);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(355, 26);
            this.label5.TabIndex = 27;
            this.label5.Text = "Если файл с правилами не создан/настроен,\r\nто его необходимо создать/настроить с " +
    "помощью кнопки ниже:\r\n";
            this.label5.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // panel3
            // 
            this.panel3.BackColor = System.Drawing.Color.Honeydew;
            this.panel3.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel3.Controls.Add(this.label3);
            this.panel3.Controls.Add(this.label12);
            this.panel3.Controls.Add(this.choseTemplate);
            this.panel3.Controls.Add(this.label2);
            this.panel3.Controls.Add(this.chosenTemplateLabel);
            this.panel3.Location = new System.Drawing.Point(9, 311);
            this.panel3.Name = "panel3";
            this.panel3.Size = new System.Drawing.Size(586, 213);
            this.panel3.TabIndex = 29;
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(3, 88);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(573, 65);
            this.label12.TabIndex = 9;
            this.label12.Text = resources.GetString("label12.Text");
            this.label12.Click += new System.EventHandler(this.label12_Click);
            // 
            // panel4
            // 
            this.panel4.BackColor = System.Drawing.Color.Azure;
            this.panel4.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel4.Controls.Add(this.startCheck);
            this.panel4.Controls.Add(this.label7);
            this.panel4.Location = new System.Drawing.Point(9, 562);
            this.panel4.Name = "panel4";
            this.panel4.Size = new System.Drawing.Size(586, 65);
            this.panel4.TabIndex = 30;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.ForeColor = System.Drawing.Color.Red;
            this.label3.Location = new System.Drawing.Point(3, 165);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(318, 13);
            this.label3.TabIndex = 10;
            this.label3.Text = "Пример абзаца должен включать не менее 200 символов!";
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(604, 645);
            this.Controls.Add(this.panel4);
            this.Controls.Add(this.panel3);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.panel1);
            this.Name = "Form1";
            this.ShowIcon = false;
            this.Tag = "11";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.panel3.ResumeLayout(false);
            this.panel3.PerformLayout();
            this.panel4.ResumeLayout(false);
            this.panel4.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.Button choseForChech;
        private System.Windows.Forms.Button choseTemplate;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label chosenDocLabel;
        private System.Windows.Forms.Label chosenTemplateLabel;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.Button startCheck;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.Button word_adjust;
        private System.Windows.Forms.Label chosenRulesFileLabel;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.Button choseRules;
        private System.Windows.Forms.Panel panel1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Panel panel3;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.Panel panel4;
        private System.Windows.Forms.Label label3;
    }
}

