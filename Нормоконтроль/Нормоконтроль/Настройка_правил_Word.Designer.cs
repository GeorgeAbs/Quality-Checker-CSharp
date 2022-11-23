
namespace Нормоконтроль
{
    partial class Настройка_правил_Word
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
            this.listOfRules = new System.Windows.Forms.CheckedListBox();
            this.save = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.saveFileDialog1 = new System.Windows.Forms.SaveFileDialog();
            this.addRule = new System.Windows.Forms.Button();
            this.deleteRule = new System.Windows.Forms.Button();
            this.label2 = new System.Windows.Forms.Label();
            this.button1 = new System.Windows.Forms.Button();
            this.panel2 = new System.Windows.Forms.Panel();
            this.closeWindow = new System.Windows.Forms.Button();
            this.ruleDescription = new System.Windows.Forms.RichTextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.panel2.SuspendLayout();
            this.SuspendLayout();
            // 
            // listOfRules
            // 
            this.listOfRules.BackColor = System.Drawing.Color.AliceBlue;
            this.listOfRules.FormattingEnabled = true;
            this.listOfRules.Items.AddRange(new object[] {
            "Проверка на размер шрифта",
            "Проверка на стиль шрифта",
            "Проверка на курсив текста",
            "Проверка на жирность текста",
            "Проверка на выравнивание текста",
            "Проверка на отступы слева/справа/первой строки",
            "Проверка на интервалы до/после",
            "Проверка на двойные пробелы",
            "Проверка на двойные знаки препинания",
            "Проверка: цвет текста везде черный ",
            "Проверка: цвет фона текста везде отсутствует "});
            this.listOfRules.Location = new System.Drawing.Point(103, 22);
            this.listOfRules.Name = "listOfRules";
            this.listOfRules.Size = new System.Drawing.Size(362, 582);
            this.listOfRules.TabIndex = 19;
            this.listOfRules.SelectedIndexChanged += new System.EventHandler(this.listOfRules_SelectedIndexChanged);
            this.listOfRules.SelectedValueChanged += new System.EventHandler(this.listOfRules_SelectedValueChanged);
            // 
            // save
            // 
            this.save.Location = new System.Drawing.Point(806, 617);
            this.save.Name = "save";
            this.save.Size = new System.Drawing.Size(81, 38);
            this.save.TabIndex = 21;
            this.save.Text = "Сохранить\r\nнастройки";
            this.save.UseVisualStyleBackColor = true;
            this.save.Click += new System.EventHandler(this.save_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(236, 6);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 24;
            this.label1.Text = "Список правил";
            this.label1.Click += new System.EventHandler(this.label1_Click);
            // 
            // openFileDialog1
            // 
            this.openFileDialog1.FileName = "openFileDialog1";
            // 
            // addRule
            // 
            this.addRule.Location = new System.Drawing.Point(12, 22);
            this.addRule.Name = "addRule";
            this.addRule.Size = new System.Drawing.Size(85, 37);
            this.addRule.TabIndex = 26;
            this.addRule.Text = "Создать правило";
            this.addRule.UseVisualStyleBackColor = true;
            this.addRule.Click += new System.EventHandler(this.addRule_Click);
            // 
            // deleteRule
            // 
            this.deleteRule.Location = new System.Drawing.Point(12, 65);
            this.deleteRule.Name = "deleteRule";
            this.deleteRule.Size = new System.Drawing.Size(85, 55);
            this.deleteRule.TabIndex = 27;
            this.deleteRule.Text = "Удалить выбранное правило";
            this.deleteRule.UseVisualStyleBackColor = true;
            this.deleteRule.Click += new System.EventHandler(this.deleteRule_Click);
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(36, 9);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(289, 26);
            this.label2.TabIndex = 33;
            this.label2.Text = "Уровни заголовков, которые не будут проверяться\r\nна форматирование параметрами ос" +
    "новного текста";
            this.label2.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // button1
            // 
            this.button1.Location = new System.Drawing.Point(135, 64);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(75, 23);
            this.button1.TabIndex = 34;
            this.button1.Text = "Настроить";
            this.button1.UseVisualStyleBackColor = true;
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // panel2
            // 
            this.panel2.BackColor = System.Drawing.Color.SeaShell;
            this.panel2.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.panel2.Controls.Add(this.label2);
            this.panel2.Controls.Add(this.button1);
            this.panel2.Location = new System.Drawing.Point(525, 505);
            this.panel2.Name = "panel2";
            this.panel2.Size = new System.Drawing.Size(362, 99);
            this.panel2.TabIndex = 35;
            // 
            // closeWindow
            // 
            this.closeWindow.Location = new System.Drawing.Point(12, 632);
            this.closeWindow.Name = "closeWindow";
            this.closeWindow.Size = new System.Drawing.Size(75, 23);
            this.closeWindow.TabIndex = 37;
            this.closeWindow.Text = "Закрыть";
            this.closeWindow.UseVisualStyleBackColor = true;
            this.closeWindow.Click += new System.EventHandler(this.closeWindow_Click);
            // 
            // ruleDescription
            // 
            this.ruleDescription.BackColor = System.Drawing.Color.AliceBlue;
            this.ruleDescription.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ruleDescription.Location = new System.Drawing.Point(525, 22);
            this.ruleDescription.Name = "ruleDescription";
            this.ruleDescription.Size = new System.Drawing.Size(362, 466);
            this.ruleDescription.TabIndex = 38;
            this.ruleDescription.Text = "";
            this.ruleDescription.TextChanged += new System.EventHandler(this.ruleDescription_TextChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(658, 6);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 13);
            this.label3.TabIndex = 39;
            this.label3.Text = "Описание правила";
            this.label3.Click += new System.EventHandler(this.label3_Click);
            // 
            // Настройка_правил_Word
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.BackColor = System.Drawing.Color.White;
            this.ClientSize = new System.Drawing.Size(902, 663);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ruleDescription);
            this.Controls.Add(this.closeWindow);
            this.Controls.Add(this.panel2);
            this.Controls.Add(this.deleteRule);
            this.Controls.Add(this.addRule);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.save);
            this.Controls.Add(this.listOfRules);
            this.Name = "Настройка_правил_Word";
            this.ShowIcon = false;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Настройка правил Word";
            this.Load += new System.EventHandler(this.Настройка_правил_Word_Load);
            this.VisibleChanged += new System.EventHandler(this.Настройка_правил_Word_VisibleChanged);
            this.panel2.ResumeLayout(false);
            this.panel2.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.CheckedListBox listOfRules;
        private System.Windows.Forms.Button save;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private System.Windows.Forms.SaveFileDialog saveFileDialog1;
        private System.Windows.Forms.Button addRule;
        private System.Windows.Forms.Button deleteRule;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Button button1;
        private System.Windows.Forms.Panel panel2;
        private System.Windows.Forms.Button closeWindow;
        private System.Windows.Forms.RichTextBox ruleDescription;
        private System.Windows.Forms.Label label3;
    }
}