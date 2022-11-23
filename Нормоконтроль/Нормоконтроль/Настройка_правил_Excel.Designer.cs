
namespace Нормоконтроль
{
    partial class Настройка_правил_Excel
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
            this.label3 = new System.Windows.Forms.Label();
            this.ruleDescription = new System.Windows.Forms.RichTextBox();
            this.closeWindow = new System.Windows.Forms.Button();
            this.deleteRule = new System.Windows.Forms.Button();
            this.addRule = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.save = new System.Windows.Forms.Button();
            this.listOfRules = new System.Windows.Forms.CheckedListBox();
            this.SuspendLayout();
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(654, 13);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(108, 13);
            this.label3.TabIndex = 47;
            this.label3.Text = "Описание правила";
            // 
            // ruleDescription
            // 
            this.ruleDescription.BackColor = System.Drawing.Color.AliceBlue;
            this.ruleDescription.BorderStyle = System.Windows.Forms.BorderStyle.FixedSingle;
            this.ruleDescription.Location = new System.Drawing.Point(521, 29);
            this.ruleDescription.Name = "ruleDescription";
            this.ruleDescription.Size = new System.Drawing.Size(362, 98);
            this.ruleDescription.TabIndex = 46;
            this.ruleDescription.Text = "";
            // 
            // closeWindow
            // 
            this.closeWindow.Location = new System.Drawing.Point(8, 639);
            this.closeWindow.Name = "closeWindow";
            this.closeWindow.Size = new System.Drawing.Size(75, 23);
            this.closeWindow.TabIndex = 45;
            this.closeWindow.Text = "Закрыть";
            this.closeWindow.UseVisualStyleBackColor = true;
            this.closeWindow.Click += new System.EventHandler(this.closeWindow_Click);
            // 
            // deleteRule
            // 
            this.deleteRule.Location = new System.Drawing.Point(8, 72);
            this.deleteRule.Name = "deleteRule";
            this.deleteRule.Size = new System.Drawing.Size(85, 55);
            this.deleteRule.TabIndex = 44;
            this.deleteRule.Text = "Удалить выбранное правило";
            this.deleteRule.UseVisualStyleBackColor = true;
            this.deleteRule.Click += new System.EventHandler(this.deleteRule_Click);
            // 
            // addRule
            // 
            this.addRule.Location = new System.Drawing.Point(8, 29);
            this.addRule.Name = "addRule";
            this.addRule.Size = new System.Drawing.Size(85, 37);
            this.addRule.TabIndex = 43;
            this.addRule.Text = "Создать правило";
            this.addRule.UseVisualStyleBackColor = true;
            this.addRule.Click += new System.EventHandler(this.addRule_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(232, 13);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(88, 13);
            this.label1.TabIndex = 42;
            this.label1.Text = "Список правил";
            // 
            // save
            // 
            this.save.Location = new System.Drawing.Point(805, 624);
            this.save.Name = "save";
            this.save.Size = new System.Drawing.Size(81, 38);
            this.save.TabIndex = 41;
            this.save.Text = "Сохранить\r\nнастройки";
            this.save.UseVisualStyleBackColor = true;
            this.save.Click += new System.EventHandler(this.save_Click);
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
            this.listOfRules.Location = new System.Drawing.Point(99, 29);
            this.listOfRules.Name = "listOfRules";
            this.listOfRules.Size = new System.Drawing.Size(362, 582);
            this.listOfRules.TabIndex = 40;
            this.listOfRules.SelectedValueChanged += new System.EventHandler(this.listOfRules_SelectedValueChanged);
            // 
            // Настройка_правил_Excel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(901, 674);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.ruleDescription);
            this.Controls.Add(this.closeWindow);
            this.Controls.Add(this.deleteRule);
            this.Controls.Add(this.addRule);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.save);
            this.Controls.Add(this.listOfRules);
            this.Name = "Настройка_правил_Excel";
            this.Text = "Настройка_правил_Excel";
            this.Load += new System.EventHandler(this.Настройка_правил_Excel_Load);
            this.VisibleChanged += new System.EventHandler(this.Настройка_правил_Excel_VisibleChanged);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.RichTextBox ruleDescription;
        private System.Windows.Forms.Button closeWindow;
        private System.Windows.Forms.Button deleteRule;
        private System.Windows.Forms.Button addRule;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button save;
        private System.Windows.Forms.CheckedListBox listOfRules;
    }
}