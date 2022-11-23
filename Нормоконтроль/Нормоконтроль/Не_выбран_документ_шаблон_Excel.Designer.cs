
namespace Нормоконтроль
{
    partial class Не_выбран_документ_шаблон_Excel
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Не_выбран_документ_шаблон_Excel));
            this.yes = new System.Windows.Forms.Button();
            this.canc = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // yes
            // 
            this.yes.Location = new System.Drawing.Point(282, 159);
            this.yes.Name = "yes";
            this.yes.Size = new System.Drawing.Size(75, 23);
            this.yes.TabIndex = 5;
            this.yes.Text = "Да";
            this.yes.UseVisualStyleBackColor = true;
            this.yes.Click += new System.EventHandler(this.yes_Click);
            // 
            // canc
            // 
            this.canc.Location = new System.Drawing.Point(15, 159);
            this.canc.Name = "canc";
            this.canc.Size = new System.Drawing.Size(75, 23);
            this.canc.TabIndex = 4;
            this.canc.Text = "Отмена";
            this.canc.UseVisualStyleBackColor = true;
            this.canc.Click += new System.EventHandler(this.canc_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(12, 9);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(345, 117);
            this.label1.TabIndex = 3;
            this.label1.Text = resources.GetString("label1.Text");
            // 
            // Не_выбран_документ_шаблон_Excel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(367, 197);
            this.Controls.Add(this.yes);
            this.Controls.Add(this.canc);
            this.Controls.Add(this.label1);
            this.Name = "Не_выбран_документ_шаблон_Excel";
            this.Text = "Не_выбран_документ_шаблон_Excel";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Не_выбран_документ_шаблон_Excel_FormClosing);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Button yes;
        private System.Windows.Forms.Button canc;
        private System.Windows.Forms.Label label1;
    }
}