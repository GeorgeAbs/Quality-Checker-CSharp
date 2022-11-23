
namespace Нормоконтроль
{
    partial class Прогресс_проверки_Excel
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
            this.checkUserRulesNumber = new System.Windows.Forms.Label();
            this.commentsForFinal = new System.Windows.Forms.Label();
            this.checkingIsDone = new System.Windows.Forms.Label();
            this.preparingForChecking = new System.Windows.Forms.Label();
            this.checkedRulesNumber = new System.Windows.Forms.Label();
            this.checkedTablesNumber = new System.Windows.Forms.Label();
            this.checkedRules = new System.Windows.Forms.Label();
            this.checkedTables = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.SuspendLayout();
            // 
            // checkUserRulesNumber
            // 
            this.checkUserRulesNumber.AutoSize = true;
            this.checkUserRulesNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkUserRulesNumber.Location = new System.Drawing.Point(335, 228);
            this.checkUserRulesNumber.Name = "checkUserRulesNumber";
            this.checkUserRulesNumber.Size = new System.Drawing.Size(0, 20);
            this.checkUserRulesNumber.TabIndex = 22;
            // 
            // commentsForFinal
            // 
            this.commentsForFinal.AutoSize = true;
            this.commentsForFinal.Location = new System.Drawing.Point(104, 330);
            this.commentsForFinal.Name = "commentsForFinal";
            this.commentsForFinal.Size = new System.Drawing.Size(311, 39);
            this.commentsForFinal.TabIndex = 21;
            this.commentsForFinal.Text = "Проверенный документ с припиской \"после проверки\" \r\n(и log-файл по пользовательск" +
    "им правилам, если были)\r\nсохранены в исходной директории.";
            this.commentsForFinal.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // checkingIsDone
            // 
            this.checkingIsDone.AutoSize = true;
            this.checkingIsDone.Font = new System.Drawing.Font("Segoe UI Semibold", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkingIsDone.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.checkingIsDone.Location = new System.Drawing.Point(159, 292);
            this.checkingIsDone.Name = "checkingIsDone";
            this.checkingIsDone.Size = new System.Drawing.Size(221, 28);
            this.checkingIsDone.TabIndex = 20;
            this.checkingIsDone.Text = "Проверка завершена!\r\n";
            // 
            // preparingForChecking
            // 
            this.preparingForChecking.AutoSize = true;
            this.preparingForChecking.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.preparingForChecking.Location = new System.Drawing.Point(141, 33);
            this.preparingForChecking.Name = "preparingForChecking";
            this.preparingForChecking.Size = new System.Drawing.Size(188, 20);
            this.preparingForChecking.TabIndex = 19;
            this.preparingForChecking.Text = "Подготовка к проверке...";
            // 
            // checkedRulesNumber
            // 
            this.checkedRulesNumber.AutoSize = true;
            this.checkedRulesNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedRulesNumber.Location = new System.Drawing.Point(335, 228);
            this.checkedRulesNumber.Name = "checkedRulesNumber";
            this.checkedRulesNumber.Size = new System.Drawing.Size(0, 20);
            this.checkedRulesNumber.TabIndex = 18;
            // 
            // checkedTablesNumber
            // 
            this.checkedTablesNumber.AutoSize = true;
            this.checkedTablesNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedTablesNumber.Location = new System.Drawing.Point(335, 82);
            this.checkedTablesNumber.Name = "checkedTablesNumber";
            this.checkedTablesNumber.Size = new System.Drawing.Size(0, 20);
            this.checkedTablesNumber.TabIndex = 16;
            // 
            // checkedRules
            // 
            this.checkedRules.AutoSize = true;
            this.checkedRules.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedRules.ForeColor = System.Drawing.Color.Silver;
            this.checkedRules.Location = new System.Drawing.Point(93, 208);
            this.checkedRules.Name = "checkedRules";
            this.checkedRules.Size = new System.Drawing.Size(236, 40);
            this.checkedRules.TabIndex = 14;
            this.checkedRules.Text = "Документ проверен \r\nпользовательскими правилами:";
            this.checkedRules.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // checkedTables
            // 
            this.checkedTables.AutoSize = true;
            this.checkedTables.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedTables.ForeColor = System.Drawing.Color.Silver;
            this.checkedTables.Location = new System.Drawing.Point(183, 82);
            this.checkedTables.Name = "checkedTables";
            this.checkedTables.Size = new System.Drawing.Size(137, 20);
            this.checkedTables.TabIndex = 13;
            this.checkedTables.Text = "Проверено строк:";
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            //this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // Прогресс_проверки_Excel
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(497, 394);
            this.Controls.Add(this.checkUserRulesNumber);
            this.Controls.Add(this.commentsForFinal);
            this.Controls.Add(this.checkingIsDone);
            this.Controls.Add(this.preparingForChecking);
            this.Controls.Add(this.checkedRulesNumber);
            this.Controls.Add(this.checkedTablesNumber);
            this.Controls.Add(this.checkedRules);
            this.Controls.Add(this.checkedTables);
            this.Name = "Прогресс_проверки_Excel";
            this.Text = "Прогресс_проверки_Excel";
            this.Load += new System.EventHandler(this.Прогресс_проверки_Excel_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label checkUserRulesNumber;
        private System.Windows.Forms.Label commentsForFinal;
        private System.Windows.Forms.Label checkingIsDone;
        private System.Windows.Forms.Label preparingForChecking;
        private System.Windows.Forms.Label checkedRulesNumber;
        private System.Windows.Forms.Label checkedTablesNumber;
        private System.Windows.Forms.Label checkedRules;
        private System.Windows.Forms.Label checkedTables;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
    }
}