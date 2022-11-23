
namespace Нормоконтроль
{
    partial class Прогресс_проверки
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
            this.checkedTables = new System.Windows.Forms.Label();
            this.checkedFormating = new System.Windows.Forms.Label();
            this.checkedRules = new System.Windows.Forms.Label();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.checkedTablesNumber = new System.Windows.Forms.Label();
            this.checkedFormatingNumber = new System.Windows.Forms.Label();
            this.checkedRulesNumber = new System.Windows.Forms.Label();
            this.preparingForChecking = new System.Windows.Forms.Label();
            this.checkingIsDone = new System.Windows.Forms.Label();
            this.commentsForFinal = new System.Windows.Forms.Label();
            this.checkUserRulesNumber = new System.Windows.Forms.Label();
            this.SuspendLayout();
            // 
            // checkedTables
            // 
            this.checkedTables.AutoSize = true;
            this.checkedTables.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedTables.ForeColor = System.Drawing.Color.Silver;
            this.checkedTables.Location = new System.Drawing.Point(166, 74);
            this.checkedTables.Name = "checkedTables";
            this.checkedTables.Size = new System.Drawing.Size(146, 20);
            this.checkedTables.TabIndex = 0;
            this.checkedTables.Text = "Проверено таблиц:";
            // 
            // checkedFormating
            // 
            this.checkedFormating.AutoSize = true;
            this.checkedFormating.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedFormating.ForeColor = System.Drawing.Color.Silver;
            this.checkedFormating.Location = new System.Drawing.Point(10, 126);
            this.checkedFormating.Name = "checkedFormating";
            this.checkedFormating.Size = new System.Drawing.Size(302, 40);
            this.checkedFormating.TabIndex = 2;
            this.checkedFormating.Text = "Проверено абзацев на форматирование \r\n(вне таблиц):";
            this.checkedFormating.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // checkedRules
            // 
            this.checkedRules.AutoSize = true;
            this.checkedRules.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedRules.ForeColor = System.Drawing.Color.Silver;
            this.checkedRules.Location = new System.Drawing.Point(76, 200);
            this.checkedRules.Name = "checkedRules";
            this.checkedRules.Size = new System.Drawing.Size(236, 40);
            this.checkedRules.TabIndex = 1;
            this.checkedRules.Text = "Документ проверен \r\nпользовательскими правилами:";
            this.checkedRules.TextAlign = System.Drawing.ContentAlignment.TopRight;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.WorkerSupportsCancellation = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker_RunWorkerCompleted);
            // 
            // checkedTablesNumber
            // 
            this.checkedTablesNumber.AutoSize = true;
            this.checkedTablesNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedTablesNumber.Location = new System.Drawing.Point(318, 74);
            this.checkedTablesNumber.Name = "checkedTablesNumber";
            this.checkedTablesNumber.Size = new System.Drawing.Size(0, 20);
            this.checkedTablesNumber.TabIndex = 4;
            // 
            // checkedFormatingNumber
            // 
            this.checkedFormatingNumber.AutoSize = true;
            this.checkedFormatingNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedFormatingNumber.Location = new System.Drawing.Point(318, 146);
            this.checkedFormatingNumber.Name = "checkedFormatingNumber";
            this.checkedFormatingNumber.Size = new System.Drawing.Size(0, 20);
            this.checkedFormatingNumber.TabIndex = 5;
            // 
            // checkedRulesNumber
            // 
            this.checkedRulesNumber.AutoSize = true;
            this.checkedRulesNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkedRulesNumber.Location = new System.Drawing.Point(318, 220);
            this.checkedRulesNumber.Name = "checkedRulesNumber";
            this.checkedRulesNumber.Size = new System.Drawing.Size(0, 20);
            this.checkedRulesNumber.TabIndex = 6;
            // 
            // preparingForChecking
            // 
            this.preparingForChecking.AutoSize = true;
            this.preparingForChecking.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.preparingForChecking.Location = new System.Drawing.Point(124, 25);
            this.preparingForChecking.Name = "preparingForChecking";
            this.preparingForChecking.Size = new System.Drawing.Size(188, 20);
            this.preparingForChecking.TabIndex = 8;
            this.preparingForChecking.Text = "Подготовка к проверке...";
            // 
            // checkingIsDone
            // 
            this.checkingIsDone.AutoSize = true;
            this.checkingIsDone.Font = new System.Drawing.Font("Segoe UI Semibold", 15F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkingIsDone.ForeColor = System.Drawing.Color.FromArgb(((int)(((byte)(0)))), ((int)(((byte)(192)))), ((int)(((byte)(0)))));
            this.checkingIsDone.Location = new System.Drawing.Point(142, 284);
            this.checkingIsDone.Name = "checkingIsDone";
            this.checkingIsDone.Size = new System.Drawing.Size(221, 28);
            this.checkingIsDone.TabIndex = 9;
            this.checkingIsDone.Text = "Проверка завершена!\r\n";
            // 
            // commentsForFinal
            // 
            this.commentsForFinal.AutoSize = true;
            this.commentsForFinal.Location = new System.Drawing.Point(87, 322);
            this.commentsForFinal.Name = "commentsForFinal";
            this.commentsForFinal.Size = new System.Drawing.Size(311, 39);
            this.commentsForFinal.TabIndex = 10;
            this.commentsForFinal.Text = "Проверенный документ с припиской \"после проверки\" \r\n(и log-файл по пользовательск" +
    "им правилам, если были)\r\nсохранены в исходной директории.";
            this.commentsForFinal.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // checkUserRulesNumber
            // 
            this.checkUserRulesNumber.AutoSize = true;
            this.checkUserRulesNumber.Font = new System.Drawing.Font("Segoe UI Semibold", 11F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.checkUserRulesNumber.Location = new System.Drawing.Point(318, 220);
            this.checkUserRulesNumber.Name = "checkUserRulesNumber";
            this.checkUserRulesNumber.Size = new System.Drawing.Size(0, 20);
            this.checkUserRulesNumber.TabIndex = 12;
            // 
            // Прогресс_проверки
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(505, 395);
            this.Controls.Add(this.checkUserRulesNumber);
            this.Controls.Add(this.commentsForFinal);
            this.Controls.Add(this.checkingIsDone);
            this.Controls.Add(this.preparingForChecking);
            this.Controls.Add(this.checkedRulesNumber);
            this.Controls.Add(this.checkedFormatingNumber);
            this.Controls.Add(this.checkedTablesNumber);
            this.Controls.Add(this.checkedRules);
            this.Controls.Add(this.checkedFormating);
            this.Controls.Add(this.checkedTables);
            this.Name = "Прогресс_проверки";
            this.ShowIcon = false;
            this.Text = "Прогресс проверки";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Прогресс_проверки_FormClosing);
            this.Load += new System.EventHandler(this.Прогресс_проверки_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label checkedTables;
        private System.Windows.Forms.Label checkedFormating;
        private System.Windows.Forms.Label checkedRules;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label checkedTablesNumber;
        private System.Windows.Forms.Label checkedFormatingNumber;
        private System.Windows.Forms.Label checkedRulesNumber;
        private System.Windows.Forms.Label preparingForChecking;
        private System.Windows.Forms.Label checkingIsDone;
        private System.Windows.Forms.Label commentsForFinal;
        private System.Windows.Forms.Label checkUserRulesNumber;
    }
}