
namespace Нормоконтроль
{
    partial class Выбор_типа_проверяемого_документа_стартовая_страница
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
            this.label1 = new System.Windows.Forms.Label();
            this.word_ = new System.Windows.Forms.Button();
            this.excel_ = new System.Windows.Forms.Button();
            this.acad_ = new System.Windows.Forms.Button();
            this.pictureBox1 = new System.Windows.Forms.PictureBox();
            this.label3 = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
            this.SuspendLayout();
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(79, 20);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(216, 13);
            this.label1.TabIndex = 0;
            this.label1.Text = "Выбрать тип проверяемого документа";
            // 
            // word_
            // 
            this.word_.Location = new System.Drawing.Point(27, 93);
            this.word_.Name = "word_";
            this.word_.Size = new System.Drawing.Size(75, 23);
            this.word_.TabIndex = 1;
            this.word_.Text = "Word";
            this.word_.UseVisualStyleBackColor = true;
            this.word_.Click += new System.EventHandler(this.word__Click);
            // 
            // excel_
            // 
            this.excel_.Location = new System.Drawing.Point(151, 93);
            this.excel_.Name = "excel_";
            this.excel_.Size = new System.Drawing.Size(75, 23);
            this.excel_.TabIndex = 2;
            this.excel_.Text = "Excel";
            this.excel_.UseVisualStyleBackColor = true;
            this.excel_.Click += new System.EventHandler(this.excel__Click);
            // 
            // acad_
            // 
            this.acad_.Enabled = false;
            this.acad_.Location = new System.Drawing.Point(282, 93);
            this.acad_.Name = "acad_";
            this.acad_.Size = new System.Drawing.Size(75, 23);
            this.acad_.TabIndex = 3;
            this.acad_.Text = "ACAD";
            this.acad_.UseVisualStyleBackColor = true;
            // 
            // pictureBox1
            // 
            this.pictureBox1.Image = global::Нормоконтроль.Properties.Resources.varik3;
            this.pictureBox1.Location = new System.Drawing.Point(174, 162);
            this.pictureBox1.Name = "pictureBox1";
            this.pictureBox1.Size = new System.Drawing.Size(224, 42);
            this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom;
            this.pictureBox1.TabIndex = 4;
            this.pictureBox1.TabStop = false;
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(276, 119);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(86, 13);
            this.label3.TabIndex = 6;
            this.label3.Text = "(в разработке)";
            // 
            // Выбор_типа_проверяемого_документа_стартовая_страница
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(388, 204);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.pictureBox1);
            this.Controls.Add(this.acad_);
            this.Controls.Add(this.excel_);
            this.Controls.Add(this.word_);
            this.Controls.Add(this.label1);
            this.Name = "Выбор_типа_проверяемого_документа_стартовая_страница";
            this.ShowIcon = false;
            this.Text = "Выбор типа документа";
            this.Load += new System.EventHandler(this.Выбор_типа_проверяемого_документа_стартовая_страница_Load);
            ((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Button word_;
        private System.Windows.Forms.Button excel_;
        private System.Windows.Forms.Button acad_;
        private System.Windows.Forms.PictureBox pictureBox1;
        private System.Windows.Forms.Label label3;
    }
}