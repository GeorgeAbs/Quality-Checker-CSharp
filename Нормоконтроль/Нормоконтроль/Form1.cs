using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using Word = Microsoft.Office.Interop.Word;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

namespace Нормоконтроль
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
            startCheck.Enabled = false;
        }
        public static string fileForCheck = "";
        public static string templateForCheck = "";
        public static string rulesFile = "";
        private void Form1_Load(object sender, EventArgs e)
        {

        }

        private void choseForChech_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "документы Word|*.doc*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileForCheck = openFileDialog1.FileName;
                chosenDocLabel.Text = "Выбранный файл:   " + fileForCheck.Substring(fileForCheck.LastIndexOf("\\") + 1);
                choseRules.Enabled = true;
            }
        }

        private void choseTemplate_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "документы Word|*.doc*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                templateForCheck = openFileDialog1.FileName;
                chosenTemplateLabel.Text = "Выбранный файл:   " + templateForCheck.Substring(templateForCheck.LastIndexOf("\\") + 1);
            }
        }


        private void startCheck_Click(object sender, EventArgs e)
        {
            if (templateForCheck == "")
            {
                Form qq = new Не_выбран_документ_шаблон();
                qq.Owner = this;
                qq.Visible = true;
                qq.TopMost = true;
                this.Visible = false;
            }
            else
            {
                Form qq = new Прогресс_проверки();
                qq.Owner = this;
                qq.Visible = true;
                qq.TopMost = true;
                this.Visible = false;
            }
        }
        
        private void word_adjust_Click(object sender, EventArgs e)
        {
            Form q = new Создать_открыть_файл_правил_Word();
            q.Visible = true;
            q.TopLevel = true;
        }

        private void label14_Click(object sender, EventArgs e)
        {

        }

        private void Form1_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Show();
        }

        private void label12_Click(object sender, EventArgs e)
        {

        }

        private void choseRules_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Файл с правилами|*.W_rules";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                rulesFile = openFileDialog1.FileName;
                chosenRulesFileLabel.Text = "Выбранный файл:   " + rulesFile.Substring(rulesFile.LastIndexOf("\\") + 1);
                startCheck.Enabled = true;
            }
        }
    }
}
