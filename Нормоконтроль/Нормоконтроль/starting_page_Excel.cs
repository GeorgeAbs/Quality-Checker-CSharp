using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace Нормоконтроль
{
    public partial class starting_page_Excel : Form
    {
        public starting_page_Excel()
        {
            InitializeComponent();
            startCheck.Enabled = false;
        }
        public static string fileForCheck = "";
        public static string templateForCheck = "";
        public static string rulesFile = "";

        private void choseForChech_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "документы Excel|*.xls*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                fileForCheck = openFileDialog1.FileName;
                chosenDocLabel.Text = "Выбранный файл:   " + fileForCheck.Substring(fileForCheck.LastIndexOf("\\") + 1);
                choseRules.Enabled = true;
            }
        }

        private void choseRules_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Файл с правилами Excel|*.E_rules";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                rulesFile = openFileDialog1.FileName;
                chosenRulesFileLabel.Text = "Выбранный файл:   " + rulesFile.Substring(rulesFile.LastIndexOf("\\") + 1);
                startCheck.Enabled = true;
            }
        }

        private void excel_adjust_Click_1(object sender, EventArgs e)
        {
            Form q = new Создать_открыть_файл_правил_Excel();
            q.Visible = true;
            q.TopLevel = true;
        }

        private void choseTemplate_Click_1(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "документы Excel|*.xls*";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                templateForCheck = openFileDialog1.FileName;
                chosenTemplateLabel.Text = "Выбранный файл:   " + templateForCheck.Substring(templateForCheck.LastIndexOf("\\") + 1);
            }
        }

        private void startCheck_Click_1(object sender, EventArgs e)
        {
            if (templateForCheck == "")
            {
                Form qq = new Не_выбран_документ_шаблон_Excel();
                qq.Owner = this;
                qq.Visible = true;
                qq.TopMost = true;
                this.Visible = false;
            }
            else
            {
                Form qq = new Прогресс_проверки_Excel();
                qq.Owner = this;
                qq.Visible = true;
                qq.TopMost = true;
                this.Visible = false;
            }
        }

        private void starting_page_Excel_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Show();
        }
    }
}
