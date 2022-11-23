using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.IO;

namespace Нормоконтроль
{
    public partial class AddExcelRule : Form
    {
        public AddExcelRule()
        {
            InitializeComponent();
            try
            {
                rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Excel.pathForChosenRures, Encoding.UTF8);
            }
            catch { }
        }
        string[] rulesFile;

        private void closeForm_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void save_Click(object sender, EventArgs e)
        {
            rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Excel.pathForChosenRures, Encoding.UTF8);
            int i = 0;
            while (i < rulesFile.Length - 2 & rulesFile[i] != nameOfRule.Text) //проверка на существующее название правила
            {
                i++;
                if (rulesFile[i] == nameOfRule.Text)
                {
                    Form qq = new Такое_правило_уже_существует();
                    qq.Visible = true;
                    qq.TopMost = true;
                    return;
                }
            }

            foreach (Control emptyControl in this.Controls) //проверка на пустые поля
            {
                if (emptyControl.Text == "" & emptyControl.Enabled == true)
                {
                    Form qq = new Не_все_поля_заполнены();
                    qq.Visible = true;
                    qq.TopMost = true;
                    return;
                }
            }


            File.Delete(Создать_открыть_файл_правил_Excel.pathForChosenRures);
            using (StreamWriter sw = new StreamWriter(Создать_открыть_файл_правил_Excel.pathForChosenRures, true))
            {
                i = 0;
                while (rulesFile[i].Contains("*Rules*") != true)
                {
                    sw.WriteLine(rulesFile[i]);
                    i++;
                }
                while (rulesFile[i].Contains("**Rules**") != true)
                {
                    sw.WriteLine(rulesFile[i]);
                    i++;
                }
                sw.WriteLine("*??Rule??*"); //начало правила
                sw.WriteLine(nameOfRule.Text);
                sw.WriteLine("checkState_YES");
                if (shallInclude.Checked == true)
                {
                    sw.WriteLine("shallInclude");
                    sw.WriteLine(shallIncludeTEXT.Text);
                }
                if (shallNotInclude.Checked == true)
                {
                    sw.WriteLine("shallNotInclude");
                    sw.WriteLine(shallNotIncludeTEXT.Text);
                }
                if (if_A_then_B.Checked == true)
                {
                    sw.WriteLine("if_A_then_B");
                    sw.WriteLine(if_A_then_B_A.Text);
                    sw.WriteLine(if_A_then_B_B.Text);
                }
                if (if_A_not_B.Checked == true)
                {
                    sw.WriteLine("if_A_not_B");
                    sw.WriteLine(if_A_not_B_A.Text);
                    sw.WriteLine(if_A_not_B_B.Text);
                }
                if (if_A_then_A_in_B.Checked == true)
                {
                    sw.WriteLine("if_A_then_A_in_B");
                    sw.WriteLine(if_A_then_A_in_B_A.Text);
                    sw.WriteLine(if_A_then_A_in_B_B.Text);
                }

                sw.WriteLine("**??Rule??**"); //конец правила

                while (i < rulesFile.Length)
                {
                    sw.WriteLine(rulesFile[i]);
                    i++;
                }
            }


            Form q = new Added();
            q.Visible = true;
            q.TopMost = true;
        }

        private void AddExcelRule_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Visible = true;
        }

        private void shallInclude_CheckedChanged(object sender, EventArgs e)
        {
            if (shallInclude.Checked == true)
            {
                shallIncludeTEXT.Enabled = true;
                shallNotIncludeTEXT.Enabled = false;
                if_A_then_B_A.Enabled = false;
                if_A_then_B_B.Enabled = false;
                if_A_not_B_A.Enabled = false;
                if_A_not_B_B.Enabled = false;
                if_A_then_A_in_B_A.Enabled = false;
                if_A_then_A_in_B_B.Enabled = false;
            }
        }

        private void shallNotInclude_CheckedChanged(object sender, EventArgs e)
        {
            if (shallNotInclude.Checked == true)
            {
                shallIncludeTEXT.Enabled = false;
                shallNotIncludeTEXT.Enabled = true;
                if_A_then_B_A.Enabled = false;
                if_A_then_B_B.Enabled = false;
                if_A_not_B_A.Enabled = false;
                if_A_not_B_B.Enabled = false;
                if_A_then_A_in_B_A.Enabled = false;
                if_A_then_A_in_B_B.Enabled = false;
            }
        }

        private void if_A_then_B_CheckedChanged(object sender, EventArgs e)
        {
            if (if_A_then_B.Checked == true)
            {
                shallIncludeTEXT.Enabled = false;
                shallNotIncludeTEXT.Enabled = false;
                if_A_then_B_A.Enabled = true;
                if_A_then_B_B.Enabled = true;
                if_A_not_B_A.Enabled = false;
                if_A_not_B_B.Enabled = false;
                if_A_then_A_in_B_A.Enabled = false;
                if_A_then_A_in_B_B.Enabled = false;
            }
        }

        private void if_A_not_B_CheckedChanged(object sender, EventArgs e)
        {
            if (if_A_not_B.Checked == true)
            {
                shallIncludeTEXT.Enabled = false;
                shallNotIncludeTEXT.Enabled = false;
                if_A_then_B_A.Enabled = false;
                if_A_then_B_B.Enabled = false;
                if_A_not_B_A.Enabled = true;
                if_A_not_B_B.Enabled = true;
                if_A_then_A_in_B_A.Enabled = false;
                if_A_then_A_in_B_B.Enabled = false;
            }
        }

        private void if_A_then_A_in_B_CheckedChanged(object sender, EventArgs e)
        {
            if (if_A_then_A_in_B.Checked == true)
            {
                shallIncludeTEXT.Enabled = false;
                shallNotIncludeTEXT.Enabled = false;
                if_A_then_B_A.Enabled = false;
                if_A_then_B_B.Enabled = false;
                if_A_not_B_A.Enabled = false;
                if_A_not_B_B.Enabled = false;
                if_A_then_A_in_B_A.Enabled = true;
                if_A_then_A_in_B_B.Enabled = true;
            }
        }
    }
}
