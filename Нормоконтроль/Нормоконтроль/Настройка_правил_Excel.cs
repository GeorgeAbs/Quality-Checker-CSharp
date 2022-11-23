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
    public partial class Настройка_правил_Excel : Form
    {
        public Настройка_правил_Excel()
        {
            InitializeComponent();
            try
            {
                rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Excel.pathForChosenRures, Encoding.UTF8);
            }
            catch { }
            deleteRule.Enabled = false;
        }
        string[] rulesFile;

        private void addRule_Click(object sender, EventArgs e)
        {
            Form q = new AddExcelRule();
            q.Owner = this;
            q.Visible = true;
            q.TopLevel = true;
            this.Visible = false;
        }

        private void deleteRule_Click(object sender, EventArgs e)
        {
            if (listOfRules.SelectedItem != null)
                listOfRules.Items.RemoveAt(listOfRules.SelectedIndex);
        }

        private void closeWindow_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void save_Click(object sender, EventArgs e)
        {
            
            rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Excel.pathForChosenRures, Encoding.UTF8);
            File.Delete(Создать_открыть_файл_правил_Excel.pathForChosenRures);
            using (StreamWriter sw = new StreamWriter(Создать_открыть_файл_правил_Excel.pathForChosenRures, true))
            {
                int i = 0;
                sw.WriteLine("*Rules*");
                while (i < listOfRules.Items.Count)
                {
                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine(listOfRules.Items[i].ToString()); //1 line
                    if (listOfRules.GetItemChecked(i) == true) //2 line
                    {
                        sw.WriteLine("checkState_YES");
                    }
                    else
                    {
                        sw.WriteLine("checkState_NO");
                    }
                    int lineInFile = 0;
                    while (lineInFile < rulesFile.Length & rulesFile[lineInFile] != listOfRules.Items[i].ToString()) //3... lines
                    {

                        lineInFile++;
                        if (rulesFile[lineInFile] == listOfRules.Items[i].ToString())
                        {
                            lineInFile = lineInFile + 2;
                            while (rulesFile[lineInFile] != "**??Rule??**")
                            {
                                sw.WriteLine(rulesFile[lineInFile]);
                                lineInFile++;
                            }
                            break;
                        }
                    }
                    sw.WriteLine("**??Rule??**");
                    i++;
                }
                sw.WriteLine("**Rules**");



                i = 0;
                while (rulesFile[i] != "**Rules**")
                {
                    i++;
                }
                i++;
                while (i < rulesFile.Length)
                {
                    sw.WriteLine(rulesFile[i]);
                    i++;
                }
            }



            Form q = new Saved();
            q.Visible = true;
            q.TopMost = true;
        }

        private void listOfRules_SelectedValueChanged(object sender, EventArgs e)
        {
            try
            {
                ruleDescription.Text = "Название: " + listOfRules.Items[listOfRules.SelectedIndex].ToString();
                int i = 0;
                bool c = false;
                bool w = false;
                while (rulesFile[i].Contains(listOfRules.Items[listOfRules.SelectedIndex].ToString()) != true & i < rulesFile.Length - 2)
                {
                    i++;
                    if (rulesFile[i].Contains(listOfRules.Items[listOfRules.SelectedIndex].ToString()) == true)
                    {
                        c = true;
                    }
                }
                if (c == true)
                {
                    if (rulesFile[i + 2].Contains("shallInclude"))
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Документ должен включать";
                        w = true;
                    }
                    if (rulesFile[i + 2].Contains("shallNotInclude"))
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Документ не должен включать";
                        w = true;
                    }
                    if (rulesFile[i + 2].Contains("if_A_then_B"))
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Если есть А, то должно быть Б";
                        w = true;
                    }
                    if (rulesFile[i + 2].Contains("if_A_not_B"))
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Если есть А, то не должно быть Б";
                        w = true;
                    }
                    if (rulesFile[i + 2].Contains("if_A_then_A_in_B"))
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Если есть А, то A должно быть в Б";
                        w = true;
                    }
                    if (w == false)
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Базовое правило";
                    }
                }
                else
                {
                    if (w == false)
                    {
                        ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Тип правила: Базовое правило";
                    }
                }
                if (rulesFile[i + 2].Contains("shallInclude") | rulesFile[i + 2].Contains("shallNotInclude"))
                {
                    int t = i + 3;
                    string B = "";
                    while (rulesFile[t] != "**??Rule??**")
                    {
                        if (rulesFile[t] == "")
                        {
                            B = B + "\n";
                        }
                        else
                        {
                            if (B == "")
                            {
                                B = rulesFile[t];
                            }
                            else
                            {
                                B = B + "\n" + rulesFile[t];
                            }
                        }
                        t++;
                    }
                    ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "Что: " + "\n" + B;
                }
                if (rulesFile[i + 2].Contains("if_A_then_B") | rulesFile[i + 2].Contains("if_A_not_B") | rulesFile[i + 2].Contains("if_A_then_A_in_B"))
                {
                    int t = i + 4;
                    string B = "";
                    while (rulesFile[t] != "**??Rule??**")
                    {
                        if (rulesFile[t] == "")
                        {
                            B = B + "\n";
                        }
                        else
                        {
                            if (B == "")
                            {
                                B = rulesFile[t];
                            }
                            else
                            {
                                B = B + "\n" + rulesFile[t];
                            }
                        }
                        t++;
                    }
                    ruleDescription.Text = ruleDescription.Text + "\n" + "\n" + "A = " + rulesFile[i + 3] + "\n" + "Б = " + "\n" + B;
                }
                if (listOfRules.SelectedItem != null & w != false)
                {
                    deleteRule.Enabled = true;
                }
                else
                {
                    deleteRule.Enabled = false;
                }
            }
            catch { }
        }

        private void Настройка_правил_Excel_Load(object sender, EventArgs e)
        {
            int i = 0;
            // правила
            string nameOfRule = "";
            string checkState = "";
            while (rulesFile[i].Contains("*Rules*") != true)
            {
                i++;
            }
            while (rulesFile[i].Contains("**Rules**") != true & i < rulesFile.Length)
            {
                while (rulesFile[i].Contains("*??Rule??*") != true & i < rulesFile.Length)
                {
                    i++;
                }
                if (i < rulesFile.Length - 2)
                {

                    nameOfRule = rulesFile[i + 1];
                    checkState = rulesFile[i + 2];
                    if (checkState.Contains("YES"))
                    {
                        listOfRules.Items.Add(nameOfRule, true);
                    }
                    else
                    {
                        listOfRules.Items.Add(nameOfRule, false);
                    }
                }
                while (rulesFile[i].Contains("**??Rule??**") != true & i < rulesFile.Length - 2)
                {
                    i++;
                }
                i++;
            }
        }

        private void Настройка_правил_Excel_VisibleChanged(object sender, EventArgs e)
        {
            rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Excel.pathForChosenRures, Encoding.UTF8);
            listOfRules.Items.Clear();
            int i = 0;
            string nameOfRule = "";
            string checkState = "";
            while (rulesFile[i].Contains("*Rules*") != true)
            {
                i++;
            }
            while (rulesFile[i].Contains("**Rules**") != true & i < rulesFile.Length)
            {
                while (rulesFile[i].Contains("*??Rule??*") != true & i < rulesFile.Length)
                {
                    i++;
                }
                if (i < rulesFile.Length - 2)
                {

                    nameOfRule = rulesFile[i + 1];
                    checkState = rulesFile[i + 2];
                    if (checkState.Contains("YES"))
                    {
                        listOfRules.Items.Add(nameOfRule, true);
                    }
                    else
                    {
                        listOfRules.Items.Add(nameOfRule, false);
                    }
                }
                while (rulesFile[i].Contains("**??Rule??**") != true & i < rulesFile.Length - 2)
                {
                    i++;
                }
                i++;
            }
        }
    }
}
