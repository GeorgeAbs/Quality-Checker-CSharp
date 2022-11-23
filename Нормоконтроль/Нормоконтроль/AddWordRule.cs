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
    public partial class AddWordRule : Form
    {
        public AddWordRule()
        {
            InitializeComponent();
            try
            {
                rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Word.pathForChosenRures, Encoding.UTF8);
            }
            catch { }
            this.Controls.Remove(A);
            this.Controls.Remove(B);
            this.Controls.Remove(label1);
            this.Controls.Remove(label2);
        }
        string[] rulesFile;

        private void closeForm_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void save_Click(object sender, EventArgs e)
        {
            rulesFile = System.IO.File.ReadAllLines(Создать_открыть_файл_правил_Word.pathForChosenRures, Encoding.UTF8);
            int i = 0;
            while (i < rulesFile.Length-2 & rulesFile[i] != nameOfRule.Text) //проверка на существующее название правила
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


            File.Delete(Создать_открыть_файл_правил_Word.pathForChosenRures);
            using (StreamWriter sw = new StreamWriter(Создать_открыть_файл_правил_Word.pathForChosenRures, true))
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
                if (набор_правил.SelectedItem.ToString() == "1. Документ должен включать А")
                {
                    sw.WriteLine("shallInclude");
                    if (checkRunningTitle.Checked == true)
                    {
                        sw.WriteLine("runningTitle_YES");
                    }
                    else
                    {
                        sw.WriteLine("runningTitle_NO");
                    }
                    int ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "B")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }

                }
                if (набор_правил.SelectedItem.ToString() == "2. Документ не должен включать А")
                {
                    sw.WriteLine("shallNotInclude");
                    if (checkRunningTitle.Checked == true)
                    {
                        sw.WriteLine("runningTitle_YES");
                    }
                    else
                    {
                        sw.WriteLine("runningTitle_NO");
                    }
                    int ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "B")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                }
                if (набор_правил.SelectedItem.ToString() == "3. Если есть А, то должно быть Б")
                {
                    sw.WriteLine("if_A_then_B");
                    if (checkRunningTitle.Checked == true)
                    {
                        sw.WriteLine("runningTitle_YES");
                    }
                    else
                    {
                        sw.WriteLine("runningTitle_NO");
                    }
                    int ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "A")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                    ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "B")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                }
                if (набор_правил.SelectedItem.ToString() == "4. Если есть А, то не должно быть Б")
                {
                    sw.WriteLine("if_A_not_B");
                    if (checkRunningTitle.Checked == true)
                    {
                        sw.WriteLine("runningTitle_YES");
                    }
                    else
                    {
                        sw.WriteLine("runningTitle_NO");
                    }
                    int ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "A")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                    ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "B")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                }
                if (набор_правил.SelectedItem.ToString() == "5. Если есть А, то А должно быть в Б")
                {
                    sw.WriteLine("if_A_then_A_in_B");
                    if (checkRunningTitle.Checked == true)
                    {
                        sw.WriteLine("runningTitle_YES");
                    }
                    else
                    {
                        sw.WriteLine("runningTitle_NO");
                    }
                    int ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "A")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                    ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "B")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                }
                if (набор_правил.SelectedItem.ToString() == "6. Если есть А, то должно быть Б (А - многострочное)")
                {
                    sw.WriteLine("if_A_then_B_multistringA");
                    if (checkRunningTitle.Checked == true)
                    {
                        sw.WriteLine("runningTitle_YES");
                    }
                    else
                    {
                        sw.WriteLine("runningTitle_NO");
                    }
                    int ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "A")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                    ii = 0;
                    while (ii < this.Controls.Count)
                    {
                        if (Controls[ii].Name == "B")
                        {
                            sw.WriteLine(Controls[ii].Text);
                        }
                        ii++;
                    }
                }

                sw.WriteLine("**??Rule??**"); //конец правила

                while (i < rulesFile.Length)
                {
                    sw.WriteLine(rulesFile[i]);
                    i++;
                }
            }

            checkRunningTitle.Checked = false;
            Form q = new Added();
            q.Visible = true;
            q.TopMost = true;
        }

        

        private void AddWordRule_FormClosing(object sender, FormClosingEventArgs e)
        {
            Owner.Visible = true;
        }

        private void набор_правил_SelectedValueChanged(object sender, EventArgs e)
        {
            int i = 0;
            try
            {
                while (i < this.Controls.Count)
                {
                    if (Controls[i].Name == "A")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "B")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "label1")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "label2")
                        this.Controls.Remove(Controls[i]);
                    i++;
                }
            }
            catch { }
            i = 0;
            try
            {
                while (i < this.Controls.Count)
                {
                    if (Controls[i].Name == "A")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "B")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "label1")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "label2")
                        this.Controls.Remove(Controls[i]);
                    i++;
                }
            }
            catch { }
            i = 0;
            try
            {
                while (i < this.Controls.Count)
                {
                    if (Controls[i].Name == "A")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "B")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "label1")
                        this.Controls.Remove(Controls[i]);
                    if (Controls[i].Name == "label2")
                        this.Controls.Remove(Controls[i]);
                    i++;
                }
            }
            catch { }
            

            if (набор_правил.SelectedItem.ToString() == "1. Документ должен включать А")
            {
                RichTextBox B = new System.Windows.Forms.RichTextBox();
                B.Location = new System.Drawing.Point(634, 110);
                B.Name = "B";
                B.Size = new System.Drawing.Size(370, 484);
                B.TabIndex = 33;
                B.Text = "Сюда можно написать несколько \"А\". Чтобы это сделать," +
                    " каждое \"А\" должно отделяться от предыдущего пустой строкой (перед первым - без пустой строки). Пример:" + "\n" +
                    "условие 1" + "\n" + "\n" +
                    "условие 2" + "\n" + "\n" +
                    "условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3" + "\n" + "\n" +
                    "условие 4";

                Label label1 = new System.Windows.Forms.Label();
                label1.AutoSize = true;
                label1.Location = new System.Drawing.Point(823, 94);
                label1.Name = "label1";
                label1.Size = new System.Drawing.Size(14, 13);
                label1.TabIndex = 32;
                label1.Text = "А";

                this.Controls.Add(B);
                this.Controls.Add(label1);
            }
            if (набор_правил.SelectedItem.ToString() == "2. Документ не должен включать А")
            {
                RichTextBox B = new System.Windows.Forms.RichTextBox();
                B.Location = new System.Drawing.Point(634, 110);
                B.Name = "B";
                B.Size = new System.Drawing.Size(370, 484);
                B.TabIndex = 33;
                B.Text = "Сюда можно написать несколько \"А\". Чтобы это сделать," +
                    " каждое \"А\" должно отделяться от предыдущего пустой строкой (перед первым - без пустой строки). Пример:" + "\n" +
                    "условие 1" + "\n" + "\n" +
                    "условие 2" + "\n" + "\n" +
                    "условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3" + "\n" + "\n" +
                    "условие 4";

                Label label1 = new System.Windows.Forms.Label();
                label1.AutoSize = true;
                label1.Location = new System.Drawing.Point(823, 94);
                label1.Name = "label1";
                label1.Size = new System.Drawing.Size(14, 13);
                label1.TabIndex = 32;
                label1.Text = "А";

                this.Controls.Add(B);
                this.Controls.Add(label1);
            }
            if (набор_правил.SelectedItem.ToString() == "3. Если есть А, то должно быть Б")
            {
                RichTextBox B = new System.Windows.Forms.RichTextBox();
                B.Location = new System.Drawing.Point(634, 110);
                B.Name = "B";
                B.Size = new System.Drawing.Size(370, 484);
                B.TabIndex = 33;
                B.Text = "Сюда можно написать несколько \"Б\". Чтобы это сделать," +
                    " каждое \"Б\" должно отделяться от предыдущего пустой строкой (перед первым - без пустой строки). Пример:" + "\n" +
                    "условие 1" + "\n" + "\n" +
                    "условие 2" + "\n" + "\n" +
                    "условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3" + "\n" + "\n" +
                    "условие 4";

                Label label1 = new System.Windows.Forms.Label();
                label1.AutoSize = true;
                label1.Location = new System.Drawing.Point(823, 94);
                label1.Name = "label1";
                label1.Size = new System.Drawing.Size(14, 13);
                label1.TabIndex = 32;
                label1.Text = "Б";

                Label label2 = new System.Windows.Forms.Label();
                label2.AutoSize = true;
                label2.Location = new System.Drawing.Point(467, 94);
                label2.Name = "label2";
                label2.Size = new System.Drawing.Size(14, 13);
                label2.TabIndex = 32;
                label2.Text = "А";

                TextBox A = new System.Windows.Forms.TextBox();
                A.Location = new System.Drawing.Point(297, 110);
                A.Name = "A";
                A.Size = new System.Drawing.Size(330, 22);
                A.TabIndex = 30;

                this.Controls.Add(B);
                this.Controls.Add(A);
                this.Controls.Add(label1);
                this.Controls.Add(label2);
            }
            if (набор_правил.SelectedItem.ToString() == "4. Если есть А, то не должно быть Б")
            {
                RichTextBox B = new System.Windows.Forms.RichTextBox();
                B.Location = new System.Drawing.Point(634, 110);
                B.Name = "B";
                B.Size = new System.Drawing.Size(370, 484);
                B.TabIndex = 33;
                B.Text = "Сюда можно написать несколько \"Б\". Чтобы это сделать," +
                    " каждое \"Б\" должно отделяться от предыдущего пустой строкой (перед первым - без пустой строки). Пример:" + "\n" +
                    "условие 1" + "\n" + "\n" +
                    "условие 2" + "\n" + "\n" +
                    "условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3" + "\n" + "\n" +
                    "условие 4";

                Label label1 = new System.Windows.Forms.Label();
                label1.AutoSize = true;
                label1.Location = new System.Drawing.Point(823, 94);
                label1.Name = "label1";
                label1.Size = new System.Drawing.Size(14, 13);
                label1.TabIndex = 32;
                label1.Text = "Б";

                Label label2 = new System.Windows.Forms.Label();
                label2.AutoSize = true;
                label2.Location = new System.Drawing.Point(467, 94);
                label2.Name = "label2";
                label2.Size = new System.Drawing.Size(14, 13);
                label2.TabIndex = 32;
                label2.Text = "А";

                TextBox A = new System.Windows.Forms.TextBox();
                A.Location = new System.Drawing.Point(297, 110);
                A.Name = "A";
                A.Size = new System.Drawing.Size(330, 22);
                A.TabIndex = 30;

                this.Controls.Add(B);
                this.Controls.Add(A);
                this.Controls.Add(label1);
                this.Controls.Add(label2);
            }
            if (набор_правил.SelectedItem.ToString() == "5. Если есть А, то А должно быть в Б")
            {
                RichTextBox B = new System.Windows.Forms.RichTextBox();
                B.Location = new System.Drawing.Point(634, 110);
                B.Name = "B";
                B.Size = new System.Drawing.Size(370, 484);
                B.TabIndex = 33;
                B.Text = "Сюда можно написать несколько \"Б\". Чтобы это сделать," +
                    " каждое \"Б\" должно отделяться от предыдущего пустой строкой (перед первым - без пустой строки). Пример:" + "\n" +
                    "условие 1" + "\n" + "\n" +
                    "условие 2" + "\n" + "\n" +
                    "условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3" + "\n" + "\n" +
                    "условие 4";

                Label label1 = new System.Windows.Forms.Label();
                label1.AutoSize = true;
                label1.Location = new System.Drawing.Point(823, 94);
                label1.Name = "label1";
                label1.Size = new System.Drawing.Size(14, 13);
                label1.TabIndex = 32;
                label1.Text = "Б";

                Label label2 = new System.Windows.Forms.Label();
                label2.AutoSize = true;
                label2.Location = new System.Drawing.Point(467, 94);
                label2.Name = "label2";
                label2.Size = new System.Drawing.Size(14, 13);
                label2.TabIndex = 32;
                label2.Text = "А";

                TextBox A = new System.Windows.Forms.TextBox();
                A.Location = new System.Drawing.Point(297, 110);
                A.Name = "A";
                A.Size = new System.Drawing.Size(330, 22);
                A.TabIndex = 30;

                this.Controls.Add(B);
                this.Controls.Add(A);
                this.Controls.Add(label1);
                this.Controls.Add(label2);
            }
            if (набор_правил.SelectedItem.ToString() == "6. Если есть А, то должно быть Б (А - многострочное)")
            {
                RichTextBox A = new System.Windows.Forms.RichTextBox();
                A.Location = new System.Drawing.Point(297, 110);
                A.Name = "A";
                A.Size = new System.Drawing.Size(330, 484);
                A.TabIndex = 33;
                A.Text = "Сюда можно написать несколько \"Б\". Чтобы это сделать," +
                    " каждое \"Б\" должно отделяться от предыдущего пустой строкой (перед первым - без пустой строки). Пример:" + "\n" +
                    "условие 1" + "\n" + "\n" +
                    "условие 2" + "\n" + "\n" +
                    "условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3 условие 3" + "\n" + "\n" +
                    "условие 4";

                Label label1 = new System.Windows.Forms.Label();
                label1.AutoSize = true;
                label1.Location = new System.Drawing.Point(823, 94);
                label1.Name = "label1";
                label1.Size = new System.Drawing.Size(14, 13);
                label1.TabIndex = 32;
                label1.Text = "Б";

                Label label2 = new System.Windows.Forms.Label();
                label2.AutoSize = true;
                label2.Location = new System.Drawing.Point(467, 94);
                label2.Name = "label2";
                label2.Size = new System.Drawing.Size(14, 13);
                label2.TabIndex = 32;
                label2.Text = "А";

                TextBox B = new System.Windows.Forms.TextBox();
                B.Location = new System.Drawing.Point(634, 110); 
                B.Name = "B";
                B.Size = new System.Drawing.Size(370, 22);
                B.TabIndex = 30;

                this.Controls.Add(B);
                this.Controls.Add(A);
                this.Controls.Add(label1);
                this.Controls.Add(label2);
            }
        }
    }
}
