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
    public partial class Настройка_заголовков_Word : Form
    {
        public Настройка_заголовков_Word()
        {
            InitializeComponent();
            controlCountWhenInitilize = this.Controls.Count;
            fileWithRules = File.ReadAllLines(Создать_открыть_файл_правил_Word.pathForChosenRures, Encoding.UTF8);
            int i = 0;
            while (fileWithRules[i].Contains("*Titles*")!=true)
            {
                i++;
            }
            i++;
            while (fileWithRules[i].Contains("**Titles**")!=true)
            {
                int controlsCountNow = this.Controls.Count;
                string[] fieldsInRow = fileWithRules[i].Split(new[] { "??" }, StringSplitOptions.RemoveEmptyEntries);

                //
                Label titleLevel = new Label();
                titleLevel.Left = labelTitleLevel;
                titleLevel.Top = top + 5 + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
                titleLevel.Name = "titleLevel" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
                titleLevel.Visible = true;
                titleLevel.Size = new System.Drawing.Size(87, 21);
                titleLevel.Text = "1";
                int ii = 0;
                while (ii < (controlsCountNow - controlCountWhenInitilize) / controlsInARow)
                {
                    titleLevel.Text = titleLevel.Text + "." + (ii + 2).ToString();
                    ii++;
                }
                this.Controls.Add(titleLevel);
                //
                //
                TextBox textBox = new TextBox();
                textBox.Left = textBoxLeft;
                textBox.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
                textBox.Name = "visibleTextBox" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
                textBox.Visible = true;
                textBox.Size = new System.Drawing.Size(100, 22);
                textBox.Text = fieldsInRow[0];
                this.Controls.Add(textBox);
                //
                //
                ComboBox comboBold = new ComboBox();
                comboBold.Left = comboBoxBoldLeft;
                comboBold.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
                comboBold.Name = "comboBold" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
                comboBold.Visible = true;
                comboBold.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboBold.Size = new System.Drawing.Size(87, 21);
                comboBold.Items.AddRange(new object[] {
            "Да",
            "Нет"});
                comboBold.Text = fieldsInRow[1];
                this.Controls.Add(comboBold);
                //
                //
                ComboBox comboItalic = new ComboBox();
                comboItalic.Left = comboBoxItalicLeft;
                comboItalic.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
                comboItalic.Name = "comboItalic" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
                comboItalic.Visible = true;
                comboItalic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboItalic.Size = new System.Drawing.Size(87, 21);
                comboItalic.Items.AddRange(new object[] {
            "Да",
            "Нет"});
                comboItalic.Text = fieldsInRow[2];
                this.Controls.Add(comboItalic);
                //
                //
                ComboBox comboEven = new ComboBox();
                comboEven.Left = comboBoxEvenLeft;
                comboEven.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
                comboEven.Name = "comboEven" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
                comboEven.Visible = true;
                comboEven.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
                comboEven.Size = new System.Drawing.Size(87, 21);
                comboEven.Items.AddRange(new object[] {
            "Слева",
            "Справа",
            "По центру",
            "По ширине"});
                comboEven.Text = fieldsInRow[3];
                this.Controls.Add(comboEven);
                //

                i++;
            }
        }
        string[] fileWithRules;
        int controlCountWhenInitilize;
        int textBoxLeft = 112;
        int comboBoxBoldLeft = 218;
        int comboBoxItalicLeft = 311;
        int comboBoxEvenLeft = 404;
        int labelTitleLevel = 18;
        int controlsInARow = 5;
        int top = 108;
        int topStep = 30;
        private void addTitle_Click(object sender, EventArgs e)
        {
            int controlsCountNow = this.Controls.Count;
            //
            TextBox textBox = new TextBox();
            textBox.Left = textBoxLeft;
            textBox.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize)/ controlsInARow;
            textBox.Name = "visibleTextBox" + ((controlsCountNow - controlCountWhenInitilize)/ controlsInARow + 1).ToString();
            textBox.Visible = true;
            textBox.Size = new System.Drawing.Size(100, 22);
            this.Controls.Add(textBox);
            //
            //
            Label titleLevel = new Label();
            titleLevel.Left = labelTitleLevel;
            titleLevel.Top = top + 5 + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
            titleLevel.Name = "titleLevel" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
            titleLevel.Visible = true;
            titleLevel.Size = new System.Drawing.Size(87, 21);
            titleLevel.Text = "1";
            int i = 0;
            while (i < (controlsCountNow - controlCountWhenInitilize) / controlsInARow)
            {
                titleLevel.Text = titleLevel.Text + "." + (i + 2).ToString();
                i++;
            }
            this.Controls.Add(titleLevel);
            //
            //
            ComboBox comboBold = new ComboBox();
            comboBold.Left = comboBoxBoldLeft;
            comboBold.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
            comboBold.Name = "comboBold" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
            comboBold.Visible = true;
            comboBold.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboBold.Size = new System.Drawing.Size(87, 21);
            comboBold.Items.AddRange(new object[] {
            "Да",
            "Нет"});
            this.Controls.Add(comboBold);
            //
            //
            ComboBox comboItalic = new ComboBox();
            comboItalic.Left = comboBoxItalicLeft;
            comboItalic.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
            comboItalic.Name = "comboItalic" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
            comboItalic.Visible = true;
            comboItalic.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboItalic.Size = new System.Drawing.Size(87, 21);
            comboItalic.Items.AddRange(new object[] {
            "Да",
            "Нет"});
            this.Controls.Add(comboItalic);
            //
            //
            ComboBox comboEven = new ComboBox();
            comboEven.Left = comboBoxEvenLeft;
            comboEven.Top = top + topStep * (controlsCountNow - controlCountWhenInitilize) / controlsInARow;
            comboEven.Name = "comboEven" + ((controlsCountNow - controlCountWhenInitilize) / controlsInARow + 1).ToString();
            comboEven.Visible = true;
            comboEven.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
            comboEven.Size = new System.Drawing.Size(87, 21);
            comboEven.Items.AddRange(new object[] {
            "Слева",
            "Справа",
            "По центру",
            "По ширине"});
            this.Controls.Add(comboEven);
            //
            
        }

        private void remobeTitle_Click(object sender, EventArgs e)
        {
            int countOfRowsOfTitles = (this.Controls.Count - controlCountWhenInitilize) / controlsInARow;
            int i = 0;
            int cc = this.Controls.Count;
            while (i<cc)
            {
                foreach (Control control in this.Controls)
                {
                    if (control.Name.Contains(countOfRowsOfTitles.ToString()))
                    {
                        this.Controls.Remove(control);
                    }
                }
                i++;
            }
            
        }

        private void Close_Click(object sender, EventArgs e)
        {
            Close();
        }

        private void Save_Click(object sender, EventArgs e)
        {
            foreach (Control emptyControl in this.Controls) //проверка на пустые поля
            {
                if (emptyControl.Text == "")
                {
                    Form qq = new Не_все_поля_заполнены();
                    qq.Visible = true;
                    qq.TopMost = true;
                    return;
                }
            }
            File.Delete(Создать_открыть_файл_правил_Word.pathForChosenRures);
            int rows = (this.Controls.Count - controlCountWhenInitilize)/5;
            using (StreamWriter sw = new StreamWriter(Создать_открыть_файл_правил_Word.pathForChosenRures/*Создать_открыть_файл_правил_Word.pathForChosenRures.Substring(0, Создать_открыть_файл_правил_Word.pathForChosenRures.LastIndexOf("\\") +1) + "~" + Создать_открыть_файл_правил_Word.pathForChosenRures.Substring(Создать_открыть_файл_правил_Word.pathForChosenRures.LastIndexOf("\\")+1)*/, true))
            {
                bool b = false;
                int t = 0;
                try
                {
                    while (t < fileWithRules.Length)
                    {
                        if (fileWithRules[t].Contains("*Titles*") == true)
                        {
                            b = true;
                        }
                        t++;
                    }
                    if (b == true)
                    {
                        t = 0;
                        while (fileWithRules[t].Contains("*Titles*") != true)
                        {
                            sw.WriteLine(fileWithRules[t]);
                            t++;
                        }
                    }
                }
                catch { }
                sw.WriteLine("*Titles*");
                int i = 1;
                while (i <= rows)
                {
                    string line = "";
                    foreach(Control control in this.Controls)
                    {
                        if (control.Name.Contains("visibleTextBox") & control.Name.Contains(i.ToString()))
                        {
                            line = line + control.Text.ToString() + "??";
                        }
                    }
                    foreach (Control control in this.Controls)
                    {
                        if (control.Name.Contains("comboBold") & control.Name.Contains(i.ToString()))
                        {
                            line = line + control.Text + "??";
                        }
                    }
                    foreach (Control control in this.Controls)
                    {
                        if (control.Name.Contains("comboItalic") & control.Name.Contains(i.ToString()))
                        {
                            line = line + control.Text + "??";
                        }
                    }
                    foreach (Control control in this.Controls)
                    {
                        if (control.Name.Contains("comboEven") & control.Name.Contains(i.ToString()))
                        {
                            line = line + control.Text;
                        }
                    }
                    sw.WriteLine(line);
                    i++;    
                }
                sw.WriteLine("**Titles**");

                if (b == true)
                {
                    t = 0;
                    while (fileWithRules[t].Contains("**Titles**")!=true)
                    {
                        t++;
                    }
                    t++;
                    try
                    {
                        while (t < fileWithRules.Length)
                        {
                            sw.WriteLine(fileWithRules[t]);
                            t++;
                        }
                    }
                    catch { }
                }
            }
            Form q = new Saved();
            q.Visible = true;
            q.TopMost = true;
            Close();
        }
    }
}
