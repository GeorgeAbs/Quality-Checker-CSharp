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
    public partial class Создать_открыть_файл_правил_Excel : Form
    {
        public Создать_открыть_файл_правил_Excel()
        {
            InitializeComponent();
        }
        public static string pathForChosenRures;
        private void button1_Click(object sender, EventArgs e)
        {
            openFileDialog1.Filter = "Файл с правилами Excel|*.E_rules";
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathForChosenRures = openFileDialog1.FileName;
                Form q = new Настройка_правил_Excel();
                q.Visible = true;
                q.TopLevel = true;
                this.Close();
            }
        }

        private void button2_Click(object sender, EventArgs e)
        {
            saveFileDialog1.Filter = "Файл с правилами Excel|*.E_rules";
            if (saveFileDialog1.ShowDialog() == DialogResult.OK)
            {
                pathForChosenRures = saveFileDialog1.FileName;
                using (StreamWriter sw = new StreamWriter(saveFileDialog1.FileName, true))
                {
                    sw.WriteLine("*Rules*");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на размер шрифта");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на стиль шрифта");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на курсив текста");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на жирность текста");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на выравнивание текста");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на отступы слева/справа/первой строки");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на интервалы до/после");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на двойные пробелы");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка на двойные знаки препинания");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка: цвет текста везде черный");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("*??Rule??*");
                    sw.WriteLine("Проверка: цвет фона текста везде отсутствует");
                    sw.WriteLine("checkState_YES");
                    sw.WriteLine("**??Rule??**");

                    sw.WriteLine("**Rules**");
                }
                Form q = new Настройка_правил_Excel();
                q.Visible = true;
                q.TopLevel = true;
                this.Close();
            }
        }
    }
}
