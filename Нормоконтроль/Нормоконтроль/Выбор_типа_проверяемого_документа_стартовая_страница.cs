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
    public partial class Выбор_типа_проверяемого_документа_стартовая_страница : Form
    {
        public Выбор_типа_проверяемого_документа_стартовая_страница()
        {
            InitializeComponent();
        }

        private void word__Click(object sender, EventArgs e)
        {
            Form q = new Form1();
            q.Visible = true;
            q.TopLevel = true;
            q.Owner = this;
            this.Hide();
        }

        private void excel__Click(object sender, EventArgs e)
        {
            Form q = new starting_page_Excel();
            q.Visible = true;
            q.TopLevel = true;
            q.Owner = this;
            this.Hide();
        }

        private void Выбор_типа_проверяемого_документа_стартовая_страница_Load(object sender, EventArgs e)
        {
            //if (/*DateTime.Now.Month > 0 &*/ DateTime.Now.Year > 2021 | Directory.GetCurrentDirectory().Contains("Юдин") !=true | Directory.GetCurrentDirectory().Contains("Users") != true | Directory.GetCurrentDirectory().Contains("Проги") != true)
                //Close();
        }
    }
}
