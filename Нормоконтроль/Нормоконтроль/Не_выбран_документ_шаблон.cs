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
    public partial class Не_выбран_документ_шаблон : Form
    {
        public Не_выбран_документ_шаблон()
        {
            InitializeComponent();
        }
        bool b = false;
        private void yes_Click(object sender, EventArgs e)
        {
            Form q = new Прогресс_проверки();
            q.Visible = true;
            q.Owner = this.Owner;
            q.TopMost = true;
            b = true;
            Close();
        }

        private void canc_Click(object sender, EventArgs e)
        {
            Owner.Visible = true;
            Close();
        }

        private void Не_выбран_документ_шаблон_FormClosing(object sender, FormClosingEventArgs e)
        {
            if (b == true)
            {
            }
            else
            {
                Owner.Visible = true;
            }
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }
    }
}
