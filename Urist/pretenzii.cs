using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Urist
{
    public partial class pretenzii : Form
    {
        public pretenzii()
        {
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (dateEdit1.Text != "")
            {
                DateTime date_pretenzii = Convert.ToDateTime(dateEdit1.Text);
                if (Globals.prombyt == 1)
                { date_pretenzii = DateTime.Now.AddDays(7); }
                else { date_pretenzii = DateTime.Now.AddDays(5); }
                this.deloTableAdapter1.UpdatePretenzii(true, date_pretenzii, Globals.id_delo);
                Globals.sbyt();
                Close();
            }
            else { MessageBox.Show("Выберите дату!!!"); }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            Globals.sbyt();
            Close();
        }
    }
}
