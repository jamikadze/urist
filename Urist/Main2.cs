using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using System.Deployment.Application;

namespace Urist
{
    public partial class Main2 : DevExpress.XtraEditors.XtraForm
    {
        private string vershion = "";
        public Main2()
        {
            InitializeComponent();
        }

        private void tileNavCategory3_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            Form ff = new Deb();
            ff.Show();
        }

        private void Main2_Load(object sender, EventArgs e)
        {   if (Globals.id_user != 3 && Globals.id_user != 33 && Globals.id_user != 34)
            { tileNavCategory3.Enabled = false; }
            else
            { tileNavCategory3.Enabled = true; }
           
        //   Version ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
         //  string str = ver.ToString();
        //  vershion = str; 
         //  this.Text = "Вы вошли как: " + Globals.name_user + "          Подразделение: " + Globals.slujba + "    Версия программы: "+vershion;
           
        }

        private void tileNavCategory4_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            Form ff = new Deb_delo();
            ff.Show();
        }

        private void tileNavCategory7_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            Form ff = new BlackList(null);
            ff.Show();
        }

        private void navButton3_ElementClick(object sender, DevExpress.XtraBars.Navigation.NavElementEventArgs e)
        {
            Form ff = new otchet();
            ff.Show();
        }
    }
}