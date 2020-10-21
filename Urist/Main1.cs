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
    public partial class Main1 : DevExpress.XtraEditors.XtraForm
    {
        private string vershion = "";
        public Main1()
        {
            InitializeComponent();
        }

       
        private void Main1_Load(object sender, EventArgs e)
        {

           
           // Version ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
            

           // string str = ver.ToString();
           // vershion = str; 
           // this.Text = "Вы вошли как: " + Globals.name_user + "  Версия программы: "+vershion;
        }

       
       

        private void tileNavItem1_Tile_ItemClick(object sender, TileItemEventArgs e)
        {
            Form newForm = new Users();

            newForm.Show();
        }

        private void tileNavItem2_Tile_ItemClick(object sender, TileItemEventArgs e)
        {
insert_deb deb_abon = new insert_deb();
            deb_abon.byt_abon();
        }

        private void tileNavItem3_Tile_ItemClick_1(object sender, TileItemEventArgs e)
        {
  insert_prom deb_prom = new insert_prom();
            deb_prom.byt_prom();
        }
    }
}