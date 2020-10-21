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
using System.IO;

namespace Urist
{
    public partial class Users : DevExpress.XtraEditors.XtraForm
    {
        public Users()
        {
            InitializeComponent();
        }

        private void Users_Load(object sender, EventArgs e)
        {
           
          
            // TODO: This line of code loads data into the 'uristDataSet1.sprSlujba' table. You can move, or remove it, as needed.
          
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
            this.sprSlujbaTableAdapter.Fill(this.uristDataSet1.sprSlujba);
        }

        private void gridView1_InitNewRow(object sender, DevExpress.XtraGrid.Views.Grid.InitNewRowEventArgs e)
        {
            usersprogrammBindingSource.AddNew();
        }

        private void gridView1_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            try
            {
                var rowHandle = gridView1.FocusedRowHandle;
                Globals.id_user = (int)gridView1.GetRowCellValue(rowHandle, "user_id");
               
                string fio=(string)gridView1.GetRowCellValue(rowHandle, "fio");
                string login=(string)gridView1.GetRowCellValue(rowHandle, "login");
                 string password=(string)gridView1.GetRowCellValue(rowHandle, "password");
                 int id_doljnost;
                 if (gridView1.GetRowCellValue(rowHandle, "id_doljnost") is DBNull)
                 { id_doljnost = 0; }
                 else
                 {
                     id_doljnost = (int)gridView1.GetRowCellValue(rowHandle, "id_doljnost");
                 }
                 double tel;
                 if (gridView1.GetRowCellValue(rowHandle, "tel") is DBNull)
                 { tel = 0; }
                 else
                 { 
                 tel=(double)gridView1.GetRowCellValue(rowHandle, "tel");}
                 string email;
                if(gridView1.GetRowCellValue(rowHandle, "email") is DBNull)
                {email="";}
                else{
                email=(string)gridView1.GetRowCellValue(rowHandle, "email");}
                DateTime? data_prin;
                if(gridView1.GetRowCellValue(rowHandle, "data_prin") is DBNull)
                {data_prin=null;}
                else{
                    data_prin = (DateTime)gridView1.GetRowCellValue(rowHandle, "data_prin");
                }
                DateTime? data_uvoln;
                if (gridView1.GetRowCellValue(rowHandle, "data_uvoln") is DBNull)
                { data_uvoln = null; }
                else
                {
                    data_uvoln = (DateTime)gridView1.GetRowCellValue(rowHandle, "data_uvoln");
                }
                 int prava_mod_spravochnika;
                if (gridView1.GetRowCellValue(rowHandle, "prava_mod_spravochnika") is DBNull)
                { prava_mod_spravochnika=0;}
                else
                {
                    prava_mod_spravochnika = (int)gridView1.GetRowCellValue(rowHandle, "prava_mod_spravochnika");
                }
                
               
                if (Globals.id_user < 0)
                { /*insert*/
                    this.users_programmTableAdapter.InsertQuery(Globals.id_slujbi1, fio, login, password, id_doljnost, tel, email, data_prin, data_uvoln, prava_mod_spravochnika, Globals.outlen,Globals.file_name);
                    this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
                }
                else
                { /*update*/
                    this.users_programmTableAdapter.UpdateQuery(Globals.id_slujbi1, fio, login, password, id_doljnost, tel, email, data_prin, data_uvoln, prava_mod_spravochnika, Globals.outlen, Globals.file_name, Globals.id_user);
                    this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
                }
            }
            catch (Exception ex) { }
        }

        private void repositoryItemLookUpEdit1_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            if (e.CloseMode == DevExpress.XtraEditors.PopupCloseMode.Normal)
            {
                Globals.id_slujbi1 = (int)(sender as DevExpress.XtraEditors.LookUpEdit).GetColumnValue("id_slujbi");
                

            }
        }

        private void repositoryItemImageEdit1_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            if (e.CloseMode == DevExpress.XtraEditors.PopupCloseMode.Normal)
            {

                ofdInput.ShowDialog();

            }

        }

        private void ofdInput_FileOk(object sender, CancelEventArgs e)
        {
            string inFile = ofdInput.FileName;
            Globals.file_name = Path.GetFileName(inFile);
            try
            {

               
                using (var fs = File.Open(inFile, FileMode.Open, FileAccess.Read, FileShare.None))
                {
                    if (fs.Length > 100000000)
                    {
                        MessageBox.Show("файл слишком большой для сохранения");
                        return;
                    }
                    Globals.outlen = new byte[fs.Length];
                    fs.Read(Globals.outlen, 0, Convert.ToInt32(fs.Length));

                }
               
            }

            catch (IOException ioex)
            {
                MessageBox.Show("Данные не сохранились");
            }
        }

        private void repositoryItemImageEdit1_Click(object sender, EventArgs e)
        {
           
        }



       
 
    }
}