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
using System.Diagnostics;

namespace Urist
{
    public partial class Podrobno : DevExpress.XtraEditors.XtraForm
    {
        public Podrobno()
        {
            InitializeComponent();
        }

        private void Podrobno_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'uristDataSet11.sprStatusSud' table. You can move, or remove it, as needed.
            this.sprStatusSudTableAdapter.Fill(this.uristDataSet11.sprStatusSud);
            // TODO: This line of code loads data into the 'uristDataSet11.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet11.users_programm);
            // TODO: This line of code loads data into the 'uristDataSet11.sprPromByt' table. You can move, or remove it, as needed.
            this.sprPromBytTableAdapter.Fill(this.uristDataSet11.sprPromByt);
            // TODO: This line of code loads data into the 'uristDataSet1.doc_prikrep_file' table. You can move, or remove it, as needed.
            
            // TODO: This line of code loads data into the 'uristDataSet1.sud_delo' table. You can move, or remove it, as needed.
            this.sud_deloTableAdapter.FillBycspot(this.uristDataSet1.sud_delo,Globals.cspot);
            int i =suddeloBindingSource.Count;
        }

        private void suddeloBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_delo1 = (int)((DataRowView)suddeloBindingSource.Current).Row["id_delo"];
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);
            }
            catch (Exception ex)
            { }
        }

        private void sudBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_sud1 = (int)((DataRowView)sudBindingSource.Current).Row["id"];
                this.doc_prikrep_fileTableAdapter.FillBySud(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1, Globals.id_sud1);

            }
            catch (Exception ex)
            { }
        }

        private void xtraTabControl3_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl3.SelectedTabPage == xtraTabPage6)
                { this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1); }


            }
            catch { }
        }

        private void gridView5_DoubleClick(object sender, EventArgs e)
        {
            try
            {
               
                byte[] b1 = ((byte[])((DataRowView)(docprikrepfileBindingSource.Current)).Row["prikr_file"]);
                string exens4grid = (string)((DataRowView)(docprikrepfileBindingSource.Current)).Row["name_file"];
                exens4grid = Path.GetExtension(exens4grid);
                string filename = System.IO.Path.GetTempFileName().Replace(".tmp", exens4grid);
                var bw = new BinaryWriter(File.Open(filename, FileMode.OpenOrCreate));
                bw.Write(b1);
                bw.Close();
                Process.Start(filename);
            }
            catch (Exception ex)
            { MessageBox.Show("Прикрепленных файлов не имеется!!!"); }
        }

        private void gridView4_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                byte[] b1 = ((byte[])((DataRowView)(docprikrepfileBindingSource.Current)).Row["prikr_file"]);
                string exens4grid = (string)((DataRowView)(docprikrepfileBindingSource.Current)).Row["name_file"];
                exens4grid = Path.GetExtension(exens4grid);
                string filename = System.IO.Path.GetTempFileName().Replace(".tmp", exens4grid);
                var bw = new BinaryWriter(File.Open(filename, FileMode.OpenOrCreate));
                bw.Write(b1);
                bw.Close();
                Process.Start(filename);
            }
            catch (Exception ex)
            { }
        }
    }
}