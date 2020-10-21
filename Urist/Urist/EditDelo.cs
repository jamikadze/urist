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
using DevExpress.XtraGrid;

namespace Urist
{
    public partial class EditDelo : DevExpress.XtraEditors.XtraForm
    {
        public EditDelo()
        {
            InitializeComponent();
        }

        private void gridControl1_Click(object sender, EventArgs e)
        {

        }

        private void EditDelo_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
            // TODO: This line of code loads data into the 'uristDataSet1.delo_deb_prom' table. You can move, or remove it, as needed.
            this.delo_deb_promTableAdapter.FillBySbyt(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi);
            this.delo_deb_bytTableAdapter.FillBySbyt(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi);

        }

        private void simpleButton18_Click(object sender, EventArgs e)
        {
            if (dateEdit6.Text != "" && dateEdit5.Text != "")
            {
                this.delo_deb_bytTableAdapter.FillBySbytDate(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi, dateEdit6.Text, dateEdit5.Text);
            }
        }

        private void simpleButton17_Click(object sender, EventArgs e)
        {
            this.delo_deb_bytTableAdapter.FillBySbyt(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi);
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.delo_deb_promTableAdapter.FillBySbyt(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi);
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (dateEdit2.Text != "" && dateEdit2.Text != "")
            {
            this.delo_deb_promTableAdapter.FillBySbytDate(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi, dateEdit2.Text, dateEdit1.Text);
        } 
        }

        private void pictureEdit2_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl2);
        }
        private void ShowGridPreview(GridControl grid)
        {
            // Check whether the GridControl can be previewed.
            if (!grid.IsPrintingAvailable)
            {
                MessageBox.Show("Библиотека не найдена", "Ошибка");
                return;
            }

            // Open the Preview window.
            grid.ShowPrintPreview();
        }

        private void pictureEdit1_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl1);
        }
    }
}