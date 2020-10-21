using DevExpress.XtraGrid;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Diagnostics;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;

namespace Urist
{
    public partial class Bankroty : Form
    {
        public Bankroty()
        {
            InitializeComponent();
        }

        private void Bankroty_Load(object sender, EventArgs e)
        {
            this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
            // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
            this.sprStatusTableAdapter.Fill(this.uristDataSet1.sprStatus);
            // TODO: This line of code loads data into the 'uristDataSet1.doc_prikrep_file' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
            // TODO: This line of code loads data into the 'uristDataSet1.sprSlujba' table. You can move, or remove it, as needed.
            this.sprSlujbaTableAdapter.Fill(this.uristDataSet1.sprSlujba);
            this.deloTableAdapter.FillByBlack(this.uristDataSet1.delo, 8);

        }

        private void deloBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_delo = (int)((DataRowView)deloBindingSource.Current).Row["id"];
                Globals.id_prom_byt = (int)((DataRowView)deloBindingSource.Current).Row["id_prom_byt"];
                Globals.prombyt = (int)((DataRowView)deloBindingSource.Current).Row["prom_byt"];


            }
            catch (Exception ex) { }
        }

        private void xtraTabControl4_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl4.SelectedTabPage == xtraTabPage8)
                { this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo); }


            }
            catch { }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (dateEdit1.Text != "" && dateEdit2.Text != "")
            {
                this.deloTableAdapter.FillByBlackDate(this.uristDataSet1.delo, dateEdit1.Text, dateEdit2.Text, 8);
            }
            else
            {
                MessageBox.Show("Выберите дату!!");
            }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.deloTableAdapter.FillByBlack(this.uristDataSet1.delo, 8);
        }

        private void pictureEdit2_Click(object sender, EventArgs e)
        {
             ShowGridPreview(gridControl7);
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

             private void gridView6_DoubleClick(object sender, EventArgs e)
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

