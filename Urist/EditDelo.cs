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
using System.IO;
using urist;
using System.Diagnostics;

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
            // TODO: This line of code loads data into the 'uristDataSet1.sprSlujba' table. You can move, or remove it, as needed.
            this.sprSlujbaTableAdapter.Fill(this.uristDataSet1.sprSlujba);
            // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
            this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
            // TODO: This line of code loads data into the 'uristDataSet1.sud_delo' table. You can move, or remove it, as needed.
           // this.sud_deloTableAdapter.Fill(this.uristDataSet1.sud_delo);
            // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
            this.sprStatusTableAdapter.FillBySbyt(this.uristDataSet1.sprStatus);
            // TODO: This line of code loads data into the 'uristDataSet1.sprPromByt' table. You can move, or remove it, as needed.
            this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
            // TODO: This line of code loads data into the 'uristDataSet1.delo' table. You can move, or remove it, as needed.
            if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
            { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 2); }
            else { 
            this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 2);}
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
            // TODO: This line of code loads data into the 'uristDataSet1.delo_deb_prom' table. You can move, or remove it, as needed.
            backstageViewControl1.SelectedTab = backstageViewTabItem1;

        }

      
      
        private void simpleButton3_Click(object sender, EventArgs e)
        {
            if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
            { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 2); }
            else
            {
                this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 2);
            }
        }

        private void deloBindingSource_CurrentChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_delo = (int)((DataRowView)deloBindingSource.Current).Row["id"];
                Globals.id_prom_byt = (int)((DataRowView)deloBindingSource.Current).Row["id_prom_byt"];
                Globals.prombyt = (int)((DataRowView)deloBindingSource.Current).Row["prom_byt"];


            }
            catch (Exception ex) { }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (dateEdit4.Text != "" && dateEdit3.Text != "")
            {
                if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                {
                    this.deloTableAdapter.FillBySbytDate1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), dateEdit4.Text, dateEdit3.Text, 2);

                }
                else { this.deloTableAdapter.FillBySbytDate(this.uristDataSet1.delo, Globals.id_slujbi, dateEdit4.Text, dateEdit3.Text, 2); }
            }
            else { MessageBox.Show("Выберите дату!!!"); }
        }

        private void pictureEdit3_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl3);
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

             private void xtraTabControl2_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
             {
                 try
                 {
                     if (xtraTabControl2.SelectedTabPage == xtraTabPage4)
                     { this.doc_prikrep_fileTableAdapter1.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo); }


                 }
                 catch { }
             }

             private void advBandedGridView1_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
             {
                 var rowHandle = advBandedGridView1.FocusedRowHandle;
                 int id = (int)(advBandedGridView1.GetRowCellValue(rowHandle, "id"));
                 string fact_spot2;
                 string fact_adress;
                 string telephone;
                 double ypeni;
                 double yenergy;
                 double yakt;
                 double y1;
                 double y2;
                 double y3;
                 double y4;
                 double y5;
                 double y6;
                 double y7;
                 double y8;
                 string dendbill;
                 string dendpay;
                 int prom_byt;
                 int id_prom_byt;
                 id_prom_byt = (int)advBandedGridView1.GetRowCellValue(rowHandle, "id_prom_byt");
                 prom_byt = (int)advBandedGridView1.GetRowCellValue(rowHandle, "prom_byt");
                 string note;
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "note") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "note") == "")
                 { note = null; }
                 else
                 {
                     note = (string)advBandedGridView1.GetRowCellValue(rowHandle, "note");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "fact_spot2") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "fact_spot2") == "")
                 { fact_spot2 = null; }
                 else
                 {
                     fact_spot2 = (string)advBandedGridView1.GetRowCellValue(rowHandle, "fact_spot2");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "fact_adress") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "fact_adress") == "")
                 { fact_adress = null; }
                 else
                 {
                     fact_adress = (string)advBandedGridView1.GetRowCellValue(rowHandle, "fact_adress");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "telephone") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "telephone") == "")
                 { telephone = null; }
                 else
                 {
                     telephone = (string)advBandedGridView1.GetRowCellValue(rowHandle, "telephone");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "dendbill") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "dendbill") == "")
                 { dendbill = null; }
                 else
                 {
                     dendbill = (string)advBandedGridView1.GetRowCellValue(rowHandle, "dendbill");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "dendpay") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "dendpay") == "")
                 { dendpay = null; }
                 else
                 {
                     dendpay = (string)advBandedGridView1.GetRowCellValue(rowHandle, "dendpay");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "ypeni") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "ypeni") == "")
                 { ypeni = 0; }
                 else
                 {
                     ypeni = (double)advBandedGridView1.GetRowCellValue(rowHandle, "ypeni");
                 }


                 if (advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") == "")
                 { yenergy = 0; }
                 else
                 {
                     yenergy = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yenergy");
                 }

                 if (advBandedGridView1.GetRowCellValue(rowHandle, "yakt") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yakt") == "")
                 { yakt = 0; }
                 else
                 {
                     yakt = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yakt");
                 }

                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y1") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y1") == "")
                 { y1 = 0; }
                 else
                 {
                     y1 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y1");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y2") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y2") == "")
                 { y2 = 0; }
                 else
                 {
                     y2 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y2");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y3") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y3") == "")
                 { y3 = 0; }
                 else
                 {
                     y3 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y3");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y4") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y4") == "")
                 { y4 = 0; }
                 else
                 {
                     y4 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y4");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y5") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y5") == "")
                 { y5 = 0; }
                 else
                 {
                     y5 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y5");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y6") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y6") == "")
                 { y6 = 0; }
                 else
                 {
                     y6 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y6");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y7") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y7") == "")
                 { y7 = 0; }
                 else
                 {
                     y7 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y7");
                 }
                 if (advBandedGridView1.GetRowCellValue(rowHandle, "y8") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "y8") == "")
                 { y8 = 0; }
                 else
                 {
                     y8 = (double)advBandedGridView1.GetRowCellValue(rowHandle, "y8");
                 }
                 this.deloTableAdapter.UpdateQuery(ypeni, yenergy, yakt, y1, y2, y3, y4, y5, y6, y7, y8, dendbill, dendpay, fact_spot2, fact_adress, telephone,4,Globals.id_user, note,id);
                 if (prom_byt == 1)
                 {
                     this.delo_deb_bytTableAdapter.UpdateSbyt(4, note, Globals.id_user, id_prom_byt);

                 }
                 else
                 {
                     this.delo_deb_promTableAdapter.UpdateSbyt(4, note, Globals.id_user, id_prom_byt);

                 }
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 2); }
                 else
                 {
                     this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 2);
                 }
             }

             private void toolStripMenuItem2_Click(object sender, EventArgs e)
             {
                 Globals.click = 1;
                 ofdInput.ShowDialog();
             }

             private void ofdInput_FileOk(object sender, CancelEventArgs e)
             {
                 string inFile = ofdInput.FileName;
                 string file_name = Path.GetFileName(inFile);
                 try
                 {

                     byte[] outlen;
                     using (var fs = File.Open(inFile, FileMode.Open, FileAccess.Read, FileShare.None))
                     {
                         if (fs.Length > 100000000000000000)
                         {
                             MessageBox.Show("файл слишком большой для сохранения");
                             return;
                         }
                         outlen = new byte[fs.Length];
                         fs.Read(outlen, 0, Convert.ToInt32(fs.Length));

                     }
                     if (outlen.Length != 0)
                     {
                         if (Globals.click == 1)
                         {
                             this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, file_name, null, Globals.id_user, 1, null);
                         }
                        
                         MessageBox.Show("Файл успешно прикреплен ;) ");

                     }
                 }

                 catch (IOException ex)
                 {
                     MessageBox.Show("Данные не сохранились");
                 }
             }

             private void отсканироватьДокументToolStripMenuItem_Click(object sender, EventArgs e)
             {
                 Globals.click = 1;
                 Form kk = new MainFrame();
                 kk.ShowDialog();
             }

             private void suddeloBindingSource_CurrentItemChanged(object sender, EventArgs e)
             {
                 try
                 {
                     Globals.id_delo1 = (int)((DataRowView)suddeloBindingSource.Current).Row["id_delo"];
                     Globals.prom_byt = (int)((DataRowView)suddeloBindingSource.Current).Row["prom_byt"];
                     if (Globals.prom_byt == 1)
                     {
                         if (((DataRowView)suddeloBindingSource.Current).Row["yenergy"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["yenergy"] == "")
                         { Globals.yenergy = 0; }
                         else
                         {
                             Globals.yenergy = (double)((DataRowView)suddeloBindingSource.Current).Row["yenergy"];
                         }
                         if (((DataRowView)suddeloBindingSource.Current).Row["ypeni"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["ypeni"] == "")
                         { Globals.ypeni = 0; }
                         else
                         { Globals.ypeni = (double)((DataRowView)suddeloBindingSource.Current).Row["ypeni"]; }
                         if (((DataRowView)suddeloBindingSource.Current).Row["yakt"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["yakt"] == "")
                         { Globals.yakt = 0; }
                         {
                             Globals.yakt = (double)((DataRowView)suddeloBindingSource.Current).Row["yakt"];
                         }



                         if (((DataRowView)suddeloBindingSource.Current).Row["yenergy_saldo"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["yenergy_saldo"] == "")
                         { Globals.yenergy_saldo = 0; }
                         else
                         {
                             Globals.yenergy_saldo = (double)((DataRowView)suddeloBindingSource.Current).Row["yenergy_saldo"];
                         }
                         if (((DataRowView)suddeloBindingSource.Current).Row["ypeny"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["ypeny"] == "")
                         { Globals.ypeny = 0; }
                         else
                         { Globals.ypeny = (double)((DataRowView)suddeloBindingSource.Current).Row["ypeny"]; }
                         if (((DataRowView)suddeloBindingSource.Current).Row["yakt1"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["yakt1"] == "")
                         { Globals.yakt1 = 0; }
                         {
                             Globals.yakt1 = (double)((DataRowView)suddeloBindingSource.Current).Row["yakt1"];
                         }

                     }
                     else
                     {
                         if (((DataRowView)suddeloBindingSource.Current).Row["yenergy"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["yenergy"] == "")
                         { Globals.yenergy = 0; }
                         else
                         {
                             Globals.yenergy = (double)((DataRowView)suddeloBindingSource.Current).Row["yenergy"];
                         }
                         if (((DataRowView)suddeloBindingSource.Current).Row["yenergy_saldo"] is DBNull || ((DataRowView)suddeloBindingSource.Current).Row["yenergy_saldo"] == "")
                         { Globals.yenergy_saldo = 0; }
                         else
                         {
                             Globals.yenergy_saldo = (double)((DataRowView)suddeloBindingSource.Current).Row["yenergy_saldo"];
                         }

                     }
                     this.sudTableAdapter1.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);

                 }
                 catch (Exception ex)
                 { }
             }

             private void xtraTabControl3_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
             {
                 try
                 {
                     if (xtraTabControl3.SelectedTabPage == xtraTabPage6)
                     { this.doc_prikrep_fileTableAdapter1.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1); }


                 }
                 catch { }
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

             private void gridView3_DoubleClick(object sender, EventArgs e)
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

             private void backstageViewTabItem3_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 3); }
                 else
                 {
                     this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 3);
                 }
             }

           

          

             

           
             private void backstageViewTabItem2_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 {
                     this.sud_deloTableAdapter.FillBySbyt1(this.uristDataSet1.sud_delo, Globals.id_slujbi.ToString());
                 }
                 else { this.sud_deloTableAdapter.FillBySbyt(this.uristDataSet1.sud_delo, Globals.id_slujbi); }
             }

             private void backstageViewTabItem4_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 {
                     this.sud_deloTableAdapter.FillBySbytZakryt1(this.uristDataSet1.sud_delo, Globals.id_slujbi.ToString());
                 }
                 else { this.sud_deloTableAdapter.FillBySbytZakryt(this.uristDataSet1.sud_delo, Globals.id_slujbi); }
             }

             private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
             {
                 try
                 {
                     if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
                     { this.doc_prikrep_fileTableAdapter1.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1); }


                 }
                 catch { }
             }

             private void pictureEdit1_Click(object sender, EventArgs e)
             {
                 ShowGridPreview(gridControl1);
             }

             private void pictureEdit4_Click(object sender, EventArgs e)
             {
                 ShowGridPreview(gridControl5);
             }

            

             private void gridView2_DoubleClick(object sender, EventArgs e)
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

             private void безнадежныйАбонентToolStripMenuItem_Click(object sender, EventArgs e)
             {
                 DialogResult result = MessageBox.Show("Вы действительно \n хотите перекинуть абонента в \n папку безнадежные абоненты? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
                 if (result == DialogResult.Yes)
                 {
                     if (Globals.prombyt == 1)
                     {
                         this.delo_deb_bytTableAdapter.UpdateQuery(Globals.id_user, 3, Globals.id_prom_byt);
                         this.deloTableAdapter.UpdateBlackList(3, Globals.id_delo);
                     }
                     if (Globals.prombyt == 2)
                     {
                         this.delo_deb_promTableAdapter.UpdateQuery(Globals.id_user, 3, Globals.id_prom_byt);
                         this.deloTableAdapter.UpdateBlackList(3, Globals.id_delo);
                     }
                     if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                     { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 2); }
                     else
                     {
                         this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 2);
                     }
                 }
                 else { return; }
               
             }

             private void gridControl1_Click_1(object sender, EventArgs e)
             {

             }

             private void backstageViewTabItem1_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 2); }
                 else
                 {
                     this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 2);
                 }
             }

             private void xtraTabControl4_SelectedPageChanged_1(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
             {
                 try
                 {
                     if (xtraTabControl4.SelectedTabPage == xtraTabPage8)
                     {this.doc_prikrep_fileTableAdapter1.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo);  }


                 }
                 catch { }
             }

             private void pictureEdit2_Click(object sender, EventArgs e)
             {
                 ShowGridPreview(gridControl7);
             }

             private void simpleButton2_Click(object sender, EventArgs e)
             {
                 if (dateEdit1.Text != "" && dateEdit2.Text != "")
                 {
                     if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                     {
                         this.deloTableAdapter.FillBySbytDate1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), dateEdit1.Text, dateEdit2.Text, 3);

                     }
                     else { this.deloTableAdapter.FillBySbytDate(this.uristDataSet1.delo, Globals.id_slujbi, dateEdit1.Text, dateEdit2.Text, 3); }
                 }
                 else { MessageBox.Show("Выберите дату"); }
             }

             private void simpleButton1_Click(object sender, EventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 3); }
                 else
                 {
                     this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 3);
                 }
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

             private void advBandedGridView2_DoubleClick(object sender, EventArgs e)
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

             private void backstageViewTabItem5_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 {
                     this.delo_deb_bytTableAdapter.FillByRes31(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi.ToString());
                 }
                 else
                 { this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi); }
             }

             private void xtraTabControl5_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
             {
                 try
                 {
                     if (xtraTabControl5.SelectedTabPage == xtraTabPage10)
                     {
                         if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                         {
                             this.delo_deb_promTableAdapter.FillByRes31(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi.ToString());
                         }
                         else { this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi); }
                     }

                 }
                 catch { }
             }

             private void simpleButton17_Click(object sender, EventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 {
                     this.delo_deb_bytTableAdapter.FillByRes31(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi.ToString());
                 }
                 else
                 { this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi); }

             }

             private void simpleButton5_Click(object sender, EventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 {
                     this.delo_deb_promTableAdapter.FillByRes31(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi.ToString());
                 }
                 else { this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi); }
             }

             private void simpleButton18_Click(object sender, EventArgs e)
             {
                 if (dateEdit6.Text != "" && dateEdit5.Text != "")
                 {
                     Globals.sdate = dateEdit6.Text;
                     Globals.podate = dateEdit5.Text;
                     this.delo_deb_bytTableAdapter.FillByRes2(this.uristDataSet1.delo_deb_byt, Globals.id_slujbi, Globals.sdate, Globals.podate);
                 }
                 else
                 { MessageBox.Show("Выберите дату!!!"); }
             }

             private void simpleButton6_Click(object sender, EventArgs e)
             {
                 if (dateEdit8.Text != "" && dateEdit7.Text != "")
                 {
                     Globals.sdate = dateEdit8.Text;
                     Globals.podate = dateEdit7.Text;
                     this.delo_deb_promTableAdapter.FillByRes2(this.uristDataSet1.delo_deb_prom, Globals.id_slujbi, Globals.sdate, Globals.podate);
                 }
                 else { MessageBox.Show("Выберите дату!!!"); }
             }

             private void pictureEdit5_Click(object sender, EventArgs e)
             {
                 ShowGridPreview(gridControl9);
             }

             private void pictureEdit6_Click(object sender, EventArgs e)
             {
                 ShowGridPreview(gridControl10);
             }

             private void backstageViewTabItem6_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 8); }
                 else
                 {
                     this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 8);
                 }
             }

             private void simpleButton8_Click(object sender, EventArgs e)
             {
                 if (dateEdit9.Text != "" && dateEdit10.Text != "")
                 {
                     if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                     {
                         this.deloTableAdapter.FillBySbytDate1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), dateEdit9.Text, dateEdit10.Text, 8);

                     }
                     else { this.deloTableAdapter.FillBySbytDate(this.uristDataSet1.delo, Globals.id_slujbi, dateEdit9.Text, dateEdit10.Text, 8); }
                 }
                 else { MessageBox.Show("Выберите дату"); }
             }

             private void xtraTabControl6_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
             {
                 try
                 {
                     if (xtraTabControl4.SelectedTabPage == xtraTabPage12)
                     { this.doc_prikrep_fileTableAdapter1.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo); }


                 }
                 catch { }
             }

             private void gridView9_DoubleClick(object sender, EventArgs e)
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

             private void simpleButton7_Click(object sender, EventArgs e)
             {
                 if (Globals.id_user == 48 || Globals.id_user == 49 || Globals.id_user == 50 || Globals.id_user == 51)
                 { this.deloTableAdapter.FillBySbyt1(this.uristDataSet1.delo, Globals.id_slujbi.ToString(), 8); }
                 else
                 {
                     this.deloTableAdapter.FillBySbyt(this.uristDataSet1.delo, Globals.id_slujbi, 8);
                 }
             }

             private void pictureEdit7_Click(object sender, EventArgs e)
             {
                 ShowGridPreview(gridControl1);
             }

             private void backstageViewButtonItem1_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {
                 Globals.send_debit = 1;
                 Form ff = new Deb();
                 ff.Show();
             }

             private void backstageViewButtonItem1_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
             {

             }

             

      

      
        }

       

      
    }
