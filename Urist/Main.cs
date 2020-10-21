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
using DevExpress.XtraGrid;
using DevExpress.XtraGrid.Localization;
using System.Collections;
using DevExpress.XtraGrid.Views.Grid;
using System.IO;
using System.Diagnostics;
using urist;

namespace Urist
{
    public partial class Main : DevExpress.XtraEditors.XtraForm
    {
        private string vershion = "";
        public Main()
        {
            InitializeComponent();

        }

        private void Main_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'uristDataSet1.sprSlujba' table. You can move, or remove it, as needed.
            this.sprSlujbaTableAdapter.Fill(this.uristDataSet1.sprSlujba);
            // TODO: This line of code loads data into the 'uristDataSet1.doc_prikrep_file_jaloby' table. You can move, or remove it, as needed.
           
            // TODO: This line of code loads data into the 'uristDataSet1.jaloby' table. You can move, or remove it, as needed.
           
            // TODO: This line of code loads data into the 'uristDataSet1.iski' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.sud_delo' table. You can move, or remove it, as needed.


            Globals.click = 0;
            // TODO: This line of code loads data into the 'uristDataSet1.doc_prikrep_file' table. You can move, or remove it, as needed.
            this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);

            // TODO: This line of code loads data into the 'uristDataSet11.doc_prikrep_file' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
            this.sprStatusTableAdapter.Fill(this.uristDataSet1.sprStatus);
            // TODO: This line of code loads data into the 'uristDataSet1.sprPromByt' table. You can move, or remove it, as needed.
            this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
            // TODO: This line of code loads data into the 'uristDataSet1.delo' table. You can move, or remove it, as needed.
           // this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);
            // TODO: This line of code loads data into the 'uristDataSet1.delo_deb_prom' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);

            backstageViewControl1.SelectedTab = backstageViewTabItem3;

            /* Version ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;
             string str = ver.ToString();
             vershion = str;*/
            // ovalPictureBox1.Load(Globals.filename1);
            Delegate.RemoveAll(Globals.en_delegate, Globals.en_delegate);
            Delegate.RemoveAll(Globals.sogl, Globals.sogl);
            Delegate.RemoveAll(Globals.pssi, Globals.pssi);
            Delegate.RemoveAll(Globals.en_delegate1, Globals.en_delegate1);
            Delegate.RemoveAll(Globals.sbyt, Globals.sbyt);
            Delegate.RemoveAll(Globals.bankrot, Globals.bankrot);
            Globals.en_delegate = visible_main;
            Globals.en_delegate1 = visible_main_reload;
            Globals.sbyt = visible_main_sbyt;
            Globals.sogl = visible_main_sogl;
            Globals.bankrot = visible_main_bankrot;
            Globals.pssi = visible_main_pssi;

        // Version ver = ApplicationDeployment.CurrentDeployment.CurrentVersion;


       //  string str = ver.ToString();
        //  vershion = str;
        //   this.Text = "Вы вошли как: " + Globals.name_user + "        Подразделение: " + Globals.slujba + "  Версия программы:" + vershion;
        }

        private void visible_main_bankrot()
        {
            Globals.click = 0;
            try
            {
                this.Enabled = true;
                this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 8, 0, 0);
                backstageViewControl1.SelectedTab = backstageViewTabItem14;

            }
            catch (Exception ex) { }
        }

        private void visible_main_sbyt()
        {
            Globals.click = 0;
            try
            {
                this.Enabled = true;
                this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 1, 4,5);
                backstageViewControl1.SelectedTab = backstageViewTabItem2;

            }
            catch (Exception ex) { }
        }
        private void visible_main_sogl()
        {
            Globals.click = 0;
            try
            {
                this.Enabled = true;
                this.sud_deloTableAdapter.FillByUristSoglash(this.uristDataSet1.sud_delo, Globals.id_user);
                backstageViewControl1.SelectedTab = backstageViewTabItem11;

            }
            catch (Exception ex) { }
        }
        private void visible_main_pssi()
        {
            Globals.click = 0;
            try
            {
                this.Enabled = true;
                this.sud_deloTableAdapter.FillByPsi(this.uristDataSet1.sud_delo, Globals.id_user,17);
                backstageViewControl1.SelectedTab = ССИ;

            }
            catch (Exception ex) { }
        }

        private void visible_main_reload()
        {
            Globals.click = 0;
            try
            {
                this.Enabled = true;
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                // TODO: This line of code loads data into the 'uristDataSet11.doc_prikrep_file' table. You can move, or remove it, as needed.

                // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
                this.sprStatusTableAdapter.Fill(this.uristDataSet1.sprStatus);
                // TODO: This line of code loads data into the 'uristDataSet1.sprPromByt' table. You can move, or remove it, as needed.
                this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
                this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
                this.sud_deloTableAdapter.FillByUrist(this.uristDataSet1.sud_delo, Globals.id_user);
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);
                backstageViewControl1.SelectedTab = backstageViewTabItem5;
            }
            catch (Exception ex) { }
        }

        private void visible_main()
        {
            Globals.click = 0;
            try
            {
                this.Enabled = true;
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                // TODO: This line of code loads data into the 'uristDataSet11.doc_prikrep_file' table. You can move, or remove it, as needed.

                // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
                this.sprStatusTableAdapter.Fill(this.uristDataSet1.sprStatus);
                // TODO: This line of code loads data into the 'uristDataSet1.sprPromByt' table. You can move, or remove it, as needed.
                this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
                this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
                this.sud_deloTableAdapter.FillByUrist(this.uristDataSet1.sud_delo, Globals.id_user);


            }
            catch (Exception ex) { }
        }

        private void backstageViewButtonItem2_ItemClick(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem4_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem4_ItemHover(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem1_ItemHover(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem1_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem2_ItemHover(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem2_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem5_ItemHover(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem5_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem3_ItemHover(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem3_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem6_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

        }

        private void backstageViewTabItem6_ItemHover(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {


        }

        private void backstageViewTabItem1_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);
        }

        private void simpleButton18_Click(object sender, EventArgs e)
        {
            if (dateEdit6.Text != "" && dateEdit5.Text != "")
            {
                Globals.sdate = dateEdit6.Text;
                Globals.podate = dateEdit5.Text;
                this.delo_deb_bytTableAdapter.FillByRes2(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res, Globals.sdate, Globals.podate);
            }
            else
            { MessageBox.Show("Выберите дату!!!"); }
        }

        private void simpleButton17_Click(object sender, EventArgs e)
        {
            this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);
        }

        private void gridControl1_Click(object sender, EventArgs e)
        {

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

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage == xtraTabPage1)
            {
                ArrayList rows = new ArrayList();
                Int32[] selectedRowHandles = gridView1.GetSelectedRows();
                for (int i = 0; i < selectedRowHandles.Length; i++)
                {
                    int selectedRowHandle = selectedRowHandles[i];
                    if (selectedRowHandle >= 0)
                        rows.Add(gridView1.GetDataRow(selectedRowHandle));
                }
                try
                {
                    // gridView1.BeginUpdate();
                    for (int i = 0; i < rows.Count; i++)
                    {
                        DataRow row = rows[i] as DataRow;
                        // Change the field value.
                        object id = row["id"];
                        object cspot = row["cspot"];
                        /*тут проверка доделать */
                        this.delo_deb_bytTableAdapter.UpdateQuery(Globals.id_user, 1, (int)id);
                        queriesTableAdapter1.insert_delo_byt((int)id);
                    }



                    this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);

                }
                catch (Exception ex) { }
            }
            if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
            {
                ArrayList rows = new ArrayList();
                Int32[] selectedRowHandles = gridView2.GetSelectedRows();
                for (int i = 0; i < selectedRowHandles.Length; i++)
                {
                    int selectedRowHandle = selectedRowHandles[i];
                    if (selectedRowHandle >= 0)
                        rows.Add(gridView2.GetDataRow(selectedRowHandle));
                }
                try
                {
                    // gridView1.BeginUpdate();
                    for (int i = 0; i < rows.Count; i++)
                    {
                        DataRow row = rows[i] as DataRow;
                        // Change the field value.
                        object id = row["id"];
                        this.delo_deb_promTableAdapter.UpdateQuery(Globals.id_user, 1, (int)id);
                        queriesTableAdapter1.insert_delo_prom((int)id);
                    }



                    this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);

                }
                catch (Exception ex) { }
            }

        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (dateEdit2.Text != "" && dateEdit1.Text != "")
            {
                Globals.sdate = dateEdit2.Text;
                Globals.podate = dateEdit1.Text;
                this.delo_deb_promTableAdapter.FillByRes2(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res, Globals.sdate, Globals.podate);
            }
            else { MessageBox.Show("Выберите дату!!!"); }
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);
        }

        private void pictureEdit2_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl2);
        }

        private void pictureEdit3_EditValueChanged(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl3);
        }

        private void backstageViewTabItem2_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 1, 4,5);
        }

        private void toolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Globals.click = 1;
            ofdInput.ShowDialog();
        }

        private void advBandedGridView1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            GridView View = sender as GridView;
            if (e.RowHandle >= 0)
            {
                DateTime? pret;
                DateTime p = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]));
                if (View.GetRowCellValue(e.RowHandle, View.Columns["date_pretenzii"]) is DBNull || View.GetRowCellValue(e.RowHandle, View.Columns["date_pretenzii"]) == "")
                { pret = null; }
                else
                {
                    pret = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_pretenzii"]));
                }

                DateTime now = DateTime.Now;
                DateTime m = p.AddDays(-7);
                if (m <= now && now <= p)
                {
                    e.Appearance.BackColor = Color.Pink;
                    e.Appearance.ForeColor = Color.Black;
                }

                if (now > p)
                {
                    e.Appearance.BackColor = Color.FromArgb(0xFF, 0x99, 0xCC);
                    e.Appearance.ForeColor = Color.Black;
                }
                if (pret != null)
                {
                    if (now > pret)
                    {
                        e.Appearance.BackColor = Color.FromArgb(0xFF, 0x33, 0x00);
                        e.Appearance.ForeColor = Color.Black;
                    }
                }
                if ((int)(View.GetRowCellValue(e.RowHandle, View.Columns["status"])) == 4)
                {
                    e.Appearance.BackColor = Color.FromArgb(0xCC, 0xFF, 0xFF);
                    e.Appearance.ForeColor = Color.Black;
                }
            }
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
            string note;
            if (advBandedGridView1.GetRowCellValue(rowHandle, "fact_spot2") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "fact_spot2") == "")
            { fact_spot2 = null; }
            else
            {
                fact_spot2 = (string)advBandedGridView1.GetRowCellValue(rowHandle, "fact_spot2");
            }
            if (advBandedGridView1.GetRowCellValue(rowHandle, "note") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "note") == "")
            { note = null; }
            else
            {
                note= (string)advBandedGridView1.GetRowCellValue(rowHandle, "note");
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
            this.deloTableAdapter.UpdateQuery(ypeni, yenergy, yakt, y1, y2, y3, y4, y5, y6, y7, y8, dendbill, dendpay, fact_spot2, fact_adress, telephone, 1,Globals.id_user,note, id);
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 1, 4,5);

        }

        private void delodebbytBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {

        }

        private void ofdInput_FileOk(object sender, CancelEventArgs e)
        {
            foreach (String inFile in ofdInput.FileNames)
            {

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
                        if (Globals.click == 2)
                        {
                            this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, file_name, null, Globals.id_user, 1, Globals.id_sud);
                        }
                        if (Globals.click == 3)
                        {
                            this.queriesTableAdapter1.insert_file(Globals.id_delo1, outlen, file_name, null, Globals.id_user, Convert.ToInt32(Globals.status_sud), Globals.id_sud1);
                            this.doc_prikrep_fileTableAdapter.FillBySud(uristDataSet1.doc_prikrep_file, Globals.id_delo1, Globals.id_sud1);
                        }
                       

                    }
                }

                catch (IOException ex)
                {
                    MessageBox.Show("Данные не сохранились");
                }
            } 
            MessageBox.Show("Файл успешно прикреплен ;) ");
        }

        private void xtraTabControl2_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl2.SelectedTabPage == xtraTabPage4)
                { this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo); }


            }
            catch { }

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

        private void simpleButton21_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView3.FocusedRowHandle;

                for (int i = 0; i <= gridView3.RowCount; i++)
                {
                    var check = gridView3.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView3.GetRowCellValue(i, "id"));
                            doc_prikrep_fileTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_fileTableAdapter.FillByid_delo(uristDataSet1.doc_prikrep_file, Globals.id_delo);
            }
            else { return; }
        }

        private void отсканироватьДокументToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Globals.click = 1;
            Form kk = new MainFrame();
            kk.ShowDialog();

        }

        private void передатьВСудToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var rowHandle = advBandedGridView1.FocusedRowHandle;

            Globals.prom_byt = (int)advBandedGridView1.GetRowCellValue(rowHandle, "prom_byt");
            if (Globals.prom_byt == 1)
            {
                if (advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") == "")
                { Globals.yenergya = 0; }
                else
                {
                    Globals.yenergya = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yenergy");
                }
                if (advBandedGridView1.GetRowCellValue(rowHandle, "ypeni") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "ypeni") == "")
                { Globals.ypenia = 0; }
                else
                { Globals.ypenia = (double)advBandedGridView1.GetRowCellValue(rowHandle, "ypeni"); }
                if (advBandedGridView1.GetRowCellValue(rowHandle, "yakt") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yakt") == "")
                { Globals.yakta = 0; }
                {
                    Globals.yakta = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yakt");
                }





            }
            else
            {
                if (advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") == "")
                { Globals.yenergya = 0; }
                else
                {
                    Globals.yenergya = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yenergy");
                }

            }
            Globals.id_sud_spr = 2;
            Form ff = new sud_delo(0);
            ff.Show();

        }

        private void deloBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_delo = (int)((DataRowView)deloBindingSource.Current).Row["id"];
                Globals.id_prom_byt = (int)((DataRowView)deloBindingSource.Current).Row["id_prom_byt"];
                Globals.prombyt = (int)((DataRowView)deloBindingSource.Current).Row["prom_byt"];
                if (Globals.prombyt == 1)
                {
                    if (((DataRowView)deloBindingSource.Current).Row["yenergy"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["yenergy"] == "")
                    { Globals.yenergy = 0; }
                    else
                    {
                        Globals.yenergy = (double)((DataRowView)deloBindingSource.Current).Row["yenergy"];
                    }
                    if (((DataRowView)deloBindingSource.Current).Row["ypeni"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["ypeni"] == "")
                    { Globals.ypeni = 0; }
                    else
                    { Globals.ypeni = (double)((DataRowView)deloBindingSource.Current).Row["ypeni"]; }
                    if (((DataRowView)deloBindingSource.Current).Row["yakt"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["yakt"] == "")
                    { Globals.yakt = 0; }
                    {
                        Globals.yakt = (double)((DataRowView)deloBindingSource.Current).Row["yakt"];
                    }



                  /*  if (((DataRowView)deloBindingSource.Current).Row["yenergy_saldo"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["yenergy_saldo"] == "")
                    { Globals.yenergy_saldo = 0; }
                    else
                    {
                        Globals.yenergy_saldo = (double)((DataRowView)deloBindingSource.Current).Row["yenergy_saldo"];
                    }
                    if (((DataRowView)deloBindingSource.Current).Row["ypeny"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["ypeny"] == "")
                    { Globals.ypeny = 0; }
                    else
                    { Globals.ypeny = (double)((DataRowView)deloBindingSource.Current).Row["ypeny"]; }
                    if (((DataRowView)deloBindingSource.Current).Row["yakt1"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["yakt1"] == "")
                    { Globals.yakt1 = 0; }
                    {
                        Globals.yakt1 = (double)((DataRowView)deloBindingSource.Current).Row["yakt1"];
                    }*/

                }
                else
                {
                    if (((DataRowView)deloBindingSource.Current).Row["yenergy"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["yenergy"] == "")
                    { Globals.yenergy = 0; }
                    else
                    {
                        Globals.yenergy = (double)((DataRowView)deloBindingSource.Current).Row["yenergy"];
                    }
                    /*if (((DataRowView)deloBindingSource.Current).Row["yenergy_saldo"] is DBNull || ((DataRowView)deloBindingSource.Current).Row["yenergy_saldo"] == "")
                    { Globals.yenergy_saldo = 0; }
                    else
                    {
                        Globals.yenergy_saldo = (double)((DataRowView)deloBindingSource.Current).Row["yenergy_saldo"];
                    }*/

                }
            }
            catch (Exception ex) { }
        }


        private void pictureEdit3_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl3);
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 1, 4,5);
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (dateEdit4.Text != "" && dateEdit3.Text != "")
            {
                Globals.sdate = dateEdit4.Text;
                Globals.podate = dateEdit3.Text;
                this.deloTableAdapter.FillByUristDate(this.uristDataSet1.delo, Globals.id_user, Globals.sdate, Globals.podate, 1, 4,5);
            }
            else { MessageBox.Show("Выберите дату!!!"); }
        }

        private void backstageViewTabItem5_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            возвратВСвязиСОтсутствиемИмуществаToolStripMenuItem.Visible = false;
            this.sud_deloTableAdapter.FillByUrist(this.uristDataSet1.sud_delo, Globals.id_user);
            this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
            this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
        }

        private void suddeloBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_sud2 = (int)((DataRowView)suddeloBindingSource.Current).Row["Expr1"];
                Globals.id_delo1 = (int)((DataRowView)suddeloBindingSource.Current).Row["id_delo"];
                Globals.prom_byt = (int)((DataRowView)suddeloBindingSource.Current).Row["prom_byt"];
                Globals.id_prom_byt = (int)((DataRowView)suddeloBindingSource.Current).Row["id_prom_byt"];
            }
            catch { }
        }

        private void gridControl5_Click(object sender, EventArgs e)
        {

        }

        private void gridControl3_Click(object sender, EventArgs e)
        {

        }

        private void gridView5_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                this.doc_prikrep_fileTableAdapter.FillBySud(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1, Globals.id_sud1);
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

        private void advBandedGridView2_DoubleClick(object sender, EventArgs e)
        {
            this.Enabled = false;
            var rowHandle = advBandedGridView2.FocusedRowHandle;
            try
            {
               
                Globals.id_sud2 = (int)advBandedGridView2.GetRowCellValue(rowHandle,"Expr1");
                Globals.id_delo1 = (int)advBandedGridView2.GetRowCellValue(rowHandle,"id_delo");
                Globals.prom_byt = (int)advBandedGridView2.GetRowCellValue(rowHandle, "prom_byt");
                Globals.id_prom_byt = (int)advBandedGridView2.GetRowCellValue(rowHandle,"id_prom_byt");
               
                    if (advBandedGridView2.GetRowCellValue(rowHandle,"yenergy") is DBNull || advBandedGridView2.GetRowCellValue(rowHandle,"yenergy") == "")
                    { Globals.yenergy = 0; }
                    else
                    {
                        Globals.yenergy = (double)advBandedGridView2.GetRowCellValue(rowHandle,"yenergy");
                    }
                    if (advBandedGridView2.GetRowCellValue(rowHandle,"ypeni") is DBNull || advBandedGridView2.GetRowCellValue(rowHandle,"ypeni") == "")
                    { Globals.ypeni = 0; }
                    else
                    { Globals.ypeni = (double)advBandedGridView2.GetRowCellValue(rowHandle,"ypeni"); }
                    if (advBandedGridView2.GetRowCellValue(rowHandle,"yakt") is DBNull || advBandedGridView2.GetRowCellValue(rowHandle,"yakt") == "")
                    { Globals.yakt = 0; }
                    else
                    {
                        Globals.yakt = (double)advBandedGridView2.GetRowCellValue(rowHandle,"yakt");
                    }



                    if (advBandedGridView2.GetRowCellValue(rowHandle,"yenergy_saldo") is DBNull || advBandedGridView2.GetRowCellValue(rowHandle,"yenergy_saldo") == "")
                    { Globals.yenergy_saldo = 0; }
                    else
                    {
                        Globals.yenergy_saldo = (double)advBandedGridView2.GetRowCellValue(rowHandle,"yenergy_saldo");
                    }
                    if (advBandedGridView2.GetRowCellValue(rowHandle,"ypeny") is DBNull || advBandedGridView2.GetRowCellValue(rowHandle,"ypeny") == "")
                    { Globals.ypeny = 0; }
                    else
                    { Globals.ypeny = (double)advBandedGridView2.GetRowCellValue(rowHandle,"ypeny"); }
                    var s=advBandedGridView2.GetRowCellValue(rowHandle,"yakt1").ToString();
                    if (advBandedGridView2.GetRowCellValue(rowHandle,"yakt1") is DBNull || advBandedGridView2.GetRowCellValue(rowHandle,"yakt1") == "")
                    { Globals.yakt1 = 0; }
                    else
                    {
                        Globals.yakt1 = (double)advBandedGridView2.GetRowCellValue(rowHandle,"yakt1");
                    }

              
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);

            }
            catch (Exception ex)
            {MessageBox.Show(ex.ToString()); }
             
            Globals.id_sud_spr = 1;
            Form ff = new sud_delo(0);
            ff.Show();
        }

        private void toolStripMenuItem3_Click(object sender, EventArgs e)
        {
            Globals.click = 3;
            ofdInput.ShowDialog();

        }

        private void toolStripMenuItem4_Click(object sender, EventArgs e)
        {
            Globals.click = 3;
            Form kk = new MainFrame();
            kk.ShowDialog();
        }




        private void pictureEdit4_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl5);
        }

        private void backstageViewTabItem3_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            try
            {
                this.sud_deloTableAdapter.FillByDate(this.uristDataSet1.sud_delo, Globals.id_user, DateTime.Now, DateTime.Now.AddDays(7));
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
            }
            catch (Exception ex) { }

        }

        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
                {
                    this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);
                }

            }
            catch { }
        }

        private void backstageViewTabItem6_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            try
            {
                this.sud_deloTableAdapter.FillByUristZAkrytye(this.uristDataSet1.sud_delo, Globals.id_user);
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
            }
            catch (Exception ex) { }
        }

        private void gridControl7_Click(object sender, EventArgs e)
        {

        }

        private void передатьВСудToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var rowHandle = advBandedGridView4.FocusedRowHandle;
            try
            {

                Globals.id_sud2 = (int)advBandedGridView4.GetRowCellValue(rowHandle, "Expr1");
                Globals.id_delo1 = (int)advBandedGridView4.GetRowCellValue(rowHandle, "id_delo");
                Globals.prom_byt = (int)advBandedGridView4.GetRowCellValue(rowHandle, "prom_byt");
                Globals.id_prom_byt = (int)advBandedGridView4.GetRowCellValue(rowHandle, "id_prom_byt");
               
                    if (advBandedGridView4.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView4.GetRowCellValue(rowHandle, "yenergy") == "")
                    { Globals.yenergy = 0; }
                    else
                    {
                        Globals.yenergy = (double)advBandedGridView4.GetRowCellValue(rowHandle, "yenergy");
                    }
                    if (advBandedGridView4.GetRowCellValue(rowHandle, "ypeni") is DBNull || advBandedGridView4.GetRowCellValue(rowHandle, "ypeni") == "")
                    { Globals.ypeni = 0; }
                    else
                    { Globals.ypeni = (double)advBandedGridView4.GetRowCellValue(rowHandle, "ypeni"); }
                    if (advBandedGridView4.GetRowCellValue(rowHandle, "yakt") is DBNull || advBandedGridView4.GetRowCellValue(rowHandle, "yakt") == "")
                    { Globals.yakt = 0; }
                    else
                    {
                        Globals.yakt = (double)advBandedGridView4.GetRowCellValue(rowHandle, "yakt");
                    }



                    if (advBandedGridView4.GetRowCellValue(rowHandle, "yenergy_saldo") is DBNull || advBandedGridView4.GetRowCellValue(rowHandle, "yenergy_saldo") == "")
                    { Globals.yenergy_saldo = 0; }
                    else
                    {
                        Globals.yenergy_saldo = (double)advBandedGridView4.GetRowCellValue(rowHandle, "yenergy_saldo");
                    }
                    if (advBandedGridView4.GetRowCellValue(rowHandle, "ypeny") is DBNull || advBandedGridView4.GetRowCellValue(rowHandle, "ypeny") == "")
                    { Globals.ypeny = 0; }
                    else
                    { Globals.ypeny = (double)advBandedGridView4.GetRowCellValue(rowHandle, "ypeny"); }
                    if (advBandedGridView4.GetRowCellValue(rowHandle, "yakt1") is DBNull || advBandedGridView4.GetRowCellValue(rowHandle, "yakt1") == "")
                    { Globals.yakt1 = 0; }
                    else
                    {
                        Globals.yakt1 = (double)advBandedGridView4.GetRowCellValue(rowHandle, "yakt1");
                    }

               
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);

            }
            catch (Exception ex)
            { }
            Globals.id_sud_spr = 1;
            Form ff = new sud_delo(0);
            ff.Show();
        }

        private void advBandedGridView2_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;
            if (e.RowHandle >= 0)
            {
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]) != DBNull.Value))
                {
                    DateTime p = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]));
                    DateTime now = DateTime.Now;
                    DateTime m = p.AddDays(-4);
                    if (m <= now && now <= p)
                    {
                        e.Appearance.BackColor = Color.Pink;
                        e.Appearance.ForeColor = Color.Black;
                    }

                    if (now > p)
                    {
                        e.Appearance.BackColor = Color.FromArgb(0xFF, 0xCC, 0x99);
                        e.Appearance.ForeColor = Color.Black;
                    }
                }
            }
        }

        private void передатьНаПроработкуToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Globals.stat = 1;
            try
            {

                this.Enabled = false;
                Form dd = new PerSbyt(2);
                dd.Show();



            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void передатьНаПроработкуToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \n хотите перекинуть абонента в \n папку безнадежные абоненты? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (xtraTabControl1.SelectedTabPage == xtraTabPage1)
                {
                    ArrayList rows = new ArrayList();
                    Int32[] selectedRowHandles = gridView1.GetSelectedRows();
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView1.GetDataRow(selectedRowHandle));
                    }
                    try
                    {
                        // gridView1.BeginUpdate();
                        for (int i = 0; i < rows.Count; i++)
                        {
                            DataRow row = rows[i] as DataRow;
                            // Change the field value.
                            object id = row["id"];
                            object cspot = row["cspot"];
                            /*тут проверка доделать */
                            this.delo_deb_bytTableAdapter.UpdateQuery(Globals.id_user, 3, (int)id);
                            queriesTableAdapter1.insert_delo_byt((int)id);
                        }



                        this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);

                    }
                    catch (Exception ex) { }
                }
                if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
                {
                    ArrayList rows = new ArrayList();
                    Int32[] selectedRowHandles = gridView2.GetSelectedRows();
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView2.GetDataRow(selectedRowHandle));
                    }
                    try
                    {
                        // gridView1.BeginUpdate();
                        for (int i = 0; i < rows.Count; i++)
                        {
                            DataRow row = rows[i] as DataRow;
                            // Change the field value.
                            object id = row["id"];
                            this.delo_deb_promTableAdapter.UpdateQuery(Globals.id_user, 3, (int)id);
                            queriesTableAdapter1.insert_delo_prom((int)id);
                        }



                        this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);

                    }
                    catch (Exception ex) { }
                }
            }
            else { return; }
        }

        private void backstageViewTabItem7_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {

            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 3, 0,0);
        }



        private void simpleButton5_Click(object sender, EventArgs e)
        {
            this.delo_deb_bytTableAdapter.FillByBlackList(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);
        }

        private void simpleButton8_Click(object sender, EventArgs e)
        {
            this.delo_deb_promTableAdapter.FillByBlackList(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);
        }

        private void simpleButton6_Click(object sender, EventArgs e)
        {
            if (dateEdit8.Text != "" && dateEdit7.Text != "")
            {
                this.delo_deb_bytTableAdapter.FillByBlackListDate(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res, dateEdit8.Text, dateEdit7.Text);
            }
            else { MessageBox.Show("Выберите дату"); }
        }



        private void backstageViewTabItem8_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 2, 0,0);
        }

        private void simpleButton11_Click(object sender, EventArgs e)
        {
            if (dateEdit11.Text != "" && dateEdit12.Text != "")
            {
                this.deloTableAdapter.FillByUristDate(this.uristDataSet1.delo, Globals.id_user, dateEdit11.Text, dateEdit12.Text, 2, 0,0);
            }
            else { MessageBox.Show("Выберите дату"); }
        }

        private void simpleButton10_Click(object sender, EventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 2, 0,0);
        }

        private void xtraTabControl5_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl5.SelectedTabPage == xtraTabPage10)
                { this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo); }


            }
            catch { }
        }

        private void xtraTabControl6_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl6.SelectedTabPage == xtraTabPage12)
                { this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1); }


            }
            catch { }

        }



        private void gridView11_DoubleClick(object sender, EventArgs e)
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

        private void gridView10_DoubleClick(object sender, EventArgs e)
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
                Form f = new PerSbyt(1);
                f.Show();

            }
            else { return; }
        }

        private void претензииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            var rowHandle = advBandedGridView1.FocusedRowHandle;
            Boolean pretenzya;
            if (advBandedGridView1.GetRowCellValue(rowHandle, "pretenzii") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "pretenzii") == "")
            {
                this.Enabled=false;

            Form d = new pretenzii();
            d.Show(); 
            }
            else
            { MessageBox.Show("Этому абоненту ранее была предъявлена претензия!!!"); }
           
           
        }

        private void закрытьДелоНаОсновеПретензииToolStripMenuItem_Click(object sender, EventArgs e)
        {
            this.Enabled = false;
            var rowHandle = advBandedGridView1.FocusedRowHandle;
            Boolean pretenzya;
            if (advBandedGridView1.GetRowCellValue(rowHandle, "pretenzii") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "pretenzii") == "")
            {
                pretenzya = false;
            }
            else
            { pretenzya = true; }
            if (pretenzya == true)
            {
                Globals.prom_byt = (int)advBandedGridView1.GetRowCellValue(rowHandle, "prom_byt");
                if (Globals.prom_byt == 1)
                {
                    if (advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") == "")
                    { Globals.yenergya = 0; }
                    else
                    {
                        Globals.yenergya = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yenergy");
                    }
                    if (advBandedGridView1.GetRowCellValue(rowHandle, "ypeni") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "ypeni") == "")
                    { Globals.ypenia = 0; }
                    else
                    { Globals.ypenia = (double)advBandedGridView1.GetRowCellValue(rowHandle, "ypeni"); }
                    if (advBandedGridView1.GetRowCellValue(rowHandle, "yakt") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yakt") == "")
                    { Globals.yakta = 0; }
                    {
                        Globals.yakta = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yakt");
                    }





                }
                else
                {
                    if (advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView1.GetRowCellValue(rowHandle, "yenergy") == "")
                    { Globals.yenergya = 0; }
                    else
                    {
                        Globals.yenergya = (double)advBandedGridView1.GetRowCellValue(rowHandle, "yenergy");
                    }

                }
                Globals.id_sud_spr = 11;
                Form ff = new sud_delo(11);
                ff.Show();
            }
            else
            {
                MessageBox.Show("Вы не можете закрыть это дело на основе претензии, ранее не было предъявлено!!!!");
                this.Enabled = true;
            }
        }

        private void xtraTabControl4_SelectedPageChanged_1(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl4.SelectedTabPage == xtraTabPage8)
                {
                    this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo);
                }

            }
            catch { }


        }

        private void simpleButton5_Click_1(object sender, EventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 3, 0,0);
        }

        private void simpleButton6_Click_1(object sender, EventArgs e)
        {

            if (dateEdit7.Text != "" && dateEdit8.Text != "")
            {
                this.deloTableAdapter.FillByUristDate(this.uristDataSet1.delo, Globals.id_user, dateEdit7.Text, dateEdit8.Text, 3, 0,0);
            }
            else { MessageBox.Show("Выберите дату"); }
        }

        private void pictureEdit5_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl9);
        }

        private void gridView9_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            iskiBindingSource.AddNew();
        }

        private void gridView9_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            try
            {
                var rowHandle = gridView9.FocusedRowHandle;
                Globals.id_iska = (int)gridView9.GetRowCellValue(rowHandle, "id");
                if (gridView9.GetRowCellValue(rowHandle, "date_reg") is DBNull || gridView9.GetRowCellValue(rowHandle, "date_reg") == "")
                { Globals.date_reg = null; }
                else
                {
                    Globals.date_reg = (DateTime)gridView9.GetRowCellValue(rowHandle, "date_reg");
                }
                if (gridView9.GetRowCellValue(rowHandle, "istes") is DBNull || gridView9.GetRowCellValue(rowHandle, "istes")=="")
                { Globals.istes = ""; }
                else { 
                Globals.istes = (string)gridView9.GetRowCellValue(rowHandle, "istes");}
                if (gridView9.GetRowCellValue(rowHandle, "predmet_iska") is DBNull || gridView9.GetRowCellValue(rowHandle, "predmet_iska")=="")
                { Globals.predmet_iska = ""; }
                else { 
                Globals.predmet_iska = (string)gridView9.GetRowCellValue(rowHandle, "predmet_iska");}
                if (gridView9.GetRowCellValue(rowHandle, "one_ins") is DBNull || gridView9.GetRowCellValue(rowHandle, "one_ins")=="")
                { Globals.one_ins = ""; }
                else
                {
                Globals.one_ins = (string)gridView9.GetRowCellValue(rowHandle, "one_ins"); }
                if (gridView9.GetRowCellValue(rowHandle, "two_ins") is DBNull || gridView9.GetRowCellValue(rowHandle, "two_ins")=="")
                { Globals.two_ins = ""; }
                else { 
                Globals.two_ins = (string)gridView9.GetRowCellValue(rowHandle, "two_ins");}
                if (gridView9.GetRowCellValue(rowHandle, "three_ins") is DBNull || gridView9.GetRowCellValue(rowHandle, "three_ins")=="")
                { Globals.three_ins = ""; }
                else { 
                Globals.three_ins = (string)gridView9.GetRowCellValue(rowHandle, "three_ins");}
                if (gridView9.GetRowCellValue(rowHandle, "three_ins1") is DBNull || gridView9.GetRowCellValue(rowHandle, "three_ins1")=="")
                { Globals.three_ins1 = ""; }
                else { 
                Globals.three_ins1 = (string)gridView9.GetRowCellValue(rowHandle, "three_ins1");}
                if (gridView9.GetRowCellValue(rowHandle, "note") is DBNull || gridView9.GetRowCellValue(rowHandle, "note")=="")
                { Globals.note = ""; }
                else { 
                Globals.note = (string)gridView9.GetRowCellValue(rowHandle, "note");}

                if (gridView9.GetRowCellValue(rowHandle, "summa") is DBNull || gridView9.GetRowCellValue(rowHandle, "summa") == "")
                { Globals.summa = 0; }
                else {
                    Globals.summa = (double)gridView9.GetRowCellValue(rowHandle, "summa");
                }

                if (Globals.id_iska < 0)
                {

                    this.iskiTableAdapter.InsertQuery(Globals.id_res, Globals.id_user, DateTime.Now, Globals.date_reg, Globals.istes, Globals.predmet_iska, Globals.summa, Globals.one_ins, Globals.two_ins, Globals.three_ins, Globals.three_ins1, Globals.note);
                    this.iskiTableAdapter.FillByUristRes(this.uristDataSet1.iski, Globals.id_res, Globals.id_user);
                }
                else
                {
                    this.iskiTableAdapter.UpdateQuery(Globals.date_reg, Globals.istes, Globals.predmet_iska, Convert.ToDecimal(Globals.summa), Globals.one_ins, Globals.two_ins, Globals.three_ins, Globals.three_ins1, Globals.note, Globals.id_iska);
                    this.iskiTableAdapter.FillByUristRes(this.uristDataSet1.iski, Globals.id_res, Globals.id_user);
                }
            }
            catch (Exception ex) { }
        }

        private void backstageViewTabItem9_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.iskiTableAdapter.FillByUristRes(this.uristDataSet1.iski, Globals.id_res, Globals.id_user);
        }

        private void simpleButton8_Click_1(object sender, EventArgs e)
        {
            this.iskiTableAdapter.FillByUristRes(this.uristDataSet1.iski, Globals.id_res, Globals.id_user);
        }

        private void simpleButton9_Click(object sender, EventArgs e)
        {
            if (dateEdit9.Text != "" && dateEdit10.Text != "")
            {
                this.iskiTableAdapter.FillByUristResDate(this.uristDataSet1.iski, Globals.id_res, Globals.id_user, dateEdit9.Text, dateEdit10.Text);
            }
            else { MessageBox.Show("Выберите дату"); }
        }

        private void xtraTabControl7_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl7.SelectedTabPage == xtraTabPage14)
                {
                    this.doc_prikrep_file_iskiTableAdapter.FillByIdIski(this.uristDataSet1.doc_prikrep_file_iski, Globals.id_iska);
                }

            }
            catch { }
        }

        private void iskiBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_iska = (int)((DataRowView)iskiBindingSource.Current).Row["id"];
            }
            catch { }
        }

        private void gridView8_DoubleClick(object sender, EventArgs e)
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

        private void gridView12_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                byte[] b1 = ((byte[])((DataRowView)(docprikrepfileiskiBindingSource.Current)).Row["prikr_file"]);
                string exens4grid = (string)((DataRowView)(docprikrepfileiskiBindingSource.Current)).Row["name_file"];
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

        private void toolStripMenuItem5_Click(object sender, EventArgs e)
        {
            Globals.click = 88;
            openFileDialog1.ShowDialog();
        }

        private void openFileDialog1_FileOk(object sender, CancelEventArgs e)
        {
            foreach (String inFile in ofdInput.FileNames)
            {

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
                        if (Globals.click == 88)
                        {
                            this.queriesTableAdapter1.insert_file_iski(Globals.id_iska, outlen, file_name, null, Globals.id_user, 14, null);
                        }
                        /* if (Globals.click == 99)
                         {
                             this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, file_name, null, Globals.id_user, 1, Globals.id_sud);
                         }
                        */
                      

                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Данные не сохранились");
                }



            }  MessageBox.Show("Файл успешно прикреплен ;) ");
        }

        private void simpleButton12_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView12.FocusedRowHandle;

                for (int i = 0; i <= gridView12.RowCount; i++)
                {
                    var check = gridView12.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView12.GetRowCellValue(i, "id"));
                            doc_prikrep_file_iskiTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_file_iskiTableAdapter.FillByIdIski(uristDataSet1.doc_prikrep_file_iski, Globals.id_iska);
            }
            else { return; }
        }

        private void gridView13_InitNewRow(object sender, InitNewRowEventArgs e)
        {
            jalobyBindingSource.AddNew();

        }

        private void gridView13_RowUpdated(object sender, DevExpress.XtraGrid.Views.Base.RowObjectEventArgs e)
        {
            try
            {
                var rowHandle = gridView13.FocusedRowHandle;
                Globals.id_jaloby = (int)gridView13.GetRowCellValue(rowHandle, "id");
                if (gridView13.GetRowCellValue(rowHandle, "zayavitel") is DBNull || gridView13.GetRowCellValue(rowHandle, "zayavitel") == "")
                { Globals.zayavitel = ""; }
                else
                {
                    Globals.zayavitel = (string)gridView13.GetRowCellValue(rowHandle, "zayavitel");
                }
                if (gridView13.GetRowCellValue(rowHandle, "predmet_obr") is DBNull || gridView13.GetRowCellValue(rowHandle, "predmet_obr") == "")
                { Globals.predmet_obr = ""; }
                else
                {
                    Globals.predmet_obr = (string)gridView13.GetRowCellValue(rowHandle, "predmet_obr");
                }
                if (gridView13.GetRowCellValue(rowHandle, "vhod_number") is DBNull || gridView13.GetRowCellValue(rowHandle, "vhod_number") == "")
                { Globals.vhod_number = ""; }
                else
                {
                    Globals.vhod_number = (string)gridView13.GetRowCellValue(rowHandle, "vhod_number");
                }
                if (gridView13.GetRowCellValue(rowHandle, "ish_number") is DBNull || gridView13.GetRowCellValue(rowHandle, "ish_number") == "")
                { Globals.ish_number = ""; }
                else
                {
                    Globals.ish_number = (string)gridView13.GetRowCellValue(rowHandle, "ish_number");
                }
                if (gridView13.GetRowCellValue(rowHandle, "note") is DBNull || gridView13.GetRowCellValue(rowHandle, "note") == "")
                { Globals.note_jaloby = ""; }
                else
                {
                    Globals.note_jaloby = (string)gridView13.GetRowCellValue(rowHandle, "note");
                }
                if (gridView13.GetRowCellValue(rowHandle, "date_registr") is DBNull || gridView13.GetRowCellValue(rowHandle, "date_registr") == "")
                { Globals.date_registr = null; }
                else
                {
                    Globals.date_registr = (DateTime)gridView13.GetRowCellValue(rowHandle, "date_registr");
                }

                if (Globals.id_jaloby < 0)
                {

                    this.jalobyTableAdapter.InsertQuery(Globals.zayavitel, Globals.predmet_obr,Globals.vhod_number,Globals.ish_number,Globals.note_jaloby,Globals.date_registr,Globals.id_res,Globals.id_user, DateTime.Now);
                    this.jalobyTableAdapter.FillByUristRes(this.uristDataSet1.jaloby, Globals.id_res, Globals.id_user);
                }
                else
                {
                    this.jalobyTableAdapter.UpdateQuery(Globals.zayavitel, Globals.predmet_obr, Globals.vhod_number, Globals.ish_number, Globals.note_jaloby, Globals.date_registr,Globals.id_jaloby);
                    this.jalobyTableAdapter.FillByUristRes(this.uristDataSet1.jaloby, Globals.id_res, Globals.id_user);
                }
            }
            catch (Exception ex) { }
        }

        private void simpleButton13_Click(object sender, EventArgs e)
        {
            this.jalobyTableAdapter.FillByUristRes(this.uristDataSet1.jaloby, Globals.id_res, Globals.id_user);
        }

        private void simpleButton14_Click(object sender, EventArgs e)
        {
            if (dateEdit13.Text != "" && dateEdit14.Text != "")
            {
                this.jalobyTableAdapter.FillByUristResDate(this.uristDataSet1.jaloby, Globals.id_res, Globals.id_user, Convert.ToDateTime(dateEdit13.Text),  Convert.ToDateTime(dateEdit14.Text));
            }
            else { MessageBox.Show("Выберите дату"); }
        }

        private void jalobyBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_jaloby = (int)((DataRowView)jalobyBindingSource.Current).Row["id"];
            }
            catch { }
        }

        private void backstageViewTabItem10_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.jalobyTableAdapter.FillByUristRes(this.uristDataSet1.jaloby, Globals.id_res, Globals.id_user);
        }

        private void openFileDialog2_FileOk(object sender, CancelEventArgs e)
        {
            foreach (String inFile in ofdInput.FileNames)
            {

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

                        this.queriesTableAdapter1.insert_file_jaloby(Globals.id_jaloby, outlen, file_name, null, Globals.id_user, 15, null);

                        /* if (Globals.click == 99)
                         {
                             this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, file_name, null, Globals.id_user, 1, Globals.id_sud);
                         }
                        */
                      

                    }
                }
                catch (IOException ex)
                {
                    MessageBox.Show("Данные не сохранились");
                }


            }  MessageBox.Show("Файл успешно прикреплен ;) ");
        }

        private void toolStripMenuItem7_Click(object sender, EventArgs e)
        {
            openFileDialog2.ShowDialog();
        }

        private void gridView14_DoubleClick(object sender, EventArgs e)
        {
            try
            {
                byte[] b1 = ((byte[])((DataRowView)(docprikrepfilejalobyBindingSource.Current)).Row["prikr_file"]);
                string exens4grid = (string)((DataRowView)(docprikrepfilejalobyBindingSource.Current)).Row["name_file"];
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

        private void xtraTabControl8_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl8.SelectedTabPage == xtraTabPage16)
                {
                    this.doc_prikrep_file_jalobyTableAdapter.FillByIdJaloby(this.uristDataSet1.doc_prikrep_file_jaloby, Globals.id_jaloby);
                }

            }
            catch { }
        }

        private void simpleButton15_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView14.FocusedRowHandle;

                for (int i = 0; i <= gridView14.RowCount; i++)
                {
                    var check = gridView14.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView14.GetRowCellValue(i, "id"));
                            doc_prikrep_file_jalobyTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_file_jalobyTableAdapter.FillByIdJaloby(uristDataSet1.doc_prikrep_file_jaloby, Globals.id_jaloby);
            }
            else { return; }
        }

        private void simpleButton7_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView4.FocusedRowHandle;

                for (int i = 0; i <= gridView4.RowCount; i++)
                {
                    var check = gridView4.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView4.GetRowCellValue(i, "id"));
                            doc_prikrep_fileTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1);
            }
            else { return; }
        }

        private void прикрепитьФайлToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Globals.click = 3;
            ofdInput.ShowDialog();
        }

        private void simpleButton16_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView11.FocusedRowHandle;

                for (int i = 0; i <= gridView11.RowCount; i++)
                {
                    var check = gridView11.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView11.GetRowCellValue(i, "id"));
                            doc_prikrep_fileTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_fileTableAdapter.FillByid_delo(uristDataSet1.doc_prikrep_file, Globals.id_delo1);
            }
            else { return; }
        }

        private void toolStripMenuItem10_Click(object sender, EventArgs e)
        {
            Globals.click = 1;
            ofdInput.ShowDialog();

        }

        private void simpleButton19_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView10.FocusedRowHandle;

                for (int i = 0; i <= gridView10.RowCount; i++)
                {
                    var check = gridView10.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView10.GetRowCellValue(i, "id"));
                            doc_prikrep_fileTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_fileTableAdapter.FillByid_delo(uristDataSet1.doc_prikrep_file, Globals.id_delo1);
            }
            else { return; }
        }

        private void simpleButton20_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView8.FocusedRowHandle;

                for (int i = 0; i <= gridView8.RowCount; i++)
                {
                    var check = gridView8.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView8.GetRowCellValue(i, "id"));
                            doc_prikrep_fileTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_fileTableAdapter.FillByid_delo(uristDataSet1.doc_prikrep_file, Globals.id_delo1);
            }
            else { return; }

        }

        private void pictureEdit6_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl7);
        }

        private void pictureEdit8_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl8);
        }

        private void pictureEdit7_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl11);
        }

        private void pictureEdit9_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl14);

        }

        private void pictureEdit10_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl16);
        }

        private void simpleButton22_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить Статус? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                GridView gv = (GridView)advBandedGridView4.GetDetailView(advBandedGridView4.FocusedRowHandle, 0);

                for (int i = 0; i <= gv.RowCount; i++)
                {
                    var check = gv.GetRowCellValue(i, "vr_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gv.GetRowCellValue(i, "id"));
                            sudTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);
            }
            else { return; }
        }

        private void simpleButton23_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить Дело? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                GridView gv = (GridView)advBandedGridView2.GetDetailView(advBandedGridView2.FocusedRowHandle, 0);

                for (int i = 0; i <= gv.RowCount; i++)
                {
                    var check = gv.GetRowCellValue(i, "vr_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gv.GetRowCellValue(i, "id"));
                            sudTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);
            }
            else { return; }
        }

        private void backstageViewTabItem11_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            возвратВСвязиСОтсутствиемИмуществаToolStripMenuItem.Visible = false;
            try
            {
                this.sud_deloTableAdapter.FillByUristSoglash(this.uristDataSet1.sud_delo, Globals.id_user);
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
            }
            catch { }
        }

        private void simpleButton24_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить Дело? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                GridView gv = (GridView)advBandedGridView7.GetDetailView(advBandedGridView7.FocusedRowHandle, 0);

                for (int i = 0; i <= gv.RowCount; i++)
                {
                    var check = gv.GetRowCellValue(i, "vr_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gv.GetRowCellValue(i, "id"));
                            sudTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);
            }
            else { return; }
        }

        private void xtraTabControl9_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl9.SelectedTabPage == xtraTabPage18)
                { this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1); }


            }
            catch { }
        }

        private void simpleButton25_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \nхотите удалить документ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = gridView16.FocusedRowHandle;

                for (int i = 0; i <= gridView16.RowCount; i++)
                {
                    var check = gridView16.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            int id = Convert.ToInt32(gridView16.GetRowCellValue(i, "id"));
                            doc_prikrep_fileTableAdapter.DeleteQuery(id);

                        }
                    }
                }
                this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo1);
            }
            else { return; }
        }

        private void gridView16_DoubleClick(object sender, EventArgs e)
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

        private void advBandedGridView7_DoubleClick(object sender, EventArgs e)
        {
            this.Enabled = false;
            var rowHandle = advBandedGridView7.FocusedRowHandle;
            try
            {

                Globals.id_sud2 = (int)advBandedGridView7.GetRowCellValue(rowHandle, "Expr1");
                Globals.id_delo1 = (int)advBandedGridView7.GetRowCellValue(rowHandle, "id_delo");
                Globals.prom_byt = (int)advBandedGridView7.GetRowCellValue(rowHandle, "prom_byt");
                Globals.id_prom_byt = (int)advBandedGridView7.GetRowCellValue(rowHandle, "id_prom_byt");
                if (advBandedGridView7.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView7.GetRowCellValue(rowHandle, "yenergy") == "")
                    { Globals.yenergy = 0; }
                    else
                    {
                        Globals.yenergy = (double)advBandedGridView7.GetRowCellValue(rowHandle, "yenergy");
                    }
                    if (advBandedGridView7.GetRowCellValue(rowHandle, "ypeni") is DBNull || advBandedGridView7.GetRowCellValue(rowHandle, "ypeni") == "")
                    { Globals.ypeni = 0; }
                    else
                    { Globals.ypeni = (double)advBandedGridView7.GetRowCellValue(rowHandle, "ypeni"); }
                    if (advBandedGridView7.GetRowCellValue(rowHandle, "yakt") is DBNull || advBandedGridView7.GetRowCellValue(rowHandle, "yakt") == "")
                    { Globals.yakt = 0; }
                    else
                    {
                        Globals.yakt = (double)advBandedGridView7.GetRowCellValue(rowHandle, "yakt");
                    }



                    if (advBandedGridView7.GetRowCellValue(rowHandle, "yenergy_saldo") is DBNull || advBandedGridView7.GetRowCellValue(rowHandle, "yenergy_saldo") == "")
                    { Globals.yenergy_saldo = 0; }
                    else
                    {
                        Globals.yenergy_saldo = (double)advBandedGridView7.GetRowCellValue(rowHandle, "yenergy_saldo");
                    }
                    if (advBandedGridView7.GetRowCellValue(rowHandle, "ypeny") is DBNull || advBandedGridView7.GetRowCellValue(rowHandle, "ypeny") == "")
                    { Globals.ypeny = 0; }
                    else
                    { Globals.ypeny = (double)advBandedGridView7.GetRowCellValue(rowHandle, "ypeny"); }
                    if (advBandedGridView7.GetRowCellValue(rowHandle, "yakt1") is DBNull || advBandedGridView7.GetRowCellValue(rowHandle, "yakt1") == "")
                    { Globals.yakt1 = 0; }
                    else
                    {
                        Globals.yakt1 = (double)advBandedGridView7.GetRowCellValue(rowHandle, "yakt1");
                    }

               
                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);

            }
            catch (Exception ex)
            { }
            Globals.id_sud_spr = 5;
            Form ff = new sud_delo(0);
            ff.Show();
        }

        private void backstageViewClientControl11_Load(object sender, EventArgs e)
        {

        }

        private void ССИ_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            возвратВСвязиСОтсутствиемИмуществаToolStripMenuItem.Visible = true;
            try
            {
                this.sud_deloTableAdapter.FillByPsi(this.uristDataSet1.sud_delo, Globals.id_user,17);
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
            }
            catch { }
        }

        private void pictureEdit11_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl18);
        }

        private void pictureEdit12_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl20);
        }

        private void gridView18_DoubleClick(object sender, EventArgs e)
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

        private void advBandedGridView8_DoubleClick(object sender, EventArgs e)
        {
            this.Enabled = false;
            var rowHandle = advBandedGridView8.FocusedRowHandle;
            try
            {

                Globals.id_sud2 = (int)advBandedGridView8.GetRowCellValue(rowHandle, "Expr1");
                Globals.id_delo1 = (int)advBandedGridView8.GetRowCellValue(rowHandle, "id_delo");
                Globals.prom_byt = (int)advBandedGridView8.GetRowCellValue(rowHandle, "prom_byt");
                Globals.id_prom_byt = (int)advBandedGridView8.GetRowCellValue(rowHandle, "id_prom_byt");
               if (advBandedGridView8.GetRowCellValue(rowHandle, "yenergy") is DBNull || advBandedGridView8.GetRowCellValue(rowHandle, "yenergy") == "")
                    { Globals.yenergy = 0; }
                    else
                    {
                        Globals.yenergy = (double)advBandedGridView8.GetRowCellValue(rowHandle, "yenergy");
                    }
                    if (advBandedGridView8.GetRowCellValue(rowHandle, "ypeni") is DBNull || advBandedGridView8.GetRowCellValue(rowHandle, "ypeni") == "")
                    { Globals.ypeni = 0; }
                    else
                    { Globals.ypeni = (double)advBandedGridView8.GetRowCellValue(rowHandle, "ypeni"); }
                    if (advBandedGridView8.GetRowCellValue(rowHandle, "yakt") is DBNull || advBandedGridView8.GetRowCellValue(rowHandle, "yakt") == "")
                    { Globals.yakt = 0; }
                    else
                    {
                        Globals.yakt = (double)advBandedGridView8.GetRowCellValue(rowHandle, "yakt");
                    }



                    if (advBandedGridView8.GetRowCellValue(rowHandle, "yenergy_saldo") is DBNull || advBandedGridView8.GetRowCellValue(rowHandle, "yenergy_saldo") == "")
                    { Globals.yenergy_saldo = 0; }
                    else
                    {
                        Globals.yenergy_saldo = (double)advBandedGridView8.GetRowCellValue(rowHandle, "yenergy_saldo");
                    }
                    if (advBandedGridView8.GetRowCellValue(rowHandle, "ypeny") is DBNull || advBandedGridView8.GetRowCellValue(rowHandle, "ypeny") == "")
                    { Globals.ypeny = 0; }
                    else
                    { Globals.ypeny = (double)advBandedGridView8.GetRowCellValue(rowHandle, "ypeny"); }
                    if (advBandedGridView8.GetRowCellValue(rowHandle, "yakt1") is DBNull || advBandedGridView8.GetRowCellValue(rowHandle, "yakt1") == "")
                    { Globals.yakt1 = 0; }
                    else
                    {
                        Globals.yakt1 = (double)advBandedGridView8.GetRowCellValue(rowHandle, "yakt1");
                    }

                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, Globals.id_delo1);

            }
            catch (Exception ex)
            { }
            Globals.id_sud_spr = 1;
            Form ff = new sud_delo(0);
            ff.Show();
        }

        private void gridControl19_Click(object sender, EventArgs e)
        {

        }

        private void мирноеУрегилированиеЗаключениеСоглашениясСрокомToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Globals.id_sud_spr = 55;
            Form ff = new sud_delo(0);
            ff.Show();
        }

        private void backstageViewTabItem12_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            возвратВСвязиСОтсутствиемИмуществаToolStripMenuItem.Visible = true;
            try
            {
                this.sud_deloTableAdapter.FillByPsi(this.uristDataSet1.sud_delo, Globals.id_user,18);
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
            }
            catch { }
        }

        private void pictureEdit13_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl22);

        }

        private void gridView20_DoubleClick(object sender, EventArgs e)
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

        private void возвратВСвязиСОтсутствиемИмуществаToolStripMenuItem_Click(object sender, EventArgs e)
        { DialogResult result = MessageBox.Show("Вы действительно \nхотите перекинуть в папку Возврат в связи с отсутствием имущества? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                var rowHandle = advBandedGridView9.FocusedRowHandle;
                Globals.id_sud2 = (int)advBandedGridView9.GetRowCellValue(rowHandle, "Expr1");
                this.queriesTableAdapter1.isert_sud(Globals.id_sud2);
                this.sud_deloTableAdapter.FillByPsi(this.uristDataSet1.sud_delo, Globals.id_user, 18);
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                this.sprPromBytTableAdapter.Fill(this.uristDataSet1.sprPromByt);
                backstageViewControl1.SelectedTab = backstageViewTabItem12;

            }else{return;}
        }

        private void gridView1_CustomDrawRowIndicator(object sender, RowIndicatorCustomDrawEventArgs e)
        {
            if (e.Info.IsRowIndicator && e.RowHandle >= 0)
            {
                e.Appearance.TextOptions.HAlignment = DevExpress.Utils.HorzAlignment.Far;
                e.Info.DisplayText = (e.RowHandle + 1).ToString();
            }
        }

        private void gridControl17_Click(object sender, EventArgs e)
        {

        }

        private void передатьНаПроработкуToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Globals.stat = 2;
            try
            {

                this.Enabled = false;
                Form dd = new PerSbyt(2);
                dd.Show();



            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void backstageViewButtonItem1_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            Form dd = new otchet();
            dd.Show();
        }

        private void анулированныйАбонентToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \n хотите перекинуть абонента в \n папку анулированные абоненты? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (xtraTabControl1.SelectedTabPage == xtraTabPage1)
                {
                    ArrayList rows = new ArrayList();
                    Int32[] selectedRowHandles = gridView1.GetSelectedRows();
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView1.GetDataRow(selectedRowHandle));
                    }
                    try
                    {
                        // gridView1.BeginUpdate();
                        for (int i = 0; i < rows.Count; i++)
                        {
                            DataRow row = rows[i] as DataRow;
                            // Change the field value.
                            object id = row["id"];
                            object cspot = row["cspot"];
                            /*тут проверка доделать */
                            this.delo_deb_bytTableAdapter.UpdateQuery(Globals.id_user, 6, (int)id);
                            queriesTableAdapter1.insert_delo_byt((int)id);
                        }



                        this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);

                    }
                    catch (Exception ex) { }
                }
                if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
                {
                    ArrayList rows = new ArrayList();
                    Int32[] selectedRowHandles = gridView2.GetSelectedRows();
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView2.GetDataRow(selectedRowHandle));
                    }
                    try
                    {
                        // gridView1.BeginUpdate();
                        for (int i = 0; i < rows.Count; i++)
                        {
                            DataRow row = rows[i] as DataRow;
                            // Change the field value.
                            object id = row["id"];
                            this.delo_deb_promTableAdapter.UpdateQuery(Globals.id_user, 6, (int)id);
                            queriesTableAdapter1.insert_delo_prom((int)id);
                        }



                        this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);

                    }
                    catch (Exception ex) { }
                }
            }
            else { return; }
        }

        private void backstageViewTabItem13_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 6, 0, 0);
        }

        private void simpleButton26_Click(object sender, EventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 6, 0, 0);
        }

        private void simpleButton27_Click(object sender, EventArgs e)
        {
            if (dateEdit15.Text != "" && dateEdit16.Text != "")
            {
                this.deloTableAdapter.FillByUristDate(this.uristDataSet1.delo, Globals.id_user, dateEdit15.Text, dateEdit16.Text, 6, 0, 0);
            }
            else { MessageBox.Show("Выберите дату"); }
        }

        private void pictureEdit14_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl24);
        }

        private void backstageViewButtonItem2_ItemPressed(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            Form ff = new Deb();
            ff.Show();
        }

        private void backstageViewTabItem14_SelectedChanged(object sender, DevExpress.XtraBars.Ribbon.BackstageViewItemEventArgs e)
        {
            this.deloTableAdapter.FillByUrist(this.uristDataSet1.delo, Globals.id_user, 8, 0, 0);
        }

        private void банкротToolStripMenuItem_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \n хотите перекинуть абонента в \n папку БАНКРОТ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                if (xtraTabControl1.SelectedTabPage == xtraTabPage1)
                {
                    ArrayList rows = new ArrayList();
                    Int32[] selectedRowHandles = gridView1.GetSelectedRows();
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView1.GetDataRow(selectedRowHandle));
                    }
                    try
                    {
                        // gridView1.BeginUpdate();
                        for (int i = 0; i < rows.Count; i++)
                        {
                            DataRow row = rows[i] as DataRow;
                            // Change the field value.
                            object id = row["id"];
                            object cspot = row["cspot"];
                            /*тут проверка доделать */
                            this.delo_deb_bytTableAdapter.UpdateQuery(Globals.id_user, 8, (int)id);
                            queriesTableAdapter1.insert_delo_byt((int)id);
                        }



                        this.delo_deb_bytTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_byt, Globals.prava_mod_spr_res);

                    }
                    catch (Exception ex) { }
                }
                if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
                {
                    ArrayList rows = new ArrayList();
                    Int32[] selectedRowHandles = gridView2.GetSelectedRows();
                    for (int i = 0; i < selectedRowHandles.Length; i++)
                    {
                        int selectedRowHandle = selectedRowHandles[i];
                        if (selectedRowHandle >= 0)
                            rows.Add(gridView2.GetDataRow(selectedRowHandle));
                    }
                    try
                    {
                        // gridView1.BeginUpdate();
                        for (int i = 0; i < rows.Count; i++)
                        {
                            DataRow row = rows[i] as DataRow;
                            // Change the field value.
                            object id = row["id"];
                            this.delo_deb_promTableAdapter.UpdateQuery(Globals.id_user, 8, (int)id);
                            queriesTableAdapter1.insert_delo_prom((int)id);
                        }



                        this.delo_deb_promTableAdapter.FillByRes3(this.uristDataSet1.delo_deb_prom, Globals.prava_mod_spr_res);

                    }
                    catch (Exception ex) { }
                }
            }
            else { return; }
        }

        private void банкротToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            DialogResult result = MessageBox.Show("Вы действительно \n хотите перекинуть абонента в \n папку БАНКРОТ? \n", "Внимание!", MessageBoxButtons.YesNo, MessageBoxIcon.Question);
            if (result == DialogResult.Yes)
            {
                Form f = new Bankrot();
                f.Show();

            }
            else { return; }
        }

        private void банкротToolStripMenuItem2_Click(object sender, EventArgs e)
        {
            Globals.stat = 2;
            try
            {

                this.Enabled = false;
                Form dd = new Bankrot();
                dd.Show();



            }
            catch (Exception ex) { MessageBox.Show(ex.ToString()); }
        }

        private void xtraTabControl13_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            try
            {
                if (xtraTabControl4.SelectedTabPage == xtraTabPage8)
                {
                    this.doc_prikrep_fileTableAdapter.FillByid_delo(this.uristDataSet1.doc_prikrep_file, Globals.id_delo);
                }

            }
            catch { }


        }

        private void gridView21_DoubleClick(object sender, EventArgs e)
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

        private void gridControl25_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl25);
        }

        private void labelControl5_Click(object sender, EventArgs e)
        {

        }

        private void dateEdit5_EditValueChanged(object sender, EventArgs e)
        {

        }

        private void labelControl6_Click(object sender, EventArgs e)
        {

        }

        private void dateEdit6_EditValueChanged(object sender, EventArgs e)
        {

        }

        
    }
}