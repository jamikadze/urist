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
using DevExpress.XtraGrid.Views.Grid;

namespace Urist
{
    public partial class Deb_delo : DevExpress.XtraEditors.XtraForm
    {
        public Deb_delo()
        {
            InitializeComponent();
        }

        private void Deb_delo_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
            this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
            // TODO: This line of code loads data into the 'uristDataSet1.delo' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
            this.sudTableAdapter.Fill(this.uristDataSet1.sud);
            // TODO: This line of code loads data into the 'uristDataSet1.sud_delo' table. You can move, or remove it, as needed.
            this.sud_deloTableAdapter.Fill(this.uristDataSet1.sud_delo);
            // TODO: This line of code loads data into the 'uristDataSet1.delo_deb_prom' table. You can move, or remove it, as needed.

            // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
            this.sprStatusTableAdapter.Fill(this.uristDataSet1.sprStatus);
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
            this.sprSlujbaTableAdapter.Fill(this.uristDataSet1.sprSlujba);
            if (Globals.id_user == 3 || Globals.id_user == 34 || Globals.id_user == 9)
            { this.delo_deb_bytTableAdapter.FillBySever1(this.uristDataSet1.delo_deb_byt); }
            if (Globals.id_user == 33)
            {
                this.delo_deb_bytTableAdapter.FillByTalas1(this.uristDataSet1.delo_deb_byt);
            }
            if (Globals.id_user == 7 || Globals.id_user == 8 || Globals.id_user == 35 || Globals.id_user == 36)
            { this.delo_deb_bytTableAdapter.FillByAll(this.uristDataSet1.delo_deb_byt); }

        }

        private void sprSlujbaBindingSource_CurrentItemChanged(object sender, EventArgs e)
        {
            try
            {
                Globals.id_res = (int)((DataRowView)(sprSlujbaBindingSource.Current)).Row["id_slujbi"];
                Globals.id_zavis = (int)((DataRowView)(sprSlujbaBindingSource.Current)).Row["id_zavis"];


            }
            catch { }
        }

        private void treeList1_SelectionChanged(object sender, EventArgs e)
        {

        }

        private void simpleButton18_Click(object sender, EventArgs e)
        {
            try
            {
                Globals.sdate = dateEdit6.Text;
                Globals.podate = dateEdit5.Text;
                if (Globals.id_user == 3 || Globals.id_user == 34 || Globals.id_user == 9)
                {
                    this.delo_deb_bytTableAdapter.FillBySever(this.uristDataSet1.delo_deb_byt, Globals.sdate, Globals.podate);
                }
                if (Globals.id_user == 33)
                {
                    this.delo_deb_bytTableAdapter.FillByTalas(this.uristDataSet1.delo_deb_byt, Globals.sdate, Globals.podate);
                }
                if (Globals.id_user == 7 || Globals.id_user == 8 || Globals.id_user == 35 || Globals.id_user == 36)
                { this.delo_deb_bytTableAdapter.FillByAllDate(this.uristDataSet1.delo_deb_byt, Globals.sdate, Globals.podate); }
            }
            catch (Exception ex) { }
        }

        private void treeList1_Click(object sender, EventArgs e)
        {
            try
            {
                if (xtraTabControl1.SelectedTabPage == xtraTabPage1)
                {
                    if (Globals.id_res == 1)
                    {
                        this.delo_deb_bytTableAdapter.FillBySever1(this.uristDataSet1.delo_deb_byt);
                    }
                    else if (Globals.id_res == 14)
                    {
                        this.delo_deb_bytTableAdapter.FillByTalas1(this.uristDataSet1.delo_deb_byt);
                    }
                    else
                    {
                        if (Globals.id_zavis != 1 || Globals.id_zavis != 14)
                        {
                            if (Globals.id_res == 5 || Globals.id_res == 6 || Globals.id_res == 7 || Globals.id_res == 8 || Globals.id_res == 11 || Globals.id_res == 13 || Globals.id_res == 9)
                            {
                                this.delo_deb_bytTableAdapter.FillByRes11(this.uristDataSet1.delo_deb_byt, Globals.id_res.ToString()); ;
                            }
                            else { this.delo_deb_bytTableAdapter.FillByRes1(this.uristDataSet1.delo_deb_byt, Globals.id_res); }



                        }
                    }

                }
                if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
                {
                    if (Globals.id_res == 1)
                    {
                        this.delo_deb_promTableAdapter.FillBySever1(this.uristDataSet1.delo_deb_prom);
                    }
                    else if (Globals.id_res == 14)
                    {
                        this.delo_deb_promTableAdapter.FillByTalas1(this.uristDataSet1.delo_deb_prom);
                    }
                    else
                    {
                        if (Globals.id_zavis != 1 || Globals.id_zavis != 14)
                        {
                            if (Globals.id_res == 3 || Globals.id_res == 2)
                            { this.delo_deb_promTableAdapter.FillByRes1(this.uristDataSet1.delo_deb_prom, Globals.id_res); }
                            else
                            {
                                this.delo_deb_promTableAdapter.FillByUchastok1(this.uristDataSet1.delo_deb_prom, Globals.id_res.ToString());
                            }
                        }
                    }
                }


            }
            catch { }
        }

        private void simpleButton17_Click(object sender, EventArgs e)
        {
            if (Globals.id_user == 3 || Globals.id_user == 34 || Globals.id_user == 9)
            { this.delo_deb_bytTableAdapter.FillBySever1(this.uristDataSet1.delo_deb_byt); }
            if (Globals.id_user == 33)
            {
                this.delo_deb_bytTableAdapter.FillByTalas1(this.uristDataSet1.delo_deb_byt);
            }
            if (Globals.id_user == 7 || Globals.id_user == 8 || Globals.id_user == 35 || Globals.id_user == 36)
            { this.delo_deb_bytTableAdapter.FillByAll(this.uristDataSet1.delo_deb_byt); }
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

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (Globals.id_user == 3 || Globals.id_user == 34 || Globals.id_user == 9)
            { this.delo_deb_promTableAdapter.FillBySever1(this.uristDataSet1.delo_deb_prom); }
            if (Globals.id_user == 33)
            {
                this.delo_deb_promTableAdapter.FillByTalas1(this.uristDataSet1.delo_deb_prom);
            }
            if (Globals.id_user == 7 || Globals.id_user == 8 || Globals.id_user == 35 || Globals.id_user == 36)
            { this.delo_deb_promTableAdapter.FillByAll(this.uristDataSet1.delo_deb_prom); }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            try
            {
                Globals.sdate = dateEdit2.Text;
                Globals.podate = dateEdit1.Text;
                if (Globals.id_user == 3 || Globals.id_user == 34 || Globals.id_user == 9)
                {
                    this.delo_deb_promTableAdapter.FillBySever(this.uristDataSet1.delo_deb_prom, Globals.sdate, Globals.podate);
                }
                if (Globals.id_user == 33)
                {
                    this.delo_deb_promTableAdapter.FillByTalas(this.uristDataSet1.delo_deb_prom, Globals.sdate, Globals.podate);
                }
                if (Globals.id_user == 7 || Globals.id_user == 8 || Globals.id_user == 35 || Globals.id_user == 36)
                { this.delo_deb_promTableAdapter.FillByAllDate(this.uristDataSet1.delo_deb_prom, Globals.sdate, Globals.podate); }
            }
            catch { }
        }

        private void pictureEdit2_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl2);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage == xtraTabPage1)
            {
                try
                {
                    Globals.deb_delo_status = (int)((DataRowView)(delodebbytBindingSource.Current)).Row["status"];
                    if (Globals.deb_delo_status == 1 || Globals.deb_delo_status == 2 || Globals.deb_delo_status == 4 || Globals.deb_delo_status == 5 || Globals.deb_delo_status == 7)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebbytBindingSource.Current)).Row["cspot"];
                        this.deloTableAdapter.FillByCspot(this.uristDataSet1.delo, Globals.cspot);
                        if (deloBindingSource.Count > 0)
                        {
                            if (((DataRowView)(deloBindingSource.Current)).Row["date_peredachi_sud"] is DBNull)
                            {
                                DateTime date = (DateTime)((DataRowView)(deloBindingSource.Current)).Row["date_peredachi_sud"];
                                MessageBox.Show("По абоненту ведется работа по сбору данных. Дата передачи в суд " + date.ToString());
                            }
                            else
                            {
                                int id_delo = (int)((DataRowView)(deloBindingSource.Current)).Row["id"];
                                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, id_delo);
                                if (sudBindingSource.Count > 0)
                                {

                                    Form dd = new Podrobno();
                                    dd.Show();
                                }
                                else { MessageBox.Show("Дело по абоненту еще не передан всуд, поэтому по абоненту нет истории"); }
                            }


                        }
                    }

                    else if (Globals.deb_delo_status == 3)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebbytBindingSource.Current)).Row["cspot"];
                        Form dd = new BlackList(Globals.cspot);
                        dd.Show();
                    }
                    else { MessageBox.Show("Абонент не принят в дело, по этому абоненту нет истории"); }

                }
                catch { MessageBox.Show("По этому абоненту нет истории"); }
            }
            if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
            {
                try
                {
                    Globals.deb_delo_status = (int)((DataRowView)(delodebpromBindingSource.Current)).Row["status"];
                    if (Globals.deb_delo_status == 1 || Globals.deb_delo_status == 2 || Globals.deb_delo_status == 4 || Globals.deb_delo_status == 5 || Globals.deb_delo_status == 7)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebpromBindingSource.Current)).Row["cspot"];
                        this.deloTableAdapter.FillByCspot(this.uristDataSet1.delo, Globals.cspot);
                        if (deloBindingSource.Count > 0)
                        {

                            if (((DataRowView)(deloBindingSource.Current)).Row["date_peredachi_sud"] is DBNull)
                            {
                                DateTime date = (DateTime)((DataRowView)(deloBindingSource.Current)).Row["date_peredachi_sud"];
                                MessageBox.Show("По абоненту ведется работа по сбору данных. Дата передачи в суд " + date.ToString());
                            }
                            else
                            {
                                int id_delo = (int)((DataRowView)(deloBindingSource.Current)).Row["id"];
                                this.sudTableAdapter.FillByDelo(this.uristDataSet1.sud, id_delo);
                                if (sudBindingSource.Count > 0)
                                {

                                    Form dd = new Podrobno();
                                    dd.Show();
                                }
                                else { MessageBox.Show("Дело по абоненту еще не передан всуд, поэтому по абоненту нет истории"); }
                            }
                        }
                    }
                    else if (Globals.deb_delo_status == 3)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebpromBindingSource.Current)).Row["cspot"];
                        Form dd = new BlackList(Globals.cspot);
                        dd.Show();
                    }
                    else { MessageBox.Show("Абонент не принят в дело, по этому абоненту нет истории"); }
                    if (Globals.deb_delo_status == 3)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebpromBindingSource.Current)).Row["cspot"];
                        Form dd = new BlackList(Globals.cspot);
                        dd.Show();
                    }


                }
                catch { MessageBox.Show("По этому абоненту нет истории"); }
            }
        }

        private void xtraTabControl1_SelectedPageChanged(object sender, DevExpress.XtraTab.TabPageChangedEventArgs e)
        {
            if (xtraTabControl1.SelectedTabPage == xtraTabPage2)
            {
                if (Globals.id_user == 3 || Globals.id_user == 34 || Globals.id_user == 9)
                { this.delo_deb_promTableAdapter.FillBySever1(this.uristDataSet1.delo_deb_prom); }
                if (Globals.id_user == 33)
                {
                    this.delo_deb_promTableAdapter.FillByTalas1(this.uristDataSet1.delo_deb_prom);
                }
                if (Globals.id_user == 7 || Globals.id_user == 8 || Globals.id_user == 35 || Globals.id_user == 36)
                { this.delo_deb_promTableAdapter.FillByAll(this.uristDataSet1.delo_deb_prom); }
            }


        }

        private void gridControl1_Click(object sender, EventArgs e)
        {

        }

        private void advBandedGridView1_RowStyle(object sender, DevExpress.XtraGrid.Views.Grid.RowStyleEventArgs e)
        {
            GridView View = sender as GridView;
            if (e.RowHandle >= 0)
            {
                DateTime? r;
                DateTime? s;
                DateTime? p;
                int stat;
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["status"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["status"]) == "")))
                {
                    stat= 0;
                }
                else
                {
                    stat = (int)(View.GetRowCellValue(e.RowHandle, View.Columns["status"]));
                }
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]) == "")))
                {
                    p = null;
                }
                else
                {
                    p = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]));
                }
                if (p <= DateTime.Now && View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]) is DBNull && stat==1)
                {
                    e.Appearance.BackColor = Color.Crimson;
                    e.Appearance.ForeColor = Color.Black;
                }
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]) == "")))
                {
                    r = null;
                }
                else
                {
                    r = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]));
                }
                if (r > p && stat==1)
                {
                    e.Appearance.BackColor = Color.FromArgb(0xFF, 0x99, 0xCC);
                    e.Appearance.ForeColor = Color.Black;
                }
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]) == "")))
                {
                    s = null;
                }
                else
                {
                    s = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]));


                    if (s >= DateTime.Now && stat==1)
                    {
                        e.Appearance.BackColor = Color.FromArgb(0xFF, 0xFF, 0x99);
                        e.Appearance.ForeColor = Color.Black;
                    }

                }
            }
        }

        private void advBandedGridView2_RowStyle(object sender, RowStyleEventArgs e)
        {
            GridView View = sender as GridView;
            if (e.RowHandle >= 0)
            {
                DateTime? r;
                DateTime? s;
                DateTime? p;
                int stat;
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["status"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["status"]) == "")))
                {
                    stat = 0;
                }
                else
                {
                    stat = (int)(View.GetRowCellValue(e.RowHandle, View.Columns["status"]));
                }
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]) == "")))
                {
                    p = null;
                }
                else
                {
                    p = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_peredachi_sud"]));
                }
                if (p <= DateTime.Now && View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]) is DBNull && stat == 1)
                {
                    e.Appearance.BackColor = Color.Crimson;
                    e.Appearance.ForeColor = Color.Black;
                }
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]) == "")))
                {
                    r = null;
                }
                else
                {
                    r = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["real_date_peredachi_sud"]));
                }
                if (r > p && stat == 1)
                {
                    e.Appearance.BackColor = Color.FromArgb(0xFF, 0x99, 0xCC);
                    e.Appearance.ForeColor = Color.Black;
                }
                if ((View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]) is DBNull || (View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]) == "")))
                {
                    s = null;
                }
                else
                {
                    s = (DateTime)(View.GetRowCellValue(e.RowHandle, View.Columns["date_rassmoter"]));


                    if (s >= DateTime.Now && stat == 1)
                    {
                        e.Appearance.BackColor = Color.FromArgb(0xFF, 0xFF, 0x99);
                        e.Appearance.ForeColor = Color.Black;
                    }

                }
            }
        }
    }
}