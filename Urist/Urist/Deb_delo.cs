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
    public partial class Deb_delo : DevExpress.XtraEditors.XtraForm
    {
        public Deb_delo()
        {
            InitializeComponent();
        }

        private void Deb_delo_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'uristDataSet1.delo' table. You can move, or remove it, as needed.
           
            // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
            this.sudTableAdapter.Fill(this.uristDataSet1.sud);
            // TODO: This line of code loads data into the 'uristDataSet1.sud_delo' table. You can move, or remove it, as needed.
            this.sud_deloTableAdapter.Fill(this.uristDataSet1.sud_delo);
            // TODO: This line of code loads data into the 'uristDataSet1.delo_deb_prom' table. You can move, or remove it, as needed.
            this.delo_deb_promTableAdapter.Fill(this.uristDataSet1.delo_deb_prom);
            // TODO: This line of code loads data into the 'uristDataSet1.sprStatus' table. You can move, or remove it, as needed.
            this.sprStatusTableAdapter.Fill(this.uristDataSet1.sprStatus);
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);
            // TODO: This line of code loads data into the 'uristDataSet1.delo_deb_byt' table. You can move, or remove it, as needed.
            //this.delo_deb_bytTableAdapter.Fill(this.uristDataSet1.delo_deb_byt);
            // TODO: This line of code loads data into the 'uristDataSet1.sprSlujba' table. You can move, or remove it, as needed.
            this.sprSlujbaTableAdapter.FillByRes(this.uristDataSet1.sprSlujba);
            
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
            Globals.sdate = dateEdit6.Text;
            Globals.podate = dateEdit5.Text;
            if (Globals.id_res == 1)
            {
                this.delo_deb_bytTableAdapter.FillBySever(this.uristDataSet1.delo_deb_byt, Globals.sdate, Globals.podate);
            }
            else if (Globals.id_res == 14)
            {
                this.delo_deb_bytTableAdapter.FillByTalas(this.uristDataSet1.delo_deb_byt, Globals.sdate, Globals.podate);
            }
            else
            {
                if (Globals.id_zavis == 1 || Globals.id_zavis == 14)

                { this.delo_deb_bytTableAdapter.FillByRes(this.uristDataSet1.delo_deb_byt, Globals.id_res, Globals.sdate, Globals.podate); }
                else { this.delo_deb_bytTableAdapter.FillByUchastok(this.uristDataSet1.delo_deb_byt, Globals.sdate, Globals.podate, Globals.id_res.ToString()); }
            }
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
                        if (Globals.id_zavis == 1 || Globals.id_zavis == 14)

                        { this.delo_deb_bytTableAdapter.FillByRes1(this.uristDataSet1.delo_deb_byt, Globals.id_res); }
                        else { this.delo_deb_bytTableAdapter.FillByUchastok1(this.uristDataSet1.delo_deb_byt, Globals.id_res.ToString()); }
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
                        if (Globals.id_zavis == 1 || Globals.id_zavis == 14)

                        { this.delo_deb_promTableAdapter.FillByUchastok1(this.uristDataSet1.delo_deb_prom, Globals.id_res.ToString()); }
                    }
                }


            }
            catch { }
        }

        private void simpleButton17_Click(object sender, EventArgs e)
        {
            try
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
                    if (Globals.id_zavis == 1 || Globals.id_zavis == 14)

                    { this.delo_deb_bytTableAdapter.FillByRes1(this.uristDataSet1.delo_deb_byt, Globals.id_res); }
                    else { this.delo_deb_bytTableAdapter.FillByUchastok1(this.uristDataSet1.delo_deb_byt, Globals.id_res.ToString()); }
                }


            }
            catch { }
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
                if (Globals.id_zavis == 1 || Globals.id_zavis == 14)

                { this.delo_deb_promTableAdapter.FillByUchastok1(this.uristDataSet1.delo_deb_prom, Globals.id_res.ToString()); }
            }
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            Globals.sdate = dateEdit2.Text;
            Globals.podate = dateEdit1.Text;
            if (Globals.id_res == 1)
            {
                this.delo_deb_promTableAdapter.FillBySever(this.uristDataSet1.delo_deb_prom, Globals.sdate, Globals.podate);
            }
            else if (Globals.id_res == 14)
            {
                this.delo_deb_promTableAdapter.FillByTalas(this.uristDataSet1.delo_deb_prom, Globals.sdate, Globals.podate);
            }
            else
            {
                if (Globals.id_zavis == 1 || Globals.id_zavis == 14)

                {  this.delo_deb_promTableAdapter.FillByUchastok(this.uristDataSet1.delo_deb_prom, Globals.sdate, Globals.podate, Globals.id_res.ToString()); }
            }
        }

        private void pictureEdit2_Click(object sender, EventArgs e)
        {
            ShowGridPreview(gridControl2);
        }

        private void toolStripMenuItem1_Click(object sender, EventArgs e)
        {
            if(xtraTabControl1.SelectedTabPage==xtraTabPage1)
            {
                try
                {
                    Globals.deb_delo_status = (int)((DataRowView)(delodebbytBindingSource.Current)).Row["status"];
                    if (Globals.deb_delo_status == 1)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebbytBindingSource.Current)).Row["cspot"];
                        this.deloTableAdapter.FillByCspot(this.uristDataSet1.delo, Globals.cspot);
                        if(deloBindingSource.Count>0)
                        {
                            if (((DataRowView)(deloBindingSource.Current)).Row["status_sud_peredachi"] is DBNull)
                        {
                            DateTime date = (DateTime)((DataRowView)(deloBindingSource.Current)).Row["date_peredachi_sud"];
                            MessageBox.Show("По абоненту ведется работа по сбору данных. Дата передачи в суд "+date.ToString()); 
                         }
                        
                        else 
                        { 
    
                        Form dd = new Podrobno();
                        dd.Show();}}
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
                    if (Globals.deb_delo_status == 1)
                    {
                        Globals.cspot = (string)((DataRowView)(delodebpromBindingSource.Current)).Row["cspot"];
                        this.deloTableAdapter.FillByCspot(this.uristDataSet1.delo, Globals.cspot);
                        if (deloBindingSource.Count > 0)
                        {
                            if (((DataRowView)(deloBindingSource.Current)).Row["status_sud_peredachi"] is DBNull)
                            {
                                DateTime date = (DateTime)((DataRowView)(deloBindingSource.Current)).Row["date_peredachi_sud"];
                                MessageBox.Show("По абоненту ведется работа по сбору данных. Дата передачи в суд " + date.ToString());
                            }
                            else
                            {

                                Form dd = new Podrobno();
                                dd.Show();
                            }
                        }
                    }
                    else { MessageBox.Show("Абонент не принят в дело, по этому абоненту нет истории"); }



                }
                catch { MessageBox.Show("По этому абоненту нет истории"); }
            }
        }
    }
}