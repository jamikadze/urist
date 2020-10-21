using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using urist;

namespace Urist
{
    public partial class Bankrot : Form
    {
        public Bankrot()
        {
            InitializeComponent();
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            try
            {
                if (Globals.stat == 2)
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_delo1);
                    if (Globals.prombyt == 1)
                    { this.delo_deb_bytTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                    else
                    { this.delo_deb_promTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                    MessageBox.Show("Успешно сохранено!!!");
                }
                else { 
                this.deloTableAdapter1.UpdateStatusSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                if (Globals.prombyt == 1)
                {
                    this.delo_deb_bytTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                }
                else
                {
                    this.delo_deb_promTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                }
                }
                MessageBox.Show("Успешно сохранено!!!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Данные не сохранились!!!");
                Globals.bankrot();
                    Close();
               
            }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            Globals.bankrot();
            Close();
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            if (Globals.stat == 2)
            {
                this.deloTableAdapter1.UpdateStatusSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_delo1);
                if (Globals.prombyt == 1)
                { this.delo_deb_bytTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                else
                { this.delo_deb_promTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                MessageBox.Show("Успешно сохранено!!!");
            }
            else
            {
                this.deloTableAdapter1.UpdateStatusSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                if (Globals.prombyt == 1)
                {
                    this.delo_deb_bytTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                }
                else
                {
                    this.delo_deb_promTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                }
            }
            
            MessageBox.Show("Успешно сохранено!!!");
            simpleButton2.Enabled = false;
            simpleButton3.Enabled = false;
            ofdInput.ShowDialog();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (Globals.stat == 2)
            {
                this.deloTableAdapter1.UpdateStatusSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_delo1);
                if (Globals.prombyt == 1)
                { this.delo_deb_bytTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                else
                { this.delo_deb_promTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                MessageBox.Show("Успешно сохранено!!!");
            }
            else
            {
                this.deloTableAdapter1.UpdateStatusSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                if (Globals.prombyt == 1)
                {
                    this.delo_deb_bytTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                }
                else
                {
                    this.delo_deb_promTableAdapter1.UpdateSbyt(8, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                }
            }
            MessageBox.Show("Успешно сохранено!!!");
            simpleButton2.Enabled = false;
            simpleButton3.Enabled = false;
            MainFrame mf = new MainFrame();
            mf.Show();
        }
    }
}
