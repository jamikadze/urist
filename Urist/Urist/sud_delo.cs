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
using urist;
using System.IO;

namespace Urist
{
    public partial class sud_delo : DevExpress.XtraEditors.XtraForm
    {
        bool but_clik = false;
        bool but_clik1 = false;
        
        public sud_delo()
        {
            InitializeComponent();
        }

        private void sudBindingNavigatorSaveItem_Click(object sender, EventArgs e)
        {
            this.Validate();
            this.sudBindingSource.EndEdit();
            this.tableAdapterManager.UpdateAll(this.uristDataSet1);

        }

        private void sud_delo_Load(object sender, EventArgs e)
        {
           
            if (Globals.id_sud_spr == 2) {
                if (Globals.prom_byt == 2)
                {
                    textEdit4.Enabled = true;
                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit5.Enabled = true;
                    textEdit1.Text = Convert.ToString(Globals.yenergya);
                    textEdit1.Enabled = true;

                }
                else
                {

                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakta);
                    textEdit3.Text = Convert.ToString(Globals.ypenia);
                    textEdit1.Text = Convert.ToString(Globals.yenergya);
                    textEdit1.Enabled = true;
                }
                
            // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
            this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
            // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
            comboBoxEdit1.Visible = false;
            lookUpEdit1.Enabled = false;
            lookUpEdit1.EditValue = 2;
            groupControl2.Enabled = false;
            }
            if (Globals.id_sud_spr == 1)
            {
                this.Text = "Судебные дела";
                // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
                this.sprStatusSudTableAdapter.FillBy(this.uristDataSet1.sprStatusSud);
                // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
                comboBoxEdit1.Visible = true;
                lookUpEdit1.Enabled = true;
                checkEdit2.Checked = true;
                if (Globals.prom_byt == 2)
                {
                    textEdit4.Enabled = true;
                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit5.Enabled = true;
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;

                }
                else
                {

                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                }
            }
        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            Globals.en_delegate1();
            Close();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (Globals.id_sud_spr == 2)
            {
                Globals.click = 2;
                string nomer;
                if (vh_nomer_sudaTextEdit.Text == "")
                { nomer = null; }
                else { nomer = vh_nomer_sudaTextEdit.Text; }
                DateTime? ot_date;
                if (dateEdit4.Text == "")
                { ot_date = null; }
                else { ot_date = Convert.ToDateTime(dateEdit4.Text); }
                DateTime? date_rassm;
                if (dateEdit3.Text == "")
                { date_rassm = null; }
                else { date_rassm = Convert.ToDateTime(dateEdit3.Text); }
                int vzyskano;
                if (checkEdit1.Checked==true)
                { vzyskano = 1; }
                else
                { vzyskano = 0; }
                int oplacheno;
                if (checkEdit2.Checked == true)
                { oplacheno = 1; }
                else
                { oplacheno = 0; }

                if (!but_clik)
                {
                    but_clik = true;
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 2, DateTime.Now, 1, date_rassm, memoEdit1.Text,null,null,null,null,null);
                    this.deloTableAdapter1.UpdateSudPeredStatus(Globals.id_delo);
                    Globals.id_sud = id_sud;
                    MessageBox.Show("Дело успешно перемещена, далее отсканируйте для  прикрепления файлов к этому документу!!!");

                    simpleButton2.Enabled = false;
                    simpleButton3.Enabled = false;
                    MainFrame mf = new MainFrame();
                    mf.Show();

                }
                else
                {
                    simpleButton2.Enabled = false;
                    MainFrame mf = new MainFrame();
                    mf.Show();
                }
            }

            if (Globals.id_sud_spr == 1)
            {
                Globals.click = 3;
                string nomer;
                if (vh_nomer_sudaTextEdit.Text == "")
                { nomer = null; }
                else { nomer = vh_nomer_sudaTextEdit.Text; }
                DateTime? ot_date;
                if (dateEdit4.Text == "")
                { ot_date = null; }
                else { ot_date = Convert.ToDateTime(dateEdit4.Text); }
                DateTime? date_rassm;
                if (dateEdit3.Text == "")
                { date_rassm = null; }
                else { date_rassm = Convert.ToDateTime(dateEdit3.Text); }
                double yenergy;
                if (textEdit6.Text=="")
                { yenergy = 0; }
                else
                { yenergy=Convert.ToDouble(textEdit6.Text);}
                double ypenya;
                if (textEdit4.Text == "")
                { ypenya = 0; }
                else
                { ypenya = Convert.ToDouble(textEdit4.Text); }
                double yakt2;
                if (textEdit5.Text == "")
                { yakt2 = 0; }
                else
                { yakt2 = Convert.ToDouble(textEdit5.Text); }
                bool vzyskano;
                if (checkEdit1.Checked == true)
                { vzyskano = true; }
                else
                { vzyskano = false; }
                bool oplacheno;
                if (checkEdit2.Checked == true)
                { oplacheno = true; }
                else
                { oplacheno = false; }
                if (!but_clik)
                {
                    but_clik = true;
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, Convert.ToInt32(Globals.status_sud), DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano);
                   // this.deloTableAdapter1.UpdateSudPeredStatus(Globals.id_delo);
                    Globals.id_sud = id_sud;
                    MessageBox.Show("Успешно сохранено, далее отсканируйте для  прикрепления файлов к этому документу!!!");

                    simpleButton2.Enabled = false;
                    simpleButton3.Enabled = false;
                    MainFrame mf = new MainFrame();
                    mf.Show();

                }
                else
                {
                    simpleButton2.Enabled = false;
                    MainFrame mf = new MainFrame();
                    mf.Show();
                }
            }



        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
           
            string nomer;
            if (vh_nomer_sudaTextEdit.Text == "")
            { nomer = null; }
            else { nomer = vh_nomer_sudaTextEdit.Text; }
            DateTime? ot_date;
            if (dateEdit4.Text == "")
            { ot_date = null; }
            else { ot_date = Convert.ToDateTime(dateEdit4.Text); }
            DateTime? date_rassm;
            if (dateEdit3.Text == "")
            { date_rassm = null; }
            else { date_rassm = Convert.ToDateTime(dateEdit3.Text); }
            double yenergy;
            if (textEdit6.Text == "")
            { yenergy = 0; }
            else
            { yenergy = Convert.ToDouble(textEdit6.Text); }
            double ypenya;
            if (textEdit4.Text == "")
            { ypenya = 0; }
            else
            { ypenya = Convert.ToDouble(textEdit4.Text); }
            double yakt2;
            if (textEdit5.Text == "")
            { yakt2 = 0; }
            else
            { yakt2 = Convert.ToDouble(textEdit5.Text); }
            bool vzyskano;
            if (checkEdit1.Checked == true)
            { vzyskano = true; }
            else
            { vzyskano = false; }
            bool oplacheno;
            if (checkEdit2.Checked == true)
            { oplacheno = true; }
            else
            { oplacheno = false; }
            if (Globals.id_sud_spr == 2)
            { Globals.click = 2;
                if (!but_clik1)
                {
                    but_clik1 = true;
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 2, DateTime.Now, 1, date_rassm, memoEdit1.Text,null,null,null,null,null);
                    this.deloTableAdapter1.UpdateSudPeredStatus(Globals.id_delo);
                    Globals.id_sud = id_sud;
                    MessageBox.Show("Дело успешно перемещена, далее прикрепите файл к этому документу!!!");

                    simpleButton2.Enabled = false;
                    simpleButton3.Enabled = false;
                    ofdInput.ShowDialog();

                }
                else
                {
                    simpleButton2.Enabled = false;
                    ofdInput.ShowDialog();
                }
            }

            if (Globals.id_sud_spr == 1)
            {
                Globals.click = 3;
                if (!but_clik1)
                {
                    but_clik1 = true;
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, Convert.ToInt32(Globals.status_sud), DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano);
                   // this.deloTableAdapter1.UpdateSudPeredStatus(Globals.id_delo);
                    Globals.id_sud = id_sud;
                    MessageBox.Show("Успешно сохранено, далее прикрепите  файл к этому документу!!!");

                    simpleButton2.Enabled = false;
                    simpleButton3.Enabled = false;
                    ofdInput.ShowDialog();

                }
                else
                {
                    simpleButton2.Enabled = false;
                    ofdInput.ShowDialog();
                }
            }
        }

        private void simpleButton3_Click(object sender, EventArgs e)
        {
            string nomer;
            if (vh_nomer_sudaTextEdit.Text == "")
            { nomer = null; }
            else { nomer = vh_nomer_sudaTextEdit.Text; }
            DateTime? ot_date;
            if (dateEdit4.Text=="")
            { ot_date = null; }
            else { ot_date = Convert.ToDateTime(dateEdit4.Text); }
            DateTime? date_rassm;
            if (dateEdit3.Text == "")
            { date_rassm = null; }
            else { date_rassm = Convert.ToDateTime(dateEdit3.Text); }
            double yenergy;
            if (textEdit6.Text == "")
            { yenergy = 0; }
            else
            { yenergy = Convert.ToDouble(textEdit6.Text); }
            double ypenya;
            if (textEdit4.Text == "")
            { ypenya = 0; }
            else
            { ypenya = Convert.ToDouble(textEdit4.Text); }
            double yakt2;
            if (textEdit5.Text == "")
            { yakt2 = 0; }
            else
            { yakt2 = Convert.ToDouble(textEdit5.Text); }
            bool vzyskano;
            if (checkEdit1.Checked == true)
            { vzyskano = true; }
            else
            { vzyskano = false; }
            bool oplacheno;
            if (checkEdit2.Checked == true)
            { oplacheno = true; }
            else
            { oplacheno = false; }
            if (Globals.id_sud_spr == 2)
               {
           try
            {
                int? id_sud = 0;
                queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 2, DateTime.Now, 1, date_rassm, memoEdit1.Text,null,null,null,null,null);
                this.deloTableAdapter1.UpdateSudPeredStatus(Globals.id_delo);
                Globals.id_sud = id_sud;
                MessageBox.Show("Успешно сохранено!!!");
               

            }
           catch(Exception ex)
            {
               MessageBox.Show("Данные не сохранились!!!"); 
                Globals.en_delegate1();
                Close();
            }
            }
            if (Globals.id_sud_spr == 1)
            {
                try
                {
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, Convert.ToInt32(Globals.status_sud), DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2,  oplacheno, vzyskano);
                    //this.deloTableAdapter1.UpdateSudPeredStatus(Globals.id_delo);
                    Globals.id_sud = id_sud;
                    MessageBox.Show("Успешно сохранено!!!");


                }
                catch (Exception ex)
                {
                    MessageBox.Show("Данные не сохранились!!!");
                    Globals.en_delegate1();
                    Close();
                }
            }

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
                    if (Globals.click == 2)
                    {
                        this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, file_name, null, Globals.id_user, 1, Globals.id_sud);
                    }
                    if (Globals.click == 3)
                    {
                        this.queriesTableAdapter1.insert_file(Globals.id_delo1, outlen, file_name, null, Globals.id_user, Convert.ToInt32(Globals.status_sud), Globals.id_sud);
                    }
                    MessageBox.Show("Файл успешно прикреплен ;) ");
                }

            }

            catch (IOException ex)
            {
                MessageBox.Show("Данные не сохранились");
            }
        }

        private void lookUpEdit1_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            if (e.CloseMode == DevExpress.XtraEditors.PopupCloseMode.Normal)
            {

                Globals.status_sud = (sender as DevExpress.XtraEditors.LookUpEdit).GetColumnValue("id");
                

            }
        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit ch = sender as CheckEdit;
            if (ch.Name == "checkEdit1")
            { checkEdit2.Checked = false;  }
            else { checkEdit1.Checked = false; }

            if (ch.Name == "checkEdit2")
            {
               
                checkEdit1.Checked = false;
               
            }
            else { checkEdit2.Checked = false; }
        }
    }
}