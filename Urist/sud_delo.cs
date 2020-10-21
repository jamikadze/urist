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
        int status_sud;
        public sud_delo(int status)
        {
            status_sud = status;
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

            if (Globals.id_sud_spr == 2)
            {
                if (Globals.prom_byt == 2)
                {
                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakta);
                    textEdit3.Text = Convert.ToString(Globals.ypenia);
                    textEdit1.Text = Convert.ToString(Globals.yenergya);
                    textEdit1.Enabled = true;
                    label3.Text = "э/э+акты";
                    label4.Text = "пени";
                    label5.Text = "НДС";

                }
                else
                {

                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakta);
                    textEdit3.Text = Convert.ToString(Globals.ypenia);
                    textEdit1.Text = Convert.ToString(Globals.yenergya);
                    textEdit1.Enabled = true;
                    label3.Text = "Деб.э/э(ост)";
                    label4.Text = "Деб.пеня(ост)";
                    label5.Text = "Деб.акт(ост)";

                }

                // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
                this.sprStatusSudTableAdapter.Fill(this.uristDataSet1.sprStatusSud);
                // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
                comboBoxEdit1.Visible = false;
                lookUpEdit1.Enabled = false;
                lookUpEdit1.EditValue = 2;
                groupControl2.Enabled = false;
                checkEdit3.Enabled = false;
            }
            if (Globals.id_sud_spr == 1)
            {
                this.Text = "Судебные дела";
                // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
                this.sprStatusSudTableAdapter.FillBy(this.uristDataSet1.sprStatusSud);
                // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
                comboBoxEdit1.Visible = true;
                lookUpEdit1.Enabled = true;
                checkEdit3.Enabled = true;
                if (Globals.prom_byt == 2)
                {
                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                    label3.Text = "э/э+акты";
                    label4.Text = "пени";
                    label5.Text = "НДС";

                }
                else
                {

                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                    label3.Text = "Деб.э/э(ост)";
                    label4.Text = "Деб.пеня(ост)";
                    label5.Text = "Деб.акт(ост)";
                }

            }

            if (Globals.id_sud_spr == 11)
            {
                this.Text = "Закрыть дело на основе претензии";
                // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
                this.sprStatusSudTableAdapter.FillByZakryt(this.uristDataSet1.sprStatusSud,13);
                // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
                comboBoxEdit1.Visible = true;
                lookUpEdit1.Enabled = false;
                lookUpEdit1.EditValue = 13;
                checkEdit3.Enabled = true;
                vh_nomer_sudaTextEdit.Enabled = false;
                dateEdit3.Enabled = true;
                    comboBoxEdit1.Enabled = false;
                    dateEdit4.Enabled = false;
                  
                if (Globals.prom_byt == 2)
                {
                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                    label3.Text = "э/э+акты";
                    label4.Text = "пени";
                    label5.Text = "НДС";
                }
                else
                {

                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                    label3.Text = "Деб.э/э(ост)";
                    label4.Text = "Деб.пеня(ост)";
                    label5.Text = "Деб.акт(ост)";
                }

            }

            if (Globals.id_sud_spr == 5 || Globals.id_sud_spr == 55)
            {
                this.Text = "Мирное урегилирование- Заключение соглашения (с сроком)";
                // TODO: This line of code loads data into the 'uristDataSet1.sprStatusSud' table. You can move, or remove it, as needed.
                this.sprStatusSudTableAdapter.FillBySogl(this.uristDataSet1.sprStatusSud);
                lookUpEdit1.EditValue = 5;
                // TODO: This line of code loads data into the 'uristDataSet1.sud' table. You can move, or remove it, as needed.
                comboBoxEdit1.Visible = true;
                lookUpEdit1.Enabled = true;
                checkEdit3.Enabled = false;

                if (Globals.prom_byt == 2)
                {
                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                    label3.Text = "э/э+акты";
                    label4.Text = "пени";
                    label5.Text = "НДС";
                }
                else
                {

                    textEdit3.Enabled = true;
                    textEdit2.Enabled = true;
                    textEdit2.Text = Convert.ToString(Globals.yakt - Globals.yakt1);
                    textEdit3.Text = Convert.ToString(Globals.ypeni - Globals.ypeny);
                    textEdit1.Text = Convert.ToString(Globals.yenergy - Globals.yenergy_saldo);
                    textEdit1.Enabled = true;
                    label3.Text = "Деб.э/э(ост)";
                    label4.Text = "Деб.пеня(ост)";
                    label5.Text = "Деб.акт(ост)";
                }
            }

        }

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (Globals.id_sud_spr == 11)
            {
                if (checkEdit3.Checked == true)
                {
                    Globals.pssi();
                    Close();
                }
                else { 
                Globals.sbyt();
                Close();}
            }
            else if (Globals.id_sud_spr == 5 || Globals.id_sud_spr == 55)
            {
                if (checkEdit3.Checked == true)
                {
                    Globals.pssi();
                    Close();
                }
                else { 
                Globals.sogl();
                Close();}
           }
           
            else { if (checkEdit3.Checked == true)
            {
                Globals.pssi();
                Close();
            }
            else
            { 
                Globals.en_delegate1();
            Close();}}
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

                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 2, DateTime.Now, 1, DateTime.Now, memoEdit1.Text, null, null, null, null, null, null, null, null);
                    this.deloTableAdapter1.UpdateSudPeredStatus(2,Globals.id_delo);
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
            if (Globals.id_sud_spr == 55)
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
                if (checkEdit1.Checked == true)
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

                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 5, DateTime.Now, 1, date_rassm, memoEdit1.Text, null, null, null, null, null, null, null, null);
                    this.deloTableAdapter1.UpdateSudPeredStatus(5,Globals.id_delo);
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

            if (Globals.id_sud_spr == 1 || Globals.id_sud_spr == 5)
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
                double yenergy1;
                if (textEdit9.Text == "")
                { yenergy1 = 0; }
                else
                { yenergy1 = Convert.ToDouble(textEdit9.Text); }
                double ypenya1;
                if (textEdit7.Text == "")
                { ypenya1 = 0; }
                else
                { ypenya1 = Convert.ToDouble(textEdit7.Text); }
                double yakt1;
                if (textEdit8.Text == "")
                { yakt1 = 0; }
                else
                { yakt1 = Convert.ToDouble(textEdit8.Text); }



                if (!but_clik)
                {
                    but_clik = true;
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, Convert.ToInt32(Globals.status_sud), DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);
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

            if (Globals.id_sud_spr == 11)
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
                double yenergy1;
                if (textEdit9.Text == "")
                { yenergy1 = 0; }
                else
                { yenergy1 = Convert.ToDouble(textEdit9.Text); }
                double ypenya1;
                if (textEdit7.Text == "")
                { ypenya1 = 0; }
                else
                { ypenya1 = Convert.ToDouble(textEdit7.Text); }
                double yakt1;
                if (textEdit8.Text == "")
                { yakt1 = 0; }
                else
                { yakt1 = Convert.ToDouble(textEdit8.Text); }
                if (!but_clik)
                {
                    but_clik = true;
                    int? id_sud = 0;
                    this.deloTableAdapter1.UpdateSudPeredStatus(2,Globals.id_delo);
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 13, DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);
                  //  queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 11, DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, 0, 0, 0, oplacheno, vzyskano, 0, 0, 0);
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
            double yenergy1;
            if (textEdit9.Text == "")
            { yenergy1 = 0; }
            else
            { yenergy1 = Convert.ToDouble(textEdit9.Text); }
            double ypenya1;
            if (textEdit7.Text == "")
            { ypenya1 = 0; }
            else
            { ypenya1 = Convert.ToDouble(textEdit7.Text); }
            double yakt1;
            if (textEdit8.Text == "")
            { yakt1 = 0; }
            else
            { yakt1 = Convert.ToDouble(textEdit8.Text); }
            if (Globals.id_sud_spr == 2)
            { Globals.click = 2;
                if (!but_clik1)
                {
                    but_clik1 = true;
                    int? id_sud = 0;

                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 2, DateTime.Now, 1, DateTime.Now, memoEdit1.Text, null, null, null, null, null, null, null, null);
                    this.deloTableAdapter1.UpdateSudPeredStatus(2,Globals.id_delo);
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
           
            if (Globals.id_sud_spr == 55)
            { 
                if (dateEdit3.Text != "" && comboBoxEdit1.Text != "" && lookUpEdit1.Text != "")
            {
                Globals.click = 2;
                if (!but_clik1)
                {
                    but_clik1 = true;
                    int? id_sud = 0;

                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 5, DateTime.Now, 1, date_rassm, memoEdit1.Text, null, null, null, null, null, null, null, null);
                    this.deloTableAdapter1.UpdateSudPeredStatus(5,Globals.id_delo);
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
                else { MessageBox.Show("Заполните обьязательные поля ****"); }
            }

            if (Globals.id_sud_spr == 1 || Globals.id_sud_spr == 5)
            {
                
                if (dateEdit3.Text != "" && comboBoxEdit1.Text != "" && lookUpEdit1.Text != "")
            {
                Globals.click = 3;
                if (!but_clik1)
                {
                    but_clik1 = true;
                    int? id_sud = 0;
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, Convert.ToInt32(Globals.status_sud), DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);
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
                else { MessageBox.Show("Заполните обьязательные поля ****"); }
            }

            if (Globals.id_sud_spr == 11)
            {
                 if (dateEdit3.Text != "" && lookUpEdit1.Text != "")
            {
                Globals.click = 3;
                if (!but_clik1)
                {
                    but_clik1 = true;
                    int? id_sud = 0;
                    this.deloTableAdapter1.UpdateSudPeredStatus(2,Globals.id_delo);
                    queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 13, DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);
                   // queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 11, DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, 0, 0, 0, oplacheno, vzyskano,0,0,0);
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
                 else { MessageBox.Show("Заполните обьязательные поля ****"); }
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
            double yenergy1;
            if (textEdit9.Text == "")
            { yenergy1 = 0; }
            else
            { yenergy1 = Convert.ToDouble(textEdit9.Text); }
            double ypenya1;
            if (textEdit7.Text == "")
            { ypenya1 = 0; }
            else
            { ypenya1 = Convert.ToDouble(textEdit7.Text); }
            double yakt1;
            if (textEdit8.Text == "")
            { yakt1 = 0; }
            else
            { yakt1 = Convert.ToDouble(textEdit8.Text); }
            if (Globals.id_sud_spr == 2)
               {
           try
            {
                int? id_sud = 0;
                queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 2, DateTime.Now, 1, DateTime.Now, memoEdit1.Text, null, null, null, null, null, null, null, null);
                this.deloTableAdapter1.UpdateSudPeredStatus(2,Globals.id_delo);
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
           
                if (Globals.id_sud_spr == 55 || Globals.id_sud_spr == 5)
                {
                    if (dateEdit3.Text != "" && comboBoxEdit1.Text != "" && lookUpEdit1.Text != "")
                    {
                    try
                    {
                        if (Globals.id_sud_spr == 55)
                        {
                            int? id_sud = 0;
                            queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 5, DateTime.Now, 1, date_rassm, memoEdit1.Text, null, null, null, null, null, null, null, null);
                            this.deloTableAdapter1.UpdateSudPeredStatus(5, Globals.id_delo);
                            Globals.id_sud = id_sud;
                            MessageBox.Show("Успешно сохранено!!!");
                        }
                        if (Globals.id_sud_spr == 5)
                        {
                            int? id_sud = 0;
                            queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, 5, DateTime.Now, 0, date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);
                            this.deloTableAdapter1.UpdateSudPeredStatus(5, Globals.id_delo1);
                            Globals.id_sud = id_sud;
                            MessageBox.Show("Успешно сохранено!!!");
                        }


                    }
                    catch (Exception ex)
                    {
                        if (Globals.id_sud_spr == 55)
                        {
                            Globals.en_delegate1();
                            Close();
                        }
                        if (Globals.id_sud_spr == 5)
                        {
                            Globals.sogl();
                            Close();
                        }
                    }
                    }
                    else { MessageBox.Show("Заполните обьязательные поля ****"); }

                }

                if (Globals.id_sud_spr == 1)
                {
                    if (dateEdit3.Text != "" && comboBoxEdit1.Text != "" && lookUpEdit1.Text != "")
                    {
                    try
                    {
                        int? id_sud = 0;
                        queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, Convert.ToInt32(Globals.status_sud), DateTime.Now, Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);

                        Globals.id_sud = id_sud;
                        if (checkEdit3.Checked == true)
                        { queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, 17, DateTime.Now.AddMinutes(5), Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, null, null, null, null, null, null, null, null); }
                        MessageBox.Show("Успешно сохранено!!!");


                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Данные не сохранились!!!");
                        if (Globals.id_sud_spr == 1)
                        {
                            Globals.en_delegate1();
                            Close();
                        }


                    }
                    }
                    else { MessageBox.Show("Заполните обьязательные поля ****"); }
                }

                if (Globals.id_sud_spr == 11)
                {
                     if (dateEdit3.Text != ""  && lookUpEdit1.Text != "")
                    {
                    try
                    {
                        int? id_sud = 0;
                        this.deloTableAdapter1.UpdateSudPeredStatus(2, Globals.id_delo);
                        queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 13, DateTime.Now, 0, date_rassm, memoEdit1.Text, yenergy, ypenya, yakt2, oplacheno, vzyskano, yenergy1, ypenya1, yakt1);
                       // queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo, 11, DateTime.Now, 0, date_rassm, memoEdit1.Text, 0, 0, 0, oplacheno, vzyskano, 0, 0, 0);
                        if (checkEdit3.Checked == true)
                        { queriesTableAdapter1.insert_sud(ref id_sud, nomer, ot_date, Globals.id_delo1, 17, DateTime.Now.AddMinutes(5), Convert.ToInt32(comboBoxEdit1.Text), date_rassm, memoEdit1.Text, null, null, null, null, null, null, null, null); }
                        Globals.id_sud = id_sud;
                        MessageBox.Show("Успешно сохранено!!!");


                    }
                    catch (Exception ex)
                    {

                        MessageBox.Show("Данные не сохранились!!!");
                        Globals.sbyt();
                        Close();
                    }
                }  
                     else { MessageBox.Show("Заполните обьязательные поля ****"); }
                }
             
           

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
        }

        private void lookUpEdit1_Closed(object sender, DevExpress.XtraEditors.Controls.ClosedEventArgs e)
        {
            if (e.CloseMode == DevExpress.XtraEditors.PopupCloseMode.Normal)
            {

                Globals.status_sud = (sender as DevExpress.XtraEditors.LookUpEdit).GetColumnValue("id");
                if (Convert.ToInt32(Globals.status_sud) == 9 || Convert.ToInt32(Globals.status_sud) == 12 || Convert.ToInt32(Globals.status_sud) == 16)
                { checkEdit3.Checked = true; }

            }
        }

        private void checkEdit1_CheckedChanged(object sender, EventArgs e)
        {
            CheckEdit ch = sender as CheckEdit;
            
            textEdit4.Enabled=true;
                textEdit5.Enabled=true;
                textEdit6.Enabled = true;
           
               
           
        }

        private void label1_Click(object sender, EventArgs e)
        {

        }

        private void checkEdit1_Click(object sender, EventArgs e)
        {
            textEdit9.Enabled = true;
            textEdit7.Enabled = true;
            textEdit8.Enabled = true;
        }

        private void textEdit6_Enter(object sender, EventArgs e)
        {
            double odin;
            double dva;
            double tri;



            if (textEdit6.Text != "")
            { odin = Convert.ToDouble(textEdit6.Text); }
            else { odin = 0; }
            if (textEdit5.Text != "")
            { dva = Convert.ToDouble(textEdit5.Text); }
            else { dva= 0; }
            if (textEdit4.Text != "")
            { tri = Convert.ToDouble(textEdit4.Text); }
            else { tri = 0; }
            textEdit10.Text = (odin + dva +tri).ToString();
        }

        private void textEdit9_Validated(object sender, EventArgs e)
        {
            double odin;
            double dva;
            double tri;



            if (textEdit9.Text != "")
            { odin = Convert.ToDouble(textEdit9.Text); }
            else { odin = 0; }
            if (textEdit7.Text != "")
            { dva = Convert.ToDouble(textEdit7.Text); }
            else { dva = 0; }
            if (textEdit8.Text != "")
            { tri = Convert.ToDouble(textEdit8.Text); }
            else { tri = 0; }
            textEdit11.Text = (odin + dva + tri).ToString();
        }

       
    }
}