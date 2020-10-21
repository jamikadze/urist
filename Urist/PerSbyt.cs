using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using urist;
using System.IO;

namespace Urist
{
    public partial class PerSbyt : DevExpress.XtraEditors.XtraForm
    {
        bool but_clik = false;
        bool but_clik1 = false;
        int click_black1;
        public PerSbyt(int click_black)

        {
            click_black1 = click_black;
            InitializeComponent();
        }

        private void simpleButton1_Click(object sender, EventArgs e)
        {
            if (!but_clik)
            {
                but_clik = true;
                Globals.click = 1;
                if (click_black1 == 2)
                {

                    if (Globals.stat == 2)
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_delo1);
                        if (Globals.prombyt == 1)
                        { this.delo_deb_bytTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                        else
                        { this.delo_deb_promTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                        MessageBox.Show("Успешно сохранено!!!");
                    }
                    else
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                        if (Globals.prombyt == 1)
                        { this.delo_deb_bytTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                        else
                        { this.delo_deb_promTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                        MessageBox.Show("Успешно сохранено!!!");
                    }
                }
                if (click_black1 == 1)
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                    if (Globals.prombyt == 1)
                    {
                        this.delo_deb_bytTableAdapter1.UpdateSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    else
                    {
                        this.delo_deb_promTableAdapter1.UpdateSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    MessageBox.Show("Успешно сохранено!!!");
                }
                if (click_black1 == 3)
                {
                    try
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(5, memoEdit1.Text,Globals.id_user, Globals.id_delo );
                        if (Globals.prombyt == 1)
                        {
                            this.delo_deb_bytTableAdapter1.UpdateSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        else
                        {
                            this.delo_deb_promTableAdapter1.UpdateSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        MessageBox.Show("Успешно сохранено!!!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Данные не сохранились!!!");
                        if (click_black1 == 3)
                        {
                            Globals.black_list();
                            Close();
                        }
                        else
                        {
                            Globals.sbyt();
                            Close();
                        }
                    }

                }
                if (click_black1 == 7)
                {
                    try
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                        if (Globals.prombyt == 1)
                        {
                            this.delo_deb_bytTableAdapter1.UpdateSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        else
                        {
                            this.delo_deb_promTableAdapter1.UpdateSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        MessageBox.Show("Успешно сохранено!!!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Данные не сохранились!!!");
                        if (click_black1 == 7)
                        {
                            Globals.black_list();
                            Close();
                        }
                        else
                        {
                            Globals.sbyt();
                            Close();
                        }
                    }

                }
                MessageBox.Show("Дело успешно передано на проработку/перемещена, далее отсканируйте для  прикрепления файлов к этому документу!!!");
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

        private void simpleButton4_Click(object sender, EventArgs e)
        {
            if (click_black1 == 3 || click_black1 == 7)
            {
                Globals.black_list();
                Close(); }
            else { 
            Globals.sbyt();
            Close();}
        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            Globals.click = 1;
            if (!but_clik1)
            {
                but_clik1 = true;

                if(click_black1==2)
                {

                    if (Globals.stat == 2)
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_delo1);
                        if (Globals.prombyt == 1)
                        { this.delo_deb_bytTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                        else
                        { this.delo_deb_promTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                        MessageBox.Show("Успешно сохранено!!!");
                    }
                    else
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                        if (Globals.prombyt == 1)
                        { this.delo_deb_bytTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                        else
                        { this.delo_deb_promTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                        MessageBox.Show("Успешно сохранено!!!");
                    }
                 }
                if (click_black1 == 1)
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(3, memoEdit1.Text,  Globals.id_user,Globals.id_delo);
                    if (Globals.prombyt == 1)
                    {
                        this.delo_deb_bytTableAdapter1.UpdateSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    else
                    {
                        this.delo_deb_promTableAdapter1.UpdateSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    MessageBox.Show("Успешно сохранено!!!");
                }
                if (click_black1 == 3)
                {
                    try
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(5, memoEdit1.Text, Globals.id_user,Globals.id_delo );
                        if (Globals.prombyt == 1)
                        {
                            this.delo_deb_bytTableAdapter1.UpdateSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        else
                        {
                            this.delo_deb_promTableAdapter1.UpdateSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        MessageBox.Show("Успешно сохранено!!!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Данные не сохранились!!!");
                        if (click_black1 == 3)
                        {
                            Globals.black_list();
                            Close();
                        }
                        else
                        {
                            Globals.sbyt();
                            Close();
                        }
                    }

                }
                if (click_black1 == 7)
                {
                    try
                    {
                        this.deloTableAdapter1.UpdateStatusSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                        if (Globals.prombyt == 1)
                        {
                            this.delo_deb_bytTableAdapter1.UpdateSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        else
                        {
                            this.delo_deb_promTableAdapter1.UpdateSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                        }
                        MessageBox.Show("Успешно сохранено!!!");
                    }
                    catch (Exception ex)
                    {
                        MessageBox.Show("Данные не сохранились!!!");
                        if (click_black1 == 7)
                        {
                            Globals.black_list();
                            Close();
                        }
                        else
                        {
                            Globals.sbyt();
                            Close();
                        }
                    }

                }
                MessageBox.Show("Дело успешно передано на проработку, далее прикрепите файлы к этому документу!!!");
               

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

        private void simpleButton3_Click(object sender, EventArgs e)
        {
                if(click_black1==2)
                { 
            try
            {
                if (Globals.stat == 2)
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_delo1);
                    if (Globals.prombyt == 1)
                    { this.delo_deb_bytTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                    else
                    { this.delo_deb_promTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                    MessageBox.Show("Успешно сохранено!!!");
                }
                else
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                    if (Globals.prombyt == 1)
                    { this.delo_deb_bytTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }
                    else
                    { this.delo_deb_promTableAdapter1.UpdateSbyt(2, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); }

                    MessageBox.Show("Успешно сохранено!!!");
                }


            }
            catch (Exception ex)
            {
                MessageBox.Show("Данные не сохранились!!!");
                if (click_black1 == 3)
                {
                    Globals.black_list();
                    Close();
                }
                else
                {
                    Globals.sbyt();
                    Close();
                }
            }}

            if(click_black1==1)
            { try
            {
                this.deloTableAdapter1.UpdateStatusSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                if (Globals.prombyt == 1)
                {
                    this.delo_deb_bytTableAdapter1.UpdateSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt); 
                   
                }
                else
                {
                    this.delo_deb_promTableAdapter1.UpdateSbyt(3, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);
                    
                }
                MessageBox.Show("Успешно сохранено!!!");
            }
            catch (Exception ex)
            {
                MessageBox.Show("Данные не сохранились!!!");
                if (click_black1 == 3)
                {
                    Globals.black_list();
                    Close();
                }
                else
                {
                    Globals.sbyt();
                    Close();
                }
            }
               
            }
            if (click_black1 == 3)
            {
                try
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                    if (Globals.prombyt == 1)
                    {
                        this.delo_deb_bytTableAdapter1.UpdateSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    else
                    {
                        this.delo_deb_promTableAdapter1.UpdateSbyt(5, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    MessageBox.Show("Успешно сохранено!!!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Данные не сохранились!!!");
                    if (click_black1 == 3)
                    {
                        Globals.black_list();
                        Close();
                    }
                    else
                    {
                        Globals.sbyt();
                        Close();
                    }
                }

            }
            if (click_black1 == 7)
            {
                try
                {
                    this.deloTableAdapter1.UpdateStatusSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_delo);
                    if (Globals.prombyt == 1)
                    {
                        this.delo_deb_bytTableAdapter1.UpdateSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    else
                    {
                        this.delo_deb_promTableAdapter1.UpdateSbyt(7, memoEdit1.Text, Globals.id_user, Globals.id_prom_byt);

                    }
                    MessageBox.Show("Успешно сохранено!!!");
                }
                catch (Exception ex)
                {
                    MessageBox.Show("Данные не сохранились!!!");
                    if (click_black1 == 7)
                    {
                        Globals.black_list();
                        Close();
                    }
                    else
                    {
                        Globals.sbyt();
                        Close();
                    }
                }

            }


        }

        private void PerSbyt_Load(object sender, EventArgs e)
        {
            if (click_black1 == 1)
            {
                this.Text = "Обоснование для перемещения абонента в папку Безнадежный абоент";
            }
            if (click_black1 == 3)
            {
                this.Text = "Обоснование для перемещения абонента обратно юрисконсультанту";
            }
            if (click_black1 == 7)
            {
                this.Text = "Обоснование для списания абонента ";
            }
        }
    }
}