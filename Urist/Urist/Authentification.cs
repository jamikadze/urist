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
using Urist.Properties;
using System.IO;

namespace Urist
{
    public partial class Authentification : DevExpress.XtraEditors.XtraForm
    {
        public Authentification()
        {
            InitializeComponent();
        }

        private void Authentification_Load(object sender, EventArgs e)
        {
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.Fill(this.uristDataSet1.users_programm);

        }

        private void simpleButton2_Click(object sender, EventArgs e)
        {
            try
            {
                //this.usersTableAdapter.Fill("jjj", "jjj");
                this.users_programmTableAdapter.FillByLogin(this.uristDataSet1.users_programm, textBoxLogin.Text.Trim(), textBoxPass.Text.Trim());
                if (uristDataSet1 != null && uristDataSet1.users_programm.Count() > 0)
                {
                 //   Settings.Default.UserNowName = textBoxLogin.Text;
                    Globals.id_user = (int)((DataRowView)users_programmBindingSource.Current).Row["user_id"];
                    Globals.name_user = (string)((DataRowView)users_programmBindingSource.Current).Row["fio"];

                    Globals.id_slujbi = (int)((DataRowView)users_programmBindingSource.Current).Row["id_slujbi"];
                    Globals.prava_mod_spr_res = (int)((DataRowView)users_programmBindingSource.Current).Row["prava_mod_spravochnika"];
                    Globals.id_doljnost = (int)((DataRowView)users_programmBindingSource.Current).Row["id_doljnost"];

                    if (Globals.id_doljnost == 2)
                    {
                        Globals.iii = 3;
                        //    Form main=new Main();
                        //main.Show(); 
                        Settings.Default.Save();
                        //Application.Run(new Main());
                        Close();
                    }
                    else { 
                    if (Globals.id_slujbi == 37)
                    {
                        Globals.iii = 0;
                        //    Form main=new Main();
                        //main.Show(); 
                        Settings.Default.Save();
                        //Application.Run(new Main());
                        Close();
                    }
                    else
                    {
                        if (Globals.id_slujbi == 38)
                        {
                            Globals.iii = 2;
                            //    Form main=new Main();
                            //main.Show(); 
                            Settings.Default.Save();
                            //Application.Run(new Main());
                            Close();
                        }
                        else
                        {
                            Globals.iii = 1;
                            Settings.Default.Save();
                            //Application.Run(new Main());
                            Close();
                        }

                    }

                }}
                else
                {
                    textBoxLogin.SelectAll();
                    //textBoxPass.();
                    textBoxLogin.Focus();
                    labelControl3.Visible = true;
                    pictureBox1.Visible = false;
                }

            }

            catch (System.Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(ex.Message);
            }
        }

     
        private void simpleButton1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void textBoxPass_KeyDown(object sender, KeyEventArgs e)
        {
            if (e.KeyCode == Keys.Enter)
            {
                simpleButton2_Click(null, null);
            }
        }
    }
}