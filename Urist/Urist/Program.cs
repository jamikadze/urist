using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using DevExpress.LookAndFeel;

namespace Urist
{
    static class Program
    {
        private static bool InstanceExists()
        {
            bool createdNew;
            new System.Threading.Mutex(false, "OneInstanceApplication", out createdNew);
            return (!createdNew);
        }
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            if (InstanceExists())
            {
                MessageBox.Show("Программа уже запущенна");
                return;
            }
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);

            DevExpress.Skins.SkinManager.EnableFormSkins();
            DevExpress.UserSkins.BonusSkins.Register();
            UserLookAndFeel.Default.SetSkinStyle("London Liquid Sky");
            Application.Run(new Authentification());
            if (Globals.iii == 0)
            { Application.Run(new Main()); }
            if (Globals.iii == 1)
            { Application.Run(new Main1()); }
            if (Globals.iii == 2)
            { Application.Run(new Main2()); }
            if (Globals.iii == 3)
            { Application.Run(new EditDelo()); }
            //
        }
    }
}
