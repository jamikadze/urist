using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraEditors;
using Calendar.NET;

namespace Urist
{
    public partial class Sheduler : DevExpress.XtraEditors.XtraForm
    {
        public Sheduler()
        {
            InitializeComponent();
          
        }

        private void Sheduler_Load(object sender, EventArgs e)
        {

            this.shedulerTableAdapter.Fill(this.uristDataSet1.sheduler, DateTime.Now);
           
            }

        private void monthCalendar1_DateChanged(object sender, DateRangeEventArgs e)
        {
            DateTime ww= e.Start;
            this.shedulerTableAdapter.Fill(this.uristDataSet1.sheduler, ww);
        } 

        }

       
    }
