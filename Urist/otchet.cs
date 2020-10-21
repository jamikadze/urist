using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Linq;
using System.Windows.Forms;
using DevExpress.XtraReports.UI;
using DevExpress.XtraReports.Parameters;
using DevExpress.XtraGrid;
using DevExpress.XtraPrinting;
using System.Reflection;
using System.Data.SqlClient;

namespace Urist
{
    public partial class otchet : DevExpress.XtraEditors.XtraForm
    {
       int count_bish01;
       int count_bish02;
       int count_bish03;
       int count_bish04;
       int count_bish05;
       int count_bish06;
       int count_bish07;
       int count_bish08;
       int count_bish09;
       int count_bish010;
       int count_bish011;
       int count_bish1;
       int count_bish2;
       int count_bish3;
       int count_bish4;
       int count_bish5;
       int count_bish6;
       int count_bish7;
       int count_bish8;
       int count_bish9;
       int count_bish10;
       int count_bish11;
       int count_bisha1;
       int count_bisha2;
       int count_bisha3;
       int count_bisha4;
       int count_bisha5;
       int count_bisha6;
       int count_bisha7;
       int count_bisha8;
       int count_bisha9;
       int count_bisha10;
       int count_bisha11;


       int ccount_bish01;
       int ccount_bish02;
       int ccount_bish03;
       int ccount_bish04;
       int ccount_bish05;
       int ccount_bish06;
       int ccount_bish07;
       int ccount_bish08;
       int ccount_bish09;
       int ccount_bish010;
       int ccount_bish011;
       int ccount_bish1;
       int ccount_bish2;
       int ccount_bish3;
       int ccount_bish4;
       int ccount_bish5;
       int ccount_bish6;
       int ccount_bish7;
       int ccount_bish8;
       int ccount_bish9;
       int ccount_bish10;
       int ccount_bish11;
       int ccount_bisha1;
       int ccount_bisha2;
       int ccount_bisha3;
       int ccount_bisha4;
       int ccount_bisha5;
       int ccount_bisha6;
       int ccount_bisha7;
       int ccount_bisha8;
       int ccount_bisha9;
       int ccount_bisha10;
       int ccount_bisha11;
       
        public otchet()
        {
            InitializeComponent();
        }

        private void otchet_Load(object sender, EventArgs e)
        { 
            comboBoxEdit1.SelectedIndex =0;
            gridControl2.Visible = true;
            gridControl1.Visible = false; 
            this.sprSlujbaTableAdapter.FillByReport(this.uristDataSet1.sprSlujba);
            // TODO: This line of code loads data into the 'uristDataSet1.users_programm' table. You can move, or remove it, as needed.
            this.users_programmTableAdapter.FillByUrist(this.uristDataSet1.users_programm);

            rbByt.Checked = true;



          
        }

        private void comboBoxEdit1_SelectedValueChanged(object sender, EventArgs e)
        {
            if (comboBoxEdit1.SelectedText == "Сводный по подразделениям")
            {
                gridControl1.Visible = true;
                gridControl2.Visible = false;
            }
            else
            {
                gridControl2.Visible = true;
                gridControl1.Visible = false;
            }
        }

        private void navBarItem1_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
            Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
            if (Globals.sdate1 < Globals.podate1)
            {
                var rowHandle = gridView2.FocusedRowHandle;

                for (int i = 0; i <= gridView2.RowCount; i++)
                {
                    var check = gridView2.GetRowCellValue(i, "vrem_check");
                    if (Convert.ToString(check) != "")
                    {
                        if (Convert.ToBoolean(check) == true)
                        {
                            Globals.fio_otchet = (string)(gridView2.GetRowCellValue(i, "fio"));
                            this.users_programmTableAdapter.FillByFio(uristDataSet1.users_programm, Globals.fio_otchet);
                            int rowcount = usersprogrammBindingSource.Count;
                            if (rowcount == 1)
                            {

                                Globals.id_user1 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                Globals.id_user2 = 0;
                                Globals.id_user3 = 0;
                                Globals.id_user4 = 0;
                                Globals.id_user5 = 0;
                                svod_po_uristam report = new svod_po_uristam();
                                svod_po_uristam1 report1 = new svod_po_uristam1();
                                byt_vzyskanie report2 = new byt_vzyskanie();
                                prom_vzyskanie report3 = new prom_vzyskanie();
                                Iski report4 = new Iski();
                                Jaloby report5 = new Jaloby();
                                sogl report6 = new sogl();
                                sogl1 report7 = new sogl1();
                                report.Parameters["parameter3"].Value = Globals.id_user1;
                                report.Parameters["parameter4"].Value = Globals.id_user2;
                                report.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report.Parameters["parameter1"].Value = Globals.sdate1;
                                report.Parameters["parameter2"].Value = Globals.podate1; 
                                report.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report1.Parameters["parameter3"].Value = Globals.id_user1;
                                report1.Parameters["parameter4"].Value = Globals.id_user2;
                                report1.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report1.Parameters["parameter1"].Value = Globals.sdate1;
                                report1.Parameters["parameter2"].Value = Globals.podate1; 
                                report1.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report2.Parameters["parameter3"].Value = Globals.id_user1;
                                report2.Parameters["parameter4"].Value = Globals.id_user2;
                                report2.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report2.Parameters["parameter1"].Value = Globals.sdate1;
                                report2.Parameters["parameter2"].Value = Globals.podate1; 
                                report2.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report3.Parameters["parameter3"].Value = Globals.id_user1;
                                report3.Parameters["parameter4"].Value = Globals.id_user2;
                                report3.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report3.Parameters["parameter1"].Value = Globals.sdate1;
                                report3.Parameters["parameter2"].Value = Globals.podate1; 
                                report3.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report4.Parameters["parameter3"].Value = Globals.id_user1;
                                report4.Parameters["parameter4"].Value = Globals.id_user2;
                                report4.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report4.Parameters["parameter1"].Value = Globals.sdate1;
                                report4.Parameters["parameter2"].Value = Globals.podate1; 
                                report4.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report5.Parameters["parameter3"].Value = Globals.id_user1;
                                report5.Parameters["parameter4"].Value = Globals.id_user2;
                                report5.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report5.Parameters["parameter1"].Value = Globals.sdate1;
                                report5.Parameters["parameter2"].Value = Globals.podate1; 
                                report5.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report6.Parameters["parameter3"].Value = Globals.id_user1;
                                report6.Parameters["parameter4"].Value = Globals.id_user2;
                                report6.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report6.Parameters["parameter1"].Value = Globals.sdate1;
                                report6.Parameters["parameter2"].Value = Globals.podate1; 
                                report6.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report7.Parameters["parameter3"].Value = Globals.id_user1;
                                report7.Parameters["parameter4"].Value = Globals.id_user2;
                                report7.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report7.Parameters["parameter1"].Value = Globals.sdate1;
                                report7.Parameters["parameter2"].Value = Globals.podate1; 
                                report7.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report.RequestParameters = false;
                                report1.RequestParameters = false;
                                report2.RequestParameters = false;
                                report3.RequestParameters = false;
                                report4.RequestParameters = false;
                                report5.RequestParameters = false;
                                report6.RequestParameters = false;
                                report7.RequestParameters = false;
                                report.CreateDocument(false);
                                report1.CreateDocument(false);
                                report2.CreateDocument(false);
                                report3.CreateDocument(false);
                                report4.CreateDocument(false);
                                report5.CreateDocument(false);
                                report6.CreateDocument(false);
                                report7.CreateDocument(false);
                                report.Pages.AddRange(report1.Pages);
                                report.Pages.AddRange(report2.Pages);
                                report.Pages.AddRange(report3.Pages);
                                report.Pages.AddRange(report4.Pages);
                                report.Pages.AddRange(report5.Pages);
                                report.Pages.AddRange(report6.Pages);
                                report.Pages.AddRange(report7.Pages);
                                report.PrintingSystem.ContinuousPageNumbering = true;
                                ReportPrintTool printTool = new ReportPrintTool(report);
                                printTool.AutoShowParametersPanel = false;
                                printTool.ShowPreviewDialog();
                                this.users_programmTableAdapter.FillByUrist(this.uristDataSet1.users_programm);
                            }
                            if (rowcount == 2)
                            {

                                Globals.id_user1 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveLast();
                                Globals.id_user2 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                Globals.id_user3 = 0;
                                Globals.id_user4 = 0;
                                Globals.id_user5 = 0;
                                Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
                                Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
                                svod_po_uristam report = new svod_po_uristam();
                                svod_po_uristam1 report1 = new svod_po_uristam1();
                                byt_vzyskanie report2 = new byt_vzyskanie();
                                prom_vzyskanie report3 = new prom_vzyskanie();
                                Iski report4 = new Iski();
                                Jaloby report5 = new Jaloby();
                                sogl report6 = new sogl();
                                sogl1 report7 = new sogl1();
                                report.Parameters["parameter3"].Value = Globals.id_user1;
                                report.Parameters["parameter4"].Value = Globals.id_user2;
                                report.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report.Parameters["parameter1"].Value = Globals.sdate1;
                                report.Parameters["parameter2"].Value = Globals.podate1; 
                                report.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report1.Parameters["parameter3"].Value = Globals.id_user1;
                                report1.Parameters["parameter4"].Value = Globals.id_user2;
                                report1.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report1.Parameters["parameter1"].Value = Globals.sdate1;
                                report1.Parameters["parameter2"].Value = Globals.podate1; 
                                report1.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report2.Parameters["parameter3"].Value = Globals.id_user1;
                                report2.Parameters["parameter4"].Value = Globals.id_user2;
                                report2.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report2.Parameters["parameter1"].Value = Globals.sdate1;
                                report2.Parameters["parameter2"].Value = Globals.podate1; 
                                report2.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report3.Parameters["parameter3"].Value = Globals.id_user1;
                                report3.Parameters["parameter4"].Value = Globals.id_user2;
                                report3.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report3.Parameters["parameter1"].Value = Globals.sdate1;
                                report3.Parameters["parameter2"].Value = Globals.podate1; 
                                report3.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report4.Parameters["parameter3"].Value = Globals.id_user1;
                                report4.Parameters["parameter4"].Value = Globals.id_user2;
                                report4.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report4.Parameters["parameter1"].Value = Globals.sdate1;
                                report4.Parameters["parameter2"].Value = Globals.podate1; 
                                report4.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report5.Parameters["parameter3"].Value = Globals.id_user1;
                                report5.Parameters["parameter4"].Value = Globals.id_user2;
                                report5.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report5.Parameters["parameter1"].Value = Globals.sdate1;
                                report5.Parameters["parameter2"].Value = Globals.podate1; 
                                report5.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report6.Parameters["parameter3"].Value = Globals.id_user1;
                                report6.Parameters["parameter4"].Value = Globals.id_user2;
                                report6.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report6.Parameters["parameter1"].Value = Globals.sdate1;
                                report6.Parameters["parameter2"].Value = Globals.podate1; 
                                report6.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report7.Parameters["parameter3"].Value = Globals.id_user1;
                                report7.Parameters["parameter4"].Value = Globals.id_user2;
                                report7.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report7.Parameters["parameter1"].Value = Globals.sdate1;
                                report7.Parameters["parameter2"].Value = Globals.podate1; 
                                report7.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report.RequestParameters = false;
                                report1.RequestParameters = false;
                                report2.RequestParameters = false;
                                report3.RequestParameters = false;
                                report4.RequestParameters = false;
                                report5.RequestParameters = false;
                                report6.RequestParameters = false;
                                report7.RequestParameters = false;
                                report.CreateDocument(false);
                                report1.CreateDocument(false);
                                report2.CreateDocument(false);
                                report3.CreateDocument(false);
                                report4.CreateDocument(false);
                                report5.CreateDocument(false);
                                report6.CreateDocument(false);
                                report7.CreateDocument(false);
                                report.Pages.AddRange(report1.Pages);
                                report.Pages.AddRange(report2.Pages);
                                report.Pages.AddRange(report3.Pages);
                                report.Pages.AddRange(report4.Pages);
                                report.Pages.AddRange(report5.Pages);
                                report.Pages.AddRange(report6.Pages);
                                report.Pages.AddRange(report7.Pages);
                                report.PrintingSystem.ContinuousPageNumbering = true;
                                ReportPrintTool printTool = new ReportPrintTool(report);
                                printTool.AutoShowParametersPanel = false;
                                printTool.ShowPreviewDialog();
                                this.users_programmTableAdapter.FillByUrist(this.uristDataSet1.users_programm);
                            }
                            if (rowcount == 3)
                            {

                                Globals.id_user1 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveNext();
                                Globals.id_user2 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveLast();
                                Globals.id_user3 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                Globals.id_user4 = 0;
                                Globals.id_user5 = 0;
                                Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
                                Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
                                svod_po_uristam report = new svod_po_uristam();
                                svod_po_uristam1 report1 = new svod_po_uristam1();
                                byt_vzyskanie report2 = new byt_vzyskanie();
                                prom_vzyskanie report3 = new prom_vzyskanie();
                                Iski report4 = new Iski();
                                Jaloby report5 = new Jaloby();
                                sogl report6 = new sogl();
                                sogl1 report7 = new sogl1();
                                report.Parameters["parameter3"].Value = Globals.id_user1;
                                report.Parameters["parameter4"].Value = Globals.id_user2;
                                report.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report.Parameters["parameter1"].Value = Globals.sdate1;
                                report.Parameters["parameter2"].Value = Globals.podate1; 
                                report.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report1.Parameters["parameter3"].Value = Globals.id_user1;
                                report1.Parameters["parameter4"].Value = Globals.id_user2;
                                report1.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report1.Parameters["parameter1"].Value = Globals.sdate1;
                                report1.Parameters["parameter2"].Value = Globals.podate1; 
                                report1.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report2.Parameters["parameter3"].Value = Globals.id_user1;
                                report2.Parameters["parameter4"].Value = Globals.id_user2;
                                report2.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report2.Parameters["parameter1"].Value = Globals.sdate1;
                                report2.Parameters["parameter2"].Value = Globals.podate1; 
                                report2.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report3.Parameters["parameter3"].Value = Globals.id_user1;
                                report3.Parameters["parameter4"].Value = Globals.id_user2;
                                report3.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report3.Parameters["parameter1"].Value = Globals.sdate1;
                                report3.Parameters["parameter2"].Value = Globals.podate1; 
                                report3.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report4.Parameters["parameter3"].Value = Globals.id_user1;
                                report4.Parameters["parameter4"].Value = Globals.id_user2;
                                report4.Parameters["parameter5"].Value = Globals.id_user3;
                                 report.Parameters["parameter7"].Value = Globals.id_user4;
                                 report.Parameters["parameter8"].Value = Globals.id_user5;
                                report4.Parameters["parameter1"].Value = Globals.sdate1;
                                report4.Parameters["parameter2"].Value = Globals.podate1; 
                                report4.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report5.Parameters["parameter3"].Value = Globals.id_user1;
                                report5.Parameters["parameter4"].Value = Globals.id_user2;
                                report5.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report5.Parameters["parameter1"].Value = Globals.sdate1;
                                report5.Parameters["parameter2"].Value = Globals.podate1; 
                                report5.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report6.Parameters["parameter3"].Value = Globals.id_user1;
                                report6.Parameters["parameter4"].Value = Globals.id_user2;
                                report6.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report6.Parameters["parameter1"].Value = Globals.sdate1;
                                report6.Parameters["parameter2"].Value = Globals.podate1; 
                                report6.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report7.Parameters["parameter3"].Value = Globals.id_user1;
                                report7.Parameters["parameter4"].Value = Globals.id_user2;
                                report7.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report7.Parameters["parameter1"].Value = Globals.sdate1;
                                report7.Parameters["parameter2"].Value = Globals.podate1; 
                                report7.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report.RequestParameters = false;
                                report1.RequestParameters = false;
                                report2.RequestParameters = false;
                                report3.RequestParameters = false;
                                report4.RequestParameters = false;
                                report5.RequestParameters = false;
                                report6.RequestParameters = false;
                                report7.RequestParameters = false;
                                report.CreateDocument(false);
                                report1.CreateDocument(false);
                                report2.CreateDocument(false);
                                report3.CreateDocument(false);
                                report4.CreateDocument(false);
                                report5.CreateDocument(false);
                                report6.CreateDocument(false);
                                report7.CreateDocument(false);
                                report.Pages.AddRange(report1.Pages);
                                report.Pages.AddRange(report2.Pages);
                                report.Pages.AddRange(report3.Pages);
                                report.Pages.AddRange(report4.Pages);
                                report.Pages.AddRange(report5.Pages);
                                report.Pages.AddRange(report6.Pages);
                                report.Pages.AddRange(report7.Pages);
                                report.PrintingSystem.ContinuousPageNumbering = true;
                                ReportPrintTool printTool = new ReportPrintTool(report);
                                printTool.AutoShowParametersPanel = false;
                                printTool.ShowPreviewDialog();
                                this.users_programmTableAdapter.FillByUrist(this.uristDataSet1.users_programm);
                            }

                            if (rowcount == 5)
                            {

                                Globals.id_user1 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveNext();
                                Globals.id_user2 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveLast();
                                Globals.id_user3 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveLast();
                                Globals.id_user4 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                usersprogrammBindingSource.MoveLast();
                                Globals.id_user5 = (int)((DataRowView)(usersprogrammBindingSource.Current)).Row["user_id"];
                                Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
                                Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
                                svod_po_uristam report = new svod_po_uristam();
                                svod_po_uristam1 report1 = new svod_po_uristam1();
                                byt_vzyskanie report2 = new byt_vzyskanie();
                                prom_vzyskanie report3 = new prom_vzyskanie();
                                Iski report4 = new Iski();
                                Jaloby report5 = new Jaloby();
                                sogl report6 = new sogl();
                                sogl1 report7 = new sogl1();
                                report.Parameters["parameter3"].Value = Globals.id_user1;
                                report.Parameters["parameter4"].Value = Globals.id_user2;
                                report.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report.Parameters["parameter1"].Value = Globals.sdate1;
                                report.Parameters["parameter2"].Value = Globals.podate1;
                                report.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report1.Parameters["parameter3"].Value = Globals.id_user1;
                                report1.Parameters["parameter4"].Value = Globals.id_user2;
                                report1.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report1.Parameters["parameter1"].Value = Globals.sdate1;
                                report1.Parameters["parameter2"].Value = Globals.podate1;
                                report1.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report2.Parameters["parameter3"].Value = Globals.id_user1;
                                report2.Parameters["parameter4"].Value = Globals.id_user2;
                                report2.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report2.Parameters["parameter1"].Value = Globals.sdate1;
                                report2.Parameters["parameter2"].Value = Globals.podate1;
                                report2.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report3.Parameters["parameter3"].Value = Globals.id_user1;
                                report3.Parameters["parameter4"].Value = Globals.id_user2;
                                report3.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report3.Parameters["parameter1"].Value = Globals.sdate1;
                                report3.Parameters["parameter2"].Value = Globals.podate1;
                                report3.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report4.Parameters["parameter3"].Value = Globals.id_user1;
                                report4.Parameters["parameter4"].Value = Globals.id_user2;
                                report4.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report4.Parameters["parameter1"].Value = Globals.sdate1;
                                report4.Parameters["parameter2"].Value = Globals.podate1;
                                report4.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report5.Parameters["parameter3"].Value = Globals.id_user1;
                                report5.Parameters["parameter4"].Value = Globals.id_user2;
                                report5.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report5.Parameters["parameter1"].Value = Globals.sdate1;
                                report5.Parameters["parameter2"].Value = Globals.podate1;
                                report5.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report6.Parameters["parameter3"].Value = Globals.id_user1;
                                report6.Parameters["parameter4"].Value = Globals.id_user2;
                                report6.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report6.Parameters["parameter1"].Value = Globals.sdate1;
                                report6.Parameters["parameter2"].Value = Globals.podate1;
                                report6.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report7.Parameters["parameter3"].Value = Globals.id_user1;
                                report7.Parameters["parameter4"].Value = Globals.id_user2;
                                report7.Parameters["parameter5"].Value = Globals.id_user3;
                                report.Parameters["parameter7"].Value = Globals.id_user4;
                                report.Parameters["parameter8"].Value = Globals.id_user5;
                                report7.Parameters["parameter1"].Value = Globals.sdate1;
                                report7.Parameters["parameter2"].Value = Globals.podate1;
                                report7.Parameters["parameter6"].Value = Globals.fio_otchet;
                                report.RequestParameters = false;
                                report1.RequestParameters = false;
                                report2.RequestParameters = false;
                                report3.RequestParameters = false;
                                report4.RequestParameters = false;
                                report5.RequestParameters = false;
                                report6.RequestParameters = false;
                                report7.RequestParameters = false;
                                report.CreateDocument(false);
                                report1.CreateDocument(false);
                                report2.CreateDocument(false);
                                report3.CreateDocument(false);
                                report4.CreateDocument(false);
                                report5.CreateDocument(false);
                                report6.CreateDocument(false);
                                report7.CreateDocument(false);
                                report.Pages.AddRange(report1.Pages);
                                report.Pages.AddRange(report2.Pages);
                                report.Pages.AddRange(report3.Pages);
                                report.Pages.AddRange(report4.Pages);
                                report.Pages.AddRange(report5.Pages);
                                report.Pages.AddRange(report6.Pages);
                                report.Pages.AddRange(report7.Pages);
                                report.PrintingSystem.ContinuousPageNumbering = true;
                                ReportPrintTool printTool = new ReportPrintTool(report);
                                printTool.AutoShowParametersPanel = false;
                                printTool.ShowPreviewDialog();
                                this.users_programmTableAdapter.FillByUrist(this.uristDataSet1.users_programm);
                            }

                        }
                    }
                    else {// MessageBox.Show("Выберите юриста!!!");
                    }
                   
                }

            }
            else { MessageBox.Show("Выберите правильную дату!!! \n Дата С должна быть меньше чем ПО"); }

          
        }

        private void navBarItem2_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
            Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
            if (Globals.sdate1 < Globals.podate1)
            
            {
                
                svod_pret_res report = new svod_pret_res();
                svod_pret_res1 report1 = new svod_pret_res1();
                svod_vzys_res report2 = new svod_vzys_res();
                svod_vzys_res1 report3 = new svod_vzys_res1();
                report.Parameters["parameter2"].Value = Globals.sdate1;
                report.Parameters["parameter3"].Value = Globals.podate1;
                report1.Parameters["parameter1"].Value = Globals.sdate1;
                report1.Parameters["parameter2"].Value = Globals.podate1;
                report2.Parameters["parameter1"].Value = Globals.sdate1;
                report2.Parameters["parameter2"].Value = Globals.podate1;
                report3.Parameters["parameter1"].Value = Globals.sdate1;
                report3.Parameters["parameter2"].Value = Globals.podate1;
                report.RequestParameters = false;
                report1.RequestParameters = false;
                report2.RequestParameters = false;
                report3.RequestParameters = false;
                report.CreateDocument(false);
                report1.CreateDocument(false);
                report2.CreateDocument(false);
                report3.CreateDocument(false);
                report.Pages.AddRange(report1.Pages);
                report.Pages.AddRange(report2.Pages);
                report.Pages.AddRange(report3.Pages);
                report.PrintingSystem.ContinuousPageNumbering = true;
                ReportPrintTool printTool = new ReportPrintTool(report);
                printTool.AutoShowParametersPanel = false;
                printTool.ShowPreviewDialog();
            }
             //MessageBox.Show("Выберите правильную дату!!! \n Дата С должна быть меньше чем ПО"); }
            
        }

        private void navBarItem3_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {

            Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
            Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
            if (Globals.sdate1 < Globals.podate1)
            {

                svod_res_otpr_byt report = new svod_res_otpr_byt();
                svod_res_otpr_prom report1 = new svod_res_otpr_prom();
              //  svod_vzys_res report2 = new svod_vzys_res();
               // svod_vzys_res1 report3 = new svod_vzys_res1();
                report.Parameters["parameter1"].Value = Globals.sdate1;
                report.Parameters["parameter2"].Value = Globals.podate1;
                report1.Parameters["parameter1"].Value = Globals.sdate1;
                report1.Parameters["parameter2"].Value = Globals.podate1;
               // report2.Parameters["parameter1"].Value = Globals.sdate1;
               // report2.Parameters["parameter2"].Value = Globals.podate1;
              //  report3.Parameters["parameter1"].Value = Globals.sdate1;
              //  report3.Parameters["parameter2"].Value = Globals.podate1;
                report.RequestParameters = false;
                report1.RequestParameters = false;
              //  report2.RequestParameters = false;
              //  report3.RequestParameters = false;
                report.CreateDocument(false);
                report1.CreateDocument(false);
            //    report2.CreateDocument(false);
             //   report3.CreateDocument(false);
                report.Pages.AddRange(report1.Pages);
              //  report.Pages.AddRange(report2.Pages);
              //  report.Pages.AddRange(report3.Pages);
                report.PrintingSystem.ContinuousPageNumbering = true;
                ReportPrintTool printTool = new ReportPrintTool(report);
                printTool.AutoShowParametersPanel = false;
                printTool.ShowPreviewDialog();
            }
            else { MessageBox.Show("Выберите правильную дату!!! \n Дата С должна быть меньше чем ПО"); }
            
        }

        private void navBarItem4_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            Form dd = new Sheduler();
            dd.Show();
        }

        private void navBarItem3_LinkClicked_1(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        {
            Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
            Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
            if (Globals.sdate1 < Globals.podate1)
            {

                svod_res_otpr_byt report = new svod_res_otpr_byt();
                svod_res_otpr_prom report1 = new svod_res_otpr_prom();
                //  svod_vzys_res report2 = new svod_vzys_res();
                // svod_vzys_res1 report3 = new svod_vzys_res1();
                report.Parameters["parameter1"].Value = Globals.sdate1;
                report.Parameters["parameter2"].Value = Globals.podate1;
                report1.Parameters["parameter1"].Value = Globals.sdate1;
                report1.Parameters["parameter2"].Value = Globals.podate1;
                // report2.Parameters["parameter1"].Value = Globals.sdate1;
                // report2.Parameters["parameter2"].Value = Globals.podate1;
                //  report3.Parameters["parameter1"].Value = Globals.sdate1;
                //  report3.Parameters["parameter2"].Value = Globals.podate1;
                report.RequestParameters = false;
                report1.RequestParameters = false;
                //  report2.RequestParameters = false;
                //  report3.RequestParameters = false;
                report.CreateDocument(false);
                report1.CreateDocument(false);
                //    report2.CreateDocument(false);
                //   report3.CreateDocument(false);
                report.Pages.AddRange(report1.Pages);
                //  report.Pages.AddRange(report2.Pages);
                //  report.Pages.AddRange(report3.Pages);
                report.PrintingSystem.ContinuousPageNumbering = true;
                ReportPrintTool printTool = new ReportPrintTool(report);
                printTool.AutoShowParametersPanel = false;
                printTool.ShowPreviewDialog();
            }
            else { MessageBox.Show("Выберите правильную дату!!! \n Дата С должна быть меньше чем ПО"); }
            
        }
        void SetAllCommandTimeouts(object adapter, int timeout)
        {
            var commands = adapter.GetType().InvokeMember(
                    "CommandCollection",
                    BindingFlags.GetProperty | BindingFlags.Instance | BindingFlags.NonPublic,
                    null, adapter, new object[0]);
            var sqlCommand = (SqlCommand[])commands;
            foreach (var cmd in sqlCommand)
            {
                cmd.CommandTimeout = timeout;
            }
        }
        private void navBarItem5_LinkClicked(object sender, DevExpress.XtraNavBar.NavBarLinkEventArgs e)
        { Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
            Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
            if (Globals.sdate1 < Globals.podate1)
            {

                SetAllCommandTimeouts(byt_prosrTableAdapter, 500);
                SetAllCommandTimeouts(prom_prosrTableAdapter, 500);
            this.byt_prosrTableAdapter.Fill(this.uristDataSet1.byt_prosr,Globals.sdate1,Globals.podate1);
            this.prom_prosrTableAdapter.Fill(this.uristDataSet1.prom_prosr, Globals.sdate1, Globals.podate1);
            if (byt_prosrBindingSource.Count == 0) return;
            if (prom_prosrBindingSource.Count == 0) return;
            Microsoft.Office.Interop.Excel.Application app;
            Microsoft.Office.Interop.Excel.Workbook workbook;
           // Microsoft.Office.Interop.Excel.Worksheet worksheet;
           // Microsoft.Office.Interop.Excel.Worksheet worksheet1;
            app = new Microsoft.Office.Interop.Excel.Application();
            workbook = app.Workbooks.Add(Type.Missing);
            var xlSheets = workbook.Sheets as Microsoft.Office.Interop.Excel.Sheets;
            var worksheet = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[1], Type.Missing, Type.Missing, Type.Missing);
          
            worksheet = null;
           
            app.EnableEvents = false;
            app.Visible = false;
            //app.Visible = true;
            app.DisplayAlerts = false;
            worksheet = workbook.ActiveSheet;
            worksheet.Name = "Быт";
           
       worksheet.get_Range("a1", "a1").ColumnWidth = 4.14;
            worksheet.get_Range("b1", "b1").ColumnWidth = 17.29;
            worksheet.get_Range("c1", "c1").ColumnWidth = 8.29;
            worksheet.get_Range("d1", "d1").ColumnWidth = 8.43;
            worksheet.get_Range("e1", "e1").ColumnWidth = 8.43;
            worksheet.get_Range("f1", "f1").ColumnWidth = 8.43;
            worksheet.get_Range("g1", "g1").ColumnWidth = 8.43;
            worksheet.get_Range("h1", "h1").ColumnWidth = 8.43;
            worksheet.get_Range("i1", "i1").ColumnWidth = 8.43;
            worksheet.get_Range("j1", "j1").ColumnWidth = 8.43;
            worksheet.get_Range("k1", "k1").ColumnWidth = 9.29;
            worksheet.get_Range("l1", "l1").ColumnWidth = 10.29;
            worksheet.get_Range("m1", "m1").ColumnWidth = 8.43;
            worksheet.get_Range("n1", "n1").ColumnWidth = 18.43;
           // worksheet.get_Range("a3", "n4").Height = 40;
            worksheet.get_Range("b1", "c1").Merge();
            worksheet.get_Range("b1", "c1").Value = "Отчет по бытовым абонентам";
            worksheet.get_Range("b2", "b2").Value = "на "+ Globals.podate1;  
            worksheet.get_Range("a3", "a4").Merge();
            worksheet.get_Range("a3", "a4").Value = "№ ";
            worksheet.get_Range("b3", "b4").Merge();
            worksheet.get_Range("b3", "b4").Value = "Наименование";
            worksheet.get_Range("c3", "c4").Merge();
            worksheet.get_Range("c3", "c4").Value = "Всего от ОРЭ";
            worksheet.get_Range("d3", "h3").Merge();
            worksheet.get_Range("d3", "h3").Value = "Из них:";
            worksheet.get_Range("i3", "k3").Merge();
            worksheet.get_Range("i3", "k3").Value = "По проработке";
            worksheet.get_Range("l3", "m3").Merge();
            worksheet.get_Range("l3", "m3").Value = "Просроченные";
            worksheet.get_Range("n3", "n4").Merge();
            worksheet.get_Range("n3", "n4").Value = "Примечание";
            worksheet.get_Range("d4", "d4").Value = " Прин";
            worksheet.get_Range("e4", "e4").Value = " Не прин";
            worksheet.get_Range("f4", "f4").Value = " Безн";
            worksheet.get_Range("g4", "g4").Value = " Списан";
            worksheet.get_Range("h4", "h4").Value = " Пер.на прор";
            worksheet.get_Range("i4", "i4").Value = " Прор-н";
            worksheet.get_Range("j4", "j4").Value = " В деле";
            worksheet.get_Range("k4", "k4").Value = "Проработка(нач.сл.сб)";
            worksheet.get_Range("l4", "l4").Value = "Не переданные в суд";
            worksheet.get_Range("m4", "m4").Value = "не перед в ПССИ";

           
            
            worksheet.get_Range("a1", "n4").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet.get_Range("a1", "n4").Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;
           
            worksheet.get_Range("a1", "n4").Style.WrapText = true;
            worksheet.get_Range("a5", "b5").Merge();
            worksheet.get_Range("a5", "b5").Value = "  г. Бишкек сл.сб";
            worksheet.get_Range("a5", "b5").Font.Bold = true;
            worksheet.get_Range("a10", "b10").Merge();
            worksheet.get_Range("a10", "b10").Value = "Итого";
            worksheet.get_Range("a10", "m10").Font.Bold = true;
            worksheet.get_Range("a11", "b11").Merge();
            worksheet.get_Range("a11", "b11").Value = " Чуйская обл(РЭСы)";
            worksheet.get_Range("a11", "m11").Font.Bold = true;
            worksheet.get_Range("a22", "b22").Merge();
            worksheet.get_Range("a22", "b22").Value = "Итого";
            worksheet.get_Range("a22", "m22").Font.Bold = true;
            worksheet.get_Range("a23", "b23").Merge();
            worksheet.get_Range("a23", "b23").Value = "Таласский ф-л(РЭСы)";
            worksheet.get_Range("a23", "m23").Font.Bold = true;
            worksheet.get_Range("a29", "b29").Merge();
            worksheet.get_Range("a29", "b29").Value = "Итого";
            worksheet.get_Range("a29", "m29").Font.Bold = true;
            worksheet.get_Range("a30", "b30").Merge();
            worksheet.get_Range("a30", "b30").Value = "Итого по ОАО СЭ ";
            worksheet.get_Range("a30", "m30").Font.Bold = true;
            worksheet.get_Range("d4", "h4").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet.get_Range("d6", "h10").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet.get_Range("d12", "h22").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet.get_Range("d24", "h30").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet.get_Range("a1", "n4").Font.Bold = true;
            worksheet.get_Range("a1", "n30").Font.Size=9;
            for (int curRow = 0; curRow < byt_prosrBindingSource.Count; ++curRow)
            {

                DataRowView dtv = (DataRowView)byt_prosrBindingSource[curRow];
                if (dtv.Row["itog"].ToString().Trim() == "г. Бишкек сл.сб")
                {
 
                worksheet.get_Range("a" + (curRow + 6), "m" + (curRow + 6)).NumberFormat = "@";
                worksheet.get_Range("a" + (curRow + 6), "m" + (curRow + 6)).Font.Size = 9;
                worksheet.get_Range("a" + (curRow + 6), "a" + (curRow + 6)).Value = Convert.ToInt32(dtv.Row["number"]);
                worksheet.get_Range("b" + (curRow + 6), "b" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                worksheet.get_Range("b" + (curRow + 6), "b" + (curRow + 6)).Value = dtv.Row["res"];
                worksheet.get_Range("c" + (curRow + 6), "c" + (curRow + 6)).Font.Bold = true;
                worksheet.get_Range("c" + (curRow + 6), "c" + (curRow + 6)).Value = dtv.Row["kol"];
               // worksheet.get_Range("d" + (curRow + 6), "d" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("d" + (curRow + 6), "d" + (curRow + 6)).Value = dtv.Row["prin"];
               // worksheet.get_Range("e" + (curRow + 6), "e" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("e" + (curRow + 6), "e" + (curRow + 6)).Value = dtv.Row["ne_prin"];
             //  worksheet.get_Range("f" + (curRow + 6), "f" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("f" + (curRow + 6), "f" + (curRow + 6)).Value = dtv.Row["bezn"];
              //   worksheet.get_Range("g" + (curRow + 6), "g" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("g" + (curRow + 6), "g" + (curRow + 6)).Value = dtv.Row["spisan"];
             //   worksheet.get_Range("h" + (curRow + 6), "h" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("h" + (curRow + 6), "h" + (curRow + 6)).Value = dtv.Row["pror"];
             //   worksheet.get_Range("i" + (curRow + 6), "i" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("i" + (curRow + 6), "i" + (curRow + 6)).Value = dtv.Row["prorabotan"];
             //   worksheet.get_Range("j" + (curRow + 6), "j" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                int i = Convert.ToInt32(dtv.Row["pror"]) - Convert.ToInt32(dtv.Row["prorabotan"]) - Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                worksheet.get_Range("j" + (curRow + 6), "j" + (curRow + 6)).Value = i;
             //  worksheet.get_Range("k" + (curRow + 6), "k" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("k" + (curRow + 6), "k" + (curRow + 6)).Value = dtv.Row["prosrochka_nach"];
            //     worksheet.get_Range("l" + (curRow + 6), "l" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("l" + (curRow + 6), "l" + (curRow + 6)).Value = dtv.Row["prosrochka_sud"];
             //   worksheet.get_Range("m" + (curRow + 6), "m" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                worksheet.get_Range("m" + (curRow + 6), "m" + (curRow + 6)).Value = dtv.Row["ne_per_pssi"];
                count_bish01 = count_bish01 + Convert.ToInt32(dtv.Row["kol"]);
                count_bish02 = count_bish02 + Convert.ToInt32(dtv.Row["prin"]);
                count_bish03 = count_bish03 + Convert.ToInt32(dtv.Row["ne_prin"]);
                count_bish04 = count_bish04 + Convert.ToInt32(dtv.Row["bezn"]);
                count_bish05 = count_bish05 + Convert.ToInt32(dtv.Row["spisan"]);
                count_bish06 = count_bish06 + Convert.ToInt32(dtv.Row["pror"]);
                count_bish07 = count_bish07 + Convert.ToInt32(dtv.Row["prorabotan"]);
                count_bish08 = count_bish08 +i;
                count_bish09 = count_bish09 + Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                count_bish010 = count_bish010 + Convert.ToInt32(dtv.Row["prosrochka_sud"]);
                count_bish011 = count_bish011 + Convert.ToInt32(dtv.Row["ne_per_pssi"]);
             
                }
                if (dtv.Row["itog"].ToString().Trim() == "Чуйская обл(РЭСы)")
                {
                    
                    worksheet.get_Range("a" + (curRow + 8), "m" + (curRow + 8)).NumberFormat = "@";
                    worksheet.get_Range("a" + (curRow + 8), "m" + (curRow + 8)).Font.Size = 9;
                    worksheet.get_Range("a" + (curRow + 8), "a" + (curRow + 8)).Value = Convert.ToInt32(dtv.Row["number"]);
                    worksheet.get_Range("b" + (curRow + 8), "b" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    worksheet.get_Range("b" + (curRow + 8), "b" + (curRow + 8)).Value = dtv.Row["res"];
                    worksheet.get_Range("c" + (curRow + 8), "c" + (curRow + 8)).Font.Bold = true;
                    worksheet.get_Range("c" + (curRow + 8), "c" + (curRow + 8)).Value = dtv.Row["kol"];
                  //  worksheet.get_Range("с" + (curRow + 8), "с" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("d" + (curRow + 8), "d" + (curRow + 8)).Value = dtv.Row["prin"];
                   // worksheet.get_Range("d" + (curRow + 8), "d" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("e" + (curRow + 8), "e" + (curRow + 8)).Value = dtv.Row["ne_prin"];
                   // worksheet.get_Range("e" + (curRow + 8), "e" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("f" + (curRow + 8), "f" + (curRow + 8)).Value = dtv.Row["bezn"];
                  //  worksheet.get_Range("f" + (curRow + 8), "f" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("g" + (curRow + 8), "g" + (curRow + 8)).Value = dtv.Row["spisan"];
                  //  worksheet.get_Range("g" + (curRow + 8), "g" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("h" + (curRow + 8), "h" + (curRow + 8)).Value = dtv.Row["pror"];
                  //  worksheet.get_Range("h" + (curRow + 8), "h" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("i" + (curRow + 8), "i" + (curRow + 8)).Value = dtv.Row["prorabotan"];
                  //  worksheet.get_Range("i" + (curRow + 8), "i" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    int i = Convert.ToInt32(dtv.Row["pror"]) - Convert.ToInt32(dtv.Row["prorabotan"]) - Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    worksheet.get_Range("j" + (curRow + 8), "j" + (curRow + 8)).Value = i;
                  //  worksheet.get_Range("j" + (curRow + 8), "j" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("k" + (curRow + 8), "k" + (curRow + 8)).Value = dtv.Row["prosrochka_nach"];
                  //  worksheet.get_Range("k" + (curRow + 8), "k" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("l" + (curRow + 8), "l" + (curRow + 8)).Value = dtv.Row["prosrochka_sud"];
                 //   worksheet.get_Range("l" + (curRow + 8), "l" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("m" + (curRow + 8), "m" + (curRow + 8)).Value = dtv.Row["ne_per_pssi"];
                //    worksheet.get_Range("m" + (curRow + 8), "m" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    count_bish1 = count_bish1 + Convert.ToInt32(dtv.Row["kol"]);
                    count_bish2 = count_bish2 + Convert.ToInt32(dtv.Row["prin"]);
                    count_bish3 = count_bish3 + Convert.ToInt32(dtv.Row["ne_prin"]);
                    count_bish4 = count_bish4 + Convert.ToInt32(dtv.Row["bezn"]);
                    count_bish5 = count_bish5 + Convert.ToInt32(dtv.Row["spisan"]);
                    count_bish6 = count_bish6 + Convert.ToInt32(dtv.Row["pror"]);
                    count_bish7 = count_bish7 + Convert.ToInt32(dtv.Row["prorabotan"]);
                    count_bish8 = count_bish8 + i;
                    count_bish9 = count_bish9 + Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    count_bish10 = count_bish10 + Convert.ToInt32(dtv.Row["prosrochka_sud"]);
                    count_bish11 = count_bish11 + Convert.ToInt32(dtv.Row["ne_per_pssi"]);
                }
                if (dtv.Row["itog"].ToString().Trim() == "Таласский ф-л(РЭСы)")
                {
                   
                    worksheet.get_Range("a" + (curRow + 10), "m" + (curRow + 10)).NumberFormat = "@";
                    worksheet.get_Range("a" + (curRow + 10), "m" + (curRow + 10)).Font.Size = 9;
                    worksheet.get_Range("a" + (curRow + 10), "a" + (curRow + 10)).Value = Convert.ToInt32(dtv.Row["number"]);   
                    worksheet.get_Range("b" + (curRow + 10), "b" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    worksheet.get_Range("b" + (curRow + 10), "b" + (curRow + 10)).Value = dtv.Row["res"];
                    worksheet.get_Range("c" + (curRow + 10), "c" + (curRow + 10)).Font.Bold = true;
                    worksheet.get_Range("c" + (curRow + 10), "c" + (curRow + 10)).Value = dtv.Row["kol"];
                  //  worksheet.get_Range("с" + (curRow + 10), "с" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("d" + (curRow +10), "d" + (curRow + 10)).Value = dtv.Row["prin"];
                  //  worksheet.get_Range("d" + (curRow + 10), "d" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("e" + (curRow + 10), "e" + (curRow + 10)).Value = dtv.Row["ne_prin"];
                //    worksheet.get_Range("e" + (curRow + 10), "e" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("f" + (curRow + 10), "f" + (curRow + 10)).Value = dtv.Row["bezn"];
                //    worksheet.get_Range("f" + (curRow + 10), "f" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                  worksheet.get_Range("g" + (curRow + 10), "g" + (curRow +10)).Value = dtv.Row["spisan"];
                 //   worksheet.get_Range("g" + (curRow + 10), "g" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("h" + (curRow + 10), "h" + (curRow + 10)).Value = dtv.Row["pror"];
                //    worksheet.get_Range("h" + (curRow + 10), "h" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("i" + (curRow + 10), "i" + (curRow + 10)).Value = dtv.Row["prorabotan"];
               //     worksheet.get_Range("i" + (curRow + 10), "i" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    int i = Convert.ToInt32(dtv.Row["pror"]) - Convert.ToInt32(dtv.Row["prorabotan"]) - Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    worksheet.get_Range("j" + (curRow + 10), "j" + (curRow + 10)).Value = i;
             //       worksheet.get_Range("j" + (curRow + 10), "j" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("k" + (curRow + 10), "k" + (curRow + 10)).Value = dtv.Row["prosrochka_nach"];
            //        worksheet.get_Range("k" + (curRow + 10), "k" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("l" + (curRow + 10), "l" + (curRow + 10)).Value = dtv.Row["prosrochka_sud"];
           //         worksheet.get_Range("l" + (curRow + 10), "l" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet.get_Range("m" + (curRow + 10), "m" + (curRow + 10)).Value = dtv.Row["ne_per_pssi"];
          //          worksheet.get_Range("m" + (curRow + 10), "m" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    count_bisha1 = count_bisha1 + Convert.ToInt32(dtv.Row["kol"]);
                    count_bisha2 = count_bisha2 + Convert.ToInt32(dtv.Row["prin"]);
                    count_bisha3 = count_bisha3 + Convert.ToInt32(dtv.Row["ne_prin"]);
                    count_bisha4 = count_bisha4 + Convert.ToInt32(dtv.Row["bezn"]);
                    count_bisha5 = count_bisha5 + Convert.ToInt32(dtv.Row["spisan"]);
                    count_bisha6 = count_bisha6 + Convert.ToInt32(dtv.Row["pror"]);
                    count_bisha7 = count_bisha7 + Convert.ToInt32(dtv.Row["prorabotan"]);
                    count_bisha8 = count_bisha8 + i;
                    count_bisha9 = count_bisha9 + Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    count_bisha10 = count_bisha10 + Convert.ToInt32(dtv.Row["prosrochka_sud"]);
                    count_bisha11 = count_bisha11 + Convert.ToInt32(dtv.Row["ne_per_pssi"]);
                }
                
            }
//itog bish
            worksheet.get_Range("c10", "c10").Value = count_bish01;
            worksheet.get_Range("d10", "d10").Value = count_bish02;
            worksheet.get_Range("e10", "e10").Value = count_bish03;
            worksheet.get_Range("f10", "f10").Value = count_bish04;
            worksheet.get_Range("g10", "g10").Value = count_bish05;
            worksheet.get_Range("h10", "h10").Value = count_bish06;
            worksheet.get_Range("i10", "i10").Value = count_bish07;
            worksheet.get_Range("j10", "j10").Value = count_bish08;
            worksheet.get_Range("k10", "k10").Value = count_bish09;
            worksheet.get_Range("l10", "l10").Value = count_bish010;
            worksheet.get_Range("m10", "m10").Value = count_bish011;
            //itog chui
            worksheet.get_Range("c22", "c22").Value = count_bish1;
            worksheet.get_Range("d22", "d22").Value = count_bish2;
            worksheet.get_Range("e22", "e22").Value = count_bish3;
            worksheet.get_Range("f22", "f22").Value = count_bish4;
            worksheet.get_Range("g22", "g22").Value = count_bish5;
            worksheet.get_Range("h22", "h22").Value = count_bish6;
            worksheet.get_Range("i22", "i22").Value = count_bish7;
            worksheet.get_Range("j22", "j22").Value = count_bish8;
            worksheet.get_Range("k22", "k22").Value = count_bish9;
            worksheet.get_Range("l22", "l22").Value = count_bish10;
            worksheet.get_Range("m22", "m22").Value = count_bish11;
            //itog talas
            worksheet.get_Range("c29", "c29").Value = count_bisha1;
            worksheet.get_Range("d29", "d29").Value = count_bisha2;
            worksheet.get_Range("e29", "e29").Value = count_bisha3;
            worksheet.get_Range("f29", "f29").Value = count_bisha4;
            worksheet.get_Range("g29", "g29").Value = count_bisha5;
            worksheet.get_Range("h29", "h29").Value = count_bisha6;
            worksheet.get_Range("i29", "i29").Value = count_bisha7;
            worksheet.get_Range("j29", "j29").Value = count_bisha8;
            worksheet.get_Range("k29", "k29").Value = count_bisha9;
            worksheet.get_Range("l29", "l29").Value = count_bisha10;
            worksheet.get_Range("m29", "m29").Value = count_bisha11;
            //itog sever
            worksheet.get_Range("c30", "c30").Value = count_bisha1 + count_bish01 + count_bish1;
            worksheet.get_Range("d30", "d30").Value = count_bisha2 + count_bish02 + count_bish2;
            worksheet.get_Range("e30", "e30").Value = count_bisha3 + count_bish03 + count_bish3;
            worksheet.get_Range("f30", "f30").Value = count_bisha4+ count_bish04 + count_bish4;
            worksheet.get_Range("g30", "g30").Value = count_bisha5 + count_bish05 + count_bish5;
            worksheet.get_Range("h30", "h30").Value = count_bisha6 + count_bish06 + count_bish6;
            worksheet.get_Range("i30", "i30").Value = count_bisha7 + count_bish07 + count_bish7;
            worksheet.get_Range("j30", "j30").Value = count_bisha8 + count_bish08 + count_bish8;
            worksheet.get_Range("k30", "k30").Value = count_bisha9 + count_bish09 + count_bish9;
            worksheet.get_Range("l30", "l30").Value = count_bisha10 + count_bish010 + count_bish10;
            worksheet.get_Range("m30", "m30").Value = count_bisha11 + count_bish011 + count_bish11;

          
            worksheet.get_Range("a3", "n30").Borders.LineStyle = DataGridLineStyle.Solid;
            worksheet.Columns["A:N"].Font.Name = "TimesNewRoman";
        ///////////////////////////
           var worksheet1 = (Microsoft.Office.Interop.Excel.Worksheet)xlSheets.Add(xlSheets[2], Type.Missing, Type.Missing, Type.Missing);
            worksheet1 = null;
            worksheet1 = workbook.ActiveSheet;
            worksheet1.Name = "Пром";

            worksheet1.get_Range("a1", "a1").ColumnWidth = 4.14;
            worksheet1.get_Range("b1", "b1").ColumnWidth = 17.29;
            worksheet1.get_Range("c1", "c1").ColumnWidth = 8.29;
            worksheet1.get_Range("d1", "d1").ColumnWidth = 8.43;
            worksheet1.get_Range("e1", "e1").ColumnWidth = 8.43;
            worksheet1.get_Range("f1", "f1").ColumnWidth = 8.43;
            worksheet1.get_Range("g1", "g1").ColumnWidth = 8.43;
            worksheet1.get_Range("h1", "h1").ColumnWidth = 8.43;
            worksheet1.get_Range("i1", "i1").ColumnWidth = 8.43;
            worksheet1.get_Range("j1", "j1").ColumnWidth = 8.43;
            worksheet1.get_Range("k1", "k1").ColumnWidth = 9.29;
            worksheet1.get_Range("l1", "l1").ColumnWidth = 10.29;
            worksheet1.get_Range("m1", "m1").ColumnWidth = 8.43;
            worksheet1.get_Range("n1", "n1").ColumnWidth = 18.43;
            // worksheet1.get_Range("a3", "n4").Height = 40;
            worksheet1.get_Range("b1", "c1").Merge();
            worksheet1.get_Range("b1", "c1").Value = "Отчет по пром абонентам";
            worksheet1.get_Range("b2", "b2").Value = "на " + Globals.podate1;
            worksheet1.get_Range("a3", "a4").Merge();
            worksheet1.get_Range("a3", "a4").Value = "№ ";
            worksheet1.get_Range("b3", "b4").Merge();
            worksheet1.get_Range("b3", "b4").Value = "Наименование";
            worksheet1.get_Range("c3", "c4").Merge();
            worksheet1.get_Range("c3", "c4").Value = "Всего от ОРЭ";
            worksheet1.get_Range("d3", "h3").Merge();
            worksheet1.get_Range("d3", "h3").Value = "Из них:";
            worksheet1.get_Range("i3", "k3").Merge();
            worksheet1.get_Range("i3", "k3").Value = "По проработке";
            worksheet1.get_Range("l3", "m3").Merge();
            worksheet1.get_Range("l3", "m3").Value = "Просроченные";
            worksheet1.get_Range("n3", "n4").Merge();
            worksheet1.get_Range("n3", "n4").Value = "Примечание";
            worksheet1.get_Range("d4", "d4").Value = " Прин";
            worksheet1.get_Range("e4", "e4").Value = " Не прин";
            worksheet1.get_Range("f4", "f4").Value = " Безн";
            worksheet1.get_Range("g4", "g4").Value = " Списан";
            worksheet1.get_Range("h4", "h4").Value = " Пер.на прор";
            worksheet1.get_Range("i4", "i4").Value = " Прор-н";
            worksheet1.get_Range("j4", "j4").Value = " В деле";
            worksheet1.get_Range("k4", "k4").Value = "Проработка(нач.сл.сб)";
            worksheet1.get_Range("l4", "l4").Value = "Не переданные в суд";
            worksheet1.get_Range("m4", "m4").Value = "не перед в ПССИ";



            worksheet1.get_Range("a1", "n4").Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
            worksheet1.get_Range("a1", "n4").Style.VerticalAlignment = Microsoft.Office.Interop.Excel.XlVAlign.xlVAlignCenter;

            worksheet1.get_Range("a1", "n4").Style.WrapText = true;
            worksheet1.get_Range("a5", "b5").Merge();
            worksheet1.get_Range("a5", "b5").Value = "  г. Бишкек сл.сб";
            worksheet1.get_Range("a5", "b5").Font.Bold = true;
            worksheet1.get_Range("a10", "b10").Merge();
            worksheet1.get_Range("a10", "b10").Value = "Итого";
            worksheet1.get_Range("a10", "m10").Font.Bold = true;
            worksheet1.get_Range("a11", "b11").Merge();
            worksheet1.get_Range("a11", "b11").Value = " Чуйская обл(РЭСы)";
            worksheet1.get_Range("a11", "m11").Font.Bold = true;
            worksheet1.get_Range("a22", "b22").Merge();
            worksheet1.get_Range("a22", "b22").Value = "Итого";
            worksheet1.get_Range("a22", "m22").Font.Bold = true;
            worksheet1.get_Range("a23", "b23").Merge();
            worksheet1.get_Range("a23", "b23").Value = "Таласский ф-л(РЭСы)";
            worksheet1.get_Range("a23", "m23").Font.Bold = true;
            worksheet1.get_Range("a29", "b29").Merge();
            worksheet1.get_Range("a29", "b29").Value = "Итого";
            worksheet1.get_Range("a29", "m29").Font.Bold = true;
            worksheet1.get_Range("a30", "b30").Merge();
            worksheet1.get_Range("a30", "b30").Value = "Итого по ОАО СЭ ";
            worksheet1.get_Range("a30", "m30").Font.Bold = true;
            worksheet1.get_Range("d4", "h4").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet1.get_Range("d6", "h10").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet1.get_Range("d12", "h22").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet1.get_Range("d24", "h30").Interior.Color = System.Drawing.Color.FromArgb(220, 220, 220);
            worksheet1.get_Range("a1", "n4").Font.Bold = true;
            worksheet1.get_Range("a1", "n30").Font.Size = 9;
            for (int curRow = 0; curRow < prom_prosrBindingSource.Count; ++curRow)
            {

                DataRowView dtv = (DataRowView)prom_prosrBindingSource[curRow];
                if (dtv.Row["itog"].ToString().Trim() == "г. Бишкек сл.сб")
                {

                    worksheet1.get_Range("a" + (curRow + 6), "m" + (curRow + 6)).NumberFormat = "@";
                    worksheet1.get_Range("a" + (curRow + 6), "m" + (curRow + 6)).Font.Size = 9;
                    worksheet1.get_Range("a" + (curRow + 6), "a" + (curRow + 6)).Value = Convert.ToInt32(dtv.Row["number"]);
                    worksheet1.get_Range("b" + (curRow + 6), "b" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    worksheet1.get_Range("b" + (curRow + 6), "b" + (curRow + 6)).Value = dtv.Row["res"];
                    worksheet1.get_Range("c" + (curRow + 6), "c" + (curRow + 6)).Font.Bold = true;
                    worksheet1.get_Range("c" + (curRow + 6), "c" + (curRow + 6)).Value = dtv.Row["kol"];
                    // worksheet1.get_Range("d" + (curRow + 6), "d" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("d" + (curRow + 6), "d" + (curRow + 6)).Value = dtv.Row["prin"];
                    // worksheet1.get_Range("e" + (curRow + 6), "e" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("e" + (curRow + 6), "e" + (curRow + 6)).Value = dtv.Row["ne_prin"];
                    //  worksheet1.get_Range("f" + (curRow + 6), "f" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("f" + (curRow + 6), "f" + (curRow + 6)).Value = dtv.Row["bezn"];
                    //   worksheet1.get_Range("g" + (curRow + 6), "g" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("g" + (curRow + 6), "g" + (curRow + 6)).Value = dtv.Row["spisan"];
                    //   worksheet1.get_Range("h" + (curRow + 6), "h" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("h" + (curRow + 6), "h" + (curRow + 6)).Value = dtv.Row["pror"];
                    //   worksheet1.get_Range("i" + (curRow + 6), "i" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("i" + (curRow + 6), "i" + (curRow + 6)).Value = dtv.Row["prorabotan"];
                    //   worksheet1.get_Range("j" + (curRow + 6), "j" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    int i = Convert.ToInt32(dtv.Row["pror"]) - Convert.ToInt32(dtv.Row["prorabotan"]) - Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    worksheet1.get_Range("j" + (curRow + 6), "j" + (curRow + 6)).Value = i;
                    //  worksheet1.get_Range("k" + (curRow + 6), "k" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("k" + (curRow + 6), "k" + (curRow + 6)).Value = dtv.Row["prosrochka_nach"];
                    //     worksheet1.get_Range("l" + (curRow + 6), "l" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("l" + (curRow + 6), "l" + (curRow + 6)).Value = dtv.Row["prosr_sud"];
                    //   worksheet1.get_Range("m" + (curRow + 6), "m" + (curRow + 6)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("m" + (curRow + 6), "m" + (curRow + 6)).Value = dtv.Row["ne_per_pssi"];
                    ccount_bish01 = ccount_bish01 + Convert.ToInt32(dtv.Row["kol"]);
                    ccount_bish02 = ccount_bish02 + Convert.ToInt32(dtv.Row["prin"]);
                    ccount_bish03 = ccount_bish03 + Convert.ToInt32(dtv.Row["ne_prin"]);
                    ccount_bish04 = ccount_bish04 + Convert.ToInt32(dtv.Row["bezn"]);
                    ccount_bish05 = ccount_bish05 + Convert.ToInt32(dtv.Row["spisan"]);
                    ccount_bish06 = ccount_bish06 + Convert.ToInt32(dtv.Row["pror"]);
                    ccount_bish07 = ccount_bish07 + Convert.ToInt32(dtv.Row["prorabotan"]);
                    ccount_bish08 = ccount_bish08 + i;
                    ccount_bish09 = ccount_bish09 + Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    ccount_bish010 = ccount_bish010 + Convert.ToInt32(dtv.Row["prosr_sud"]);
                    ccount_bish011 = ccount_bish011 + Convert.ToInt32(dtv.Row["ne_per_pssi"]);

                }
                if (dtv.Row["itog"].ToString().Trim() == "Чуйская обл(РЭСы)")
                {

                    worksheet1.get_Range("a" + (curRow + 8), "m" + (curRow + 8)).NumberFormat = "@";
                    worksheet1.get_Range("a" + (curRow + 8), "m" + (curRow + 8)).Font.Size = 9;
                    worksheet1.get_Range("a" + (curRow + 8), "a" + (curRow + 8)).Value = Convert.ToInt32(dtv.Row["number"]);
                    worksheet1.get_Range("b" + (curRow + 8), "b" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    worksheet1.get_Range("b" + (curRow + 8), "b" + (curRow + 8)).Value = dtv.Row["res"];
                    worksheet1.get_Range("c" + (curRow + 8), "c" + (curRow + 8)).Font.Bold = true;
                    worksheet1.get_Range("c" + (curRow + 8), "c" + (curRow + 8)).Value = dtv.Row["kol"];
                    //  worksheet1.get_Range("с" + (curRow + 8), "с" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("d" + (curRow + 8), "d" + (curRow + 8)).Value = dtv.Row["prin"];
                    // worksheet1.get_Range("d" + (curRow + 8), "d" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("e" + (curRow + 8), "e" + (curRow + 8)).Value = dtv.Row["ne_prin"];
                    // worksheet1.get_Range("e" + (curRow + 8), "e" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("f" + (curRow + 8), "f" + (curRow + 8)).Value = dtv.Row["bezn"];
                    //  worksheet1.get_Range("f" + (curRow + 8), "f" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("g" + (curRow + 8), "g" + (curRow + 8)).Value = dtv.Row["spisan"];
                    //  worksheet1.get_Range("g" + (curRow + 8), "g" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("h" + (curRow + 8), "h" + (curRow + 8)).Value = dtv.Row["pror"];
                    //  worksheet1.get_Range("h" + (curRow + 8), "h" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("i" + (curRow + 8), "i" + (curRow + 8)).Value = dtv.Row["prorabotan"];
                    //  worksheet1.get_Range("i" + (curRow + 8), "i" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    int i = Convert.ToInt32(dtv.Row["pror"]) - Convert.ToInt32(dtv.Row["prorabotan"]) - Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    worksheet1.get_Range("j" + (curRow + 8), "j" + (curRow + 8)).Value = i;
                    //  worksheet1.get_Range("j" + (curRow + 8), "j" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("k" + (curRow + 8), "k" + (curRow + 8)).Value = dtv.Row["prosrochka_nach"];
                    //  worksheet1.get_Range("k" + (curRow + 8), "k" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("l" + (curRow + 8), "l" + (curRow + 8)).Value = dtv.Row["prosr_sud"];
                    //   worksheet1.get_Range("l" + (curRow + 8), "l" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("m" + (curRow + 8), "m" + (curRow + 8)).Value = dtv.Row["ne_per_pssi"];
                    //    worksheet1.get_Range("m" + (curRow + 8), "m" + (curRow + 8)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ccount_bish1 = ccount_bish1 + Convert.ToInt32(dtv.Row["kol"]);
                    ccount_bish2 = ccount_bish2 + Convert.ToInt32(dtv.Row["prin"]);
                    ccount_bish3 = ccount_bish3 + Convert.ToInt32(dtv.Row["ne_prin"]);
                    ccount_bish4 = ccount_bish4 + Convert.ToInt32(dtv.Row["bezn"]);
                    ccount_bish5 = ccount_bish5 + Convert.ToInt32(dtv.Row["spisan"]);
                    ccount_bish6 = ccount_bish6 + Convert.ToInt32(dtv.Row["pror"]);
                    ccount_bish7 = ccount_bish7 + Convert.ToInt32(dtv.Row["prorabotan"]);
                    ccount_bish8 = ccount_bish8 + i;
                    ccount_bish9 = ccount_bish9 + Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    ccount_bish10 = ccount_bish10 + Convert.ToInt32(dtv.Row["prosr_sud"]);
                    ccount_bish11 = ccount_bish11 + Convert.ToInt32(dtv.Row["ne_per_pssi"]);
                }
                if (dtv.Row["itog"].ToString().Trim() == "Таласский ф-л(РЭСы)")
                {

                    worksheet1.get_Range("a" + (curRow + 10), "m" + (curRow + 10)).NumberFormat = "@";
                    worksheet1.get_Range("a" + (curRow + 10), "m" + (curRow + 10)).Font.Size = 9;
                    worksheet1.get_Range("a" + (curRow + 10), "a" + (curRow + 10)).Value = Convert.ToInt32(dtv.Row["number"]);
                    worksheet1.get_Range("b" + (curRow + 10), "b" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignLeft;
                    worksheet1.get_Range("b" + (curRow + 10), "b" + (curRow + 10)).Value = dtv.Row["res"];
                    worksheet1.get_Range("c" + (curRow + 10), "c" + (curRow + 10)).Font.Bold = true;
                    worksheet1.get_Range("c" + (curRow + 10), "c" + (curRow + 10)).Value = dtv.Row["kol"];
                    //  worksheet1.get_Range("с" + (curRow + 10), "с" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("d" + (curRow + 10), "d" + (curRow + 10)).Value = dtv.Row["prin"];
                    //  worksheet1.get_Range("d" + (curRow + 10), "d" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("e" + (curRow + 10), "e" + (curRow + 10)).Value = dtv.Row["ne_prin"];
                    //    worksheet1.get_Range("e" + (curRow + 10), "e" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("f" + (curRow + 10), "f" + (curRow + 10)).Value = dtv.Row["bezn"];
                    //    worksheet1.get_Range("f" + (curRow + 10), "f" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("g" + (curRow + 10), "g" + (curRow + 10)).Value = dtv.Row["spisan"];
                    //   worksheet1.get_Range("g" + (curRow + 10), "g" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("h" + (curRow + 10), "h" + (curRow + 10)).Value = dtv.Row["pror"];
                    //    worksheet1.get_Range("h" + (curRow + 10), "h" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("i" + (curRow + 10), "i" + (curRow + 10)).Value = dtv.Row["prorabotan"];
                    //     worksheet1.get_Range("i" + (curRow + 10), "i" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    int i = Convert.ToInt32(dtv.Row["pror"]) - Convert.ToInt32(dtv.Row["prorabotan"]) - Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    worksheet1.get_Range("j" + (curRow + 10), "j" + (curRow + 10)).Value = i;
                    //       worksheet1.get_Range("j" + (curRow + 10), "j" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("k" + (curRow + 10), "k" + (curRow + 10)).Value = dtv.Row["prosrochka_nach"];
                    //        worksheet1.get_Range("k" + (curRow + 10), "k" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("l" + (curRow + 10), "l" + (curRow + 10)).Value = dtv.Row["prosr_sud"];
                    //         worksheet1.get_Range("l" + (curRow + 10), "l" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    worksheet1.get_Range("m" + (curRow + 10), "m" + (curRow + 10)).Value = dtv.Row["ne_per_pssi"];
                    //          worksheet1.get_Range("m" + (curRow + 10), "m" + (curRow + 10)).Style.HorizontalAlignment = Microsoft.Office.Interop.Excel.XlHAlign.xlHAlignCenter;
                    ccount_bisha1 = ccount_bisha1 + Convert.ToInt32(dtv.Row["kol"]);
                    ccount_bisha2 = ccount_bisha2 + Convert.ToInt32(dtv.Row["prin"]);
                    ccount_bisha3 = ccount_bisha3 + Convert.ToInt32(dtv.Row["ne_prin"]);
                    ccount_bisha4 = ccount_bisha4 + Convert.ToInt32(dtv.Row["bezn"]);
                    ccount_bisha5 = ccount_bisha5 + Convert.ToInt32(dtv.Row["spisan"]);
                    ccount_bisha6 = ccount_bisha6 + Convert.ToInt32(dtv.Row["pror"]);
                    ccount_bisha7 = ccount_bisha7 + Convert.ToInt32(dtv.Row["prorabotan"]);
                    ccount_bisha8 = ccount_bisha8 + i;
                    ccount_bisha9 = ccount_bisha9 + Convert.ToInt32(dtv.Row["prosrochka_nach"]);
                    ccount_bisha10 = ccount_bisha10 + Convert.ToInt32(dtv.Row["prosr_sud"]);
                    ccount_bisha11 = ccount_bisha11 + Convert.ToInt32(dtv.Row["ne_per_pssi"]);
                }

            }
            //itog bish
            worksheet1.get_Range("c10", "c10").Value = ccount_bish01;
            worksheet1.get_Range("d10", "d10").Value = ccount_bish02;
            worksheet1.get_Range("e10", "e10").Value = ccount_bish03;
            worksheet1.get_Range("f10", "f10").Value = ccount_bish04;
            worksheet1.get_Range("g10", "g10").Value = ccount_bish05;
            worksheet1.get_Range("h10", "h10").Value = ccount_bish06;
            worksheet1.get_Range("i10", "i10").Value = ccount_bish07;
            worksheet1.get_Range("j10", "j10").Value = ccount_bish08;
            worksheet1.get_Range("k10", "k10").Value = ccount_bish09;
            worksheet1.get_Range("l10", "l10").Value = ccount_bish010;
            worksheet1.get_Range("m10", "m10").Value = ccount_bish011;
            //itog chui
            worksheet1.get_Range("c22", "c22").Value = ccount_bish1;
            worksheet1.get_Range("d22", "d22").Value = ccount_bish2;
            worksheet1.get_Range("e22", "e22").Value = ccount_bish3;
            worksheet1.get_Range("f22", "f22").Value = ccount_bish4;
            worksheet1.get_Range("g22", "g22").Value = ccount_bish5;
            worksheet1.get_Range("h22", "h22").Value = ccount_bish6;
            worksheet1.get_Range("i22", "i22").Value = ccount_bish7;
            worksheet1.get_Range("j22", "j22").Value = ccount_bish8;
            worksheet1.get_Range("k22", "k22").Value = ccount_bish9;
            worksheet1.get_Range("l22", "l22").Value = ccount_bish10;
            worksheet1.get_Range("m22", "m22").Value = ccount_bish11;
            //itog talas
            worksheet1.get_Range("c29", "c29").Value = ccount_bisha1;
            worksheet1.get_Range("d29", "d29").Value = ccount_bisha2;
            worksheet1.get_Range("e29", "e29").Value = ccount_bisha3;
            worksheet1.get_Range("f29", "f29").Value = ccount_bisha4;
            worksheet1.get_Range("g29", "g29").Value = ccount_bisha5;
            worksheet1.get_Range("h29", "h29").Value = ccount_bisha6;
            worksheet1.get_Range("i29", "i29").Value = ccount_bisha7;
            worksheet1.get_Range("j29", "j29").Value = ccount_bisha8;
            worksheet1.get_Range("k29", "k29").Value = ccount_bisha9;
            worksheet1.get_Range("l29", "l29").Value = ccount_bisha10;
            worksheet1.get_Range("m29", "m29").Value = ccount_bisha11;
            //itog sever
            worksheet1.get_Range("c30", "c30").Value = ccount_bisha1 + ccount_bish01 + ccount_bish1;
            worksheet1.get_Range("d30", "d30").Value = ccount_bisha2 + ccount_bish02 + ccount_bish2;
            worksheet1.get_Range("e30", "e30").Value = ccount_bisha3 + ccount_bish03 + ccount_bish3;
            worksheet1.get_Range("f30", "f30").Value = ccount_bisha4 + ccount_bish04 + ccount_bish4;
            worksheet1.get_Range("g30", "g30").Value = ccount_bisha5 + ccount_bish05 + ccount_bish5;
            worksheet1.get_Range("h30", "h30").Value = ccount_bisha6 + ccount_bish06 + ccount_bish6;
            worksheet1.get_Range("i30", "i30").Value = ccount_bisha7 + ccount_bish07 + ccount_bish7;
            worksheet1.get_Range("j30", "j30").Value = ccount_bisha8 + ccount_bish08 + ccount_bish8;
            worksheet1.get_Range("k30", "k30").Value = ccount_bisha9 + ccount_bish09 + ccount_bish9;
            worksheet1.get_Range("l30", "l30").Value = ccount_bisha10 + ccount_bish010 + ccount_bish10;
            worksheet1.get_Range("m30", "m30").Value = ccount_bisha11 + ccount_bish011 + ccount_bish11;


            worksheet1.get_Range("a3", "n30").Borders.LineStyle = DataGridLineStyle.Solid;
            worksheet1.Columns["A:N"].Font.Name = "TimesNewRoman";
                ////////////////////////////////////////////////////////////////////////////
            app.Visible = true;
           
            }
                 else { MessageBox.Show("Выберите правильную дату!!! \n Дата С должна быть меньше чем ПО"); }

        }

        private void navBarControl1_Click(object sender, EventArgs e)
        {

        }

        private void linkLabel1_LinkClicked(object sender, LinkLabelLinkClickedEventArgs e)
        {
            Globals.sdate1 = Convert.ToDateTime(dateEdit1.Text);
            Globals.podate1 = Convert.ToDateTime(dateEdit2.Text);
            if (Globals.sdate1 < Globals.podate1)
            {
                if (rbByt.Checked)
                {
                    svod_prosr_det report = new svod_prosr_det();

                    report.Parameters["parameter1"].Value = Globals.sdate1;
                    report.Parameters["parameter2"].Value = Globals.podate1;
                    report.Parameters["parameter3"].Value = lookUpEdit1.Text.Trim();
                    report.RequestParameters = false;

                    report.PrintingSystem.ContinuousPageNumbering = true;
                    ReportPrintTool printTool = new ReportPrintTool(report);
                    printTool.AutoShowParametersPanel = false;
                    printTool.ShowPreviewDialog();
                }
                else if (rbProm.Checked)
                {
                    svod_prosr_detProm report = new svod_prosr_detProm();

                    report.Parameters["parameter1"].Value = Globals.sdate1;
                    report.Parameters["parameter2"].Value = Globals.podate1;
                    report.Parameters["parameter3"].Value = lookUpEdit1.Text.Trim();
                    report.RequestParameters = false;

                    report.PrintingSystem.ContinuousPageNumbering = true;
                    ReportPrintTool printTool = new ReportPrintTool(report);
                    printTool.AutoShowParametersPanel = false;
                    printTool.ShowPreviewDialog();
                }

                
            }
            //MessageBox.Show("Выберите правильную дату!!! \n Дата С должна быть меньше чем ПО"); }
        }

       
    }
}