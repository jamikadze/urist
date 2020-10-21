namespace Urist
{
    partial class Sheduler
    {
        /// <summary>
        /// Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        /// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows Form Designer generated code

        /// <summary>
        /// Required method for Designer support - do not modify
        /// the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            DevExpress.XtraScheduler.TimeRuler timeRuler1 = new DevExpress.XtraScheduler.TimeRuler();
            DevExpress.XtraScheduler.TimeRuler timeRuler2 = new DevExpress.XtraScheduler.TimeRuler();
            this.uristDataSet1 = new Urist.uristDataSet1();
            this.shedulerBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.shedulerTableAdapter = new Urist.uristDataSet1TableAdapters.shedulerTableAdapter();
            this.tableAdapterManager = new Urist.uristDataSet1TableAdapters.TableAdapterManager();
            this.schedulerControl1 = new DevExpress.XtraScheduler.SchedulerControl();
            this.schedulerStorage1 = new DevExpress.XtraScheduler.SchedulerStorage(this.components);
            this.monthCalendar1 = new System.Windows.Forms.MonthCalendar();
            ((System.ComponentModel.ISupportInitialize)(this.uristDataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.shedulerBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage1)).BeginInit();
            this.SuspendLayout();
            // 
            // uristDataSet1
            // 
            this.uristDataSet1.DataSetName = "uristDataSet1";
            this.uristDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // shedulerBindingSource
            // 
            this.shedulerBindingSource.DataMember = "sheduler";
            this.shedulerBindingSource.DataSource = this.uristDataSet1;
            // 
            // shedulerTableAdapter
            // 
            this.shedulerTableAdapter.ClearBeforeFill = true;
            // 
            // tableAdapterManager
            // 
            this.tableAdapterManager.BackupDataSetBeforeUpdate = false;
            this.tableAdapterManager.Connection = null;
            this.tableAdapterManager.deloTableAdapter = null;
            this.tableAdapterManager.doc_prikrep_file_iskiTableAdapter = null;
            this.tableAdapterManager.doc_prikrep_file_jalobyTableAdapter = null;
            this.tableAdapterManager.doc_prikrep_fileTableAdapter = null;
            this.tableAdapterManager.iskiTableAdapter = null;
            this.tableAdapterManager.jalobyTableAdapter = null;
            this.tableAdapterManager.sprPromBytTableAdapter = null;
            this.tableAdapterManager.sprSlujbaTableAdapter = null;
            this.tableAdapterManager.sprStatusSudTableAdapter = null;
            this.tableAdapterManager.sprStatusTableAdapter = null;
            this.tableAdapterManager.sudTableAdapter = null;
            this.tableAdapterManager.UpdateOrder = Urist.uristDataSet1TableAdapters.TableAdapterManager.UpdateOrderOption.InsertUpdateDelete;
            this.tableAdapterManager.users_programmTableAdapter = null;
            // 
            // schedulerControl1
            // 
            this.schedulerControl1.ActiveViewType = DevExpress.XtraScheduler.SchedulerViewType.Month;
            this.schedulerControl1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.schedulerControl1.Location = new System.Drawing.Point(1, 170);
            this.schedulerControl1.Name = "schedulerControl1";
            this.schedulerControl1.Size = new System.Drawing.Size(975, 421);
            this.schedulerControl1.Start = new System.DateTime(2019, 4, 1, 0, 0, 0, 0);
            this.schedulerControl1.Storage = this.schedulerStorage1;
            this.schedulerControl1.TabIndex = 0;
            this.schedulerControl1.Text = "schedulerControl1";
            this.schedulerControl1.Views.DayView.TimeRulers.Add(timeRuler1);
            this.schedulerControl1.Views.WorkWeekView.TimeRulers.Add(timeRuler2);
            // 
            // schedulerStorage1
            // 
            this.schedulerStorage1.Appointments.CustomFieldMappings.Add(new DevExpress.XtraScheduler.AppointmentCustomFieldMapping("Cspot2", "cspot2"));
            this.schedulerStorage1.Appointments.CustomFieldMappings.Add(new DevExpress.XtraScheduler.AppointmentCustomFieldMapping("Instansii", "instansii"));
            this.schedulerStorage1.Appointments.DataSource = this.shedulerBindingSource;
            this.schedulerStorage1.Appointments.Mappings.AllDay = "cspot";
            this.schedulerStorage1.Appointments.Mappings.Description = "fio";
            this.schedulerStorage1.Appointments.Mappings.End = "date_rassmoter";
            this.schedulerStorage1.Appointments.Mappings.Label = "nazvanie";
            this.schedulerStorage1.Appointments.Mappings.Start = "date_statusSud";
            this.schedulerStorage1.Appointments.Mappings.Subject = "note";
            // 
            // monthCalendar1
            // 
            this.monthCalendar1.Location = new System.Drawing.Point(1, 5);
            this.monthCalendar1.Name = "monthCalendar1";
            this.monthCalendar1.TabIndex = 1;
            this.monthCalendar1.DateChanged += new System.Windows.Forms.DateRangeEventHandler(this.monthCalendar1_DateChanged);
            // 
            // Sheduler
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(981, 603);
            this.Controls.Add(this.monthCalendar1);
            this.Controls.Add(this.schedulerControl1);
            this.Name = "Sheduler";
            this.Text = "Расписание судебных заседаний";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Sheduler_Load);
            ((System.ComponentModel.ISupportInitialize)(this.uristDataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.shedulerBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.schedulerStorage1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private uristDataSet1 uristDataSet1;
        private System.Windows.Forms.BindingSource shedulerBindingSource;
        private uristDataSet1TableAdapters.shedulerTableAdapter shedulerTableAdapter;
        private uristDataSet1TableAdapters.TableAdapterManager tableAdapterManager;
        private DevExpress.XtraScheduler.SchedulerControl schedulerControl1;
        private DevExpress.XtraScheduler.SchedulerStorage schedulerStorage1;
        private System.Windows.Forms.MonthCalendar monthCalendar1;
    }
}