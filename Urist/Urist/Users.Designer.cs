namespace Urist
{
    partial class Users
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Users));
            this.gridControl1 = new DevExpress.XtraGrid.GridControl();
            this.usersprogrammBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.uristDataSet1 = new Urist.uristDataSet1();
            this.gridView1 = new DevExpress.XtraGrid.Views.Grid.GridView();
            this.coluser_id = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colid_slujbi = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemLookUpEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit();
            this.sprSlujbaBindingSource = new System.Windows.Forms.BindingSource(this.components);
            this.colfio = new DevExpress.XtraGrid.Columns.GridColumn();
            this.collogin = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colpassword = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colid_doljnost = new DevExpress.XtraGrid.Columns.GridColumn();
            this.coltel = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colemail = new DevExpress.XtraGrid.Columns.GridColumn();
            this.coldata_prin = new DevExpress.XtraGrid.Columns.GridColumn();
            this.coldata_uvoln = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colprava_mod_spravochnika = new DevExpress.XtraGrid.Columns.GridColumn();
            this.colimage = new DevExpress.XtraGrid.Columns.GridColumn();
            this.repositoryItemImageEdit1 = new DevExpress.XtraEditors.Repository.RepositoryItemImageEdit();
            this.users_programmTableAdapter = new Urist.uristDataSet1TableAdapters.users_programmTableAdapter();
            this.sprSlujbaTableAdapter = new Urist.uristDataSet1TableAdapters.sprSlujbaTableAdapter();
            this.ofdInput = new System.Windows.Forms.OpenFileDialog();
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.usersprogrammBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uristDataSet1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.sprSlujbaBindingSource)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemImageEdit1)).BeginInit();
            this.SuspendLayout();
            // 
            // gridControl1
            // 
            this.gridControl1.Cursor = System.Windows.Forms.Cursors.Default;
            this.gridControl1.DataSource = this.usersprogrammBindingSource;
            this.gridControl1.Location = new System.Drawing.Point(2, 1);
            this.gridControl1.MainView = this.gridView1;
            this.gridControl1.Name = "gridControl1";
            this.gridControl1.RepositoryItems.AddRange(new DevExpress.XtraEditors.Repository.RepositoryItem[] {
            this.repositoryItemLookUpEdit1,
            this.repositoryItemImageEdit1});
            this.gridControl1.Size = new System.Drawing.Size(586, 406);
            this.gridControl1.TabIndex = 0;
            this.gridControl1.ViewCollection.AddRange(new DevExpress.XtraGrid.Views.Base.BaseView[] {
            this.gridView1});
            // 
            // usersprogrammBindingSource
            // 
            this.usersprogrammBindingSource.DataMember = "users_programm";
            this.usersprogrammBindingSource.DataSource = this.uristDataSet1;
            // 
            // uristDataSet1
            // 
            this.uristDataSet1.DataSetName = "uristDataSet1";
            this.uristDataSet1.SchemaSerializationMode = System.Data.SchemaSerializationMode.IncludeSchema;
            // 
            // gridView1
            // 
            this.gridView1.Columns.AddRange(new DevExpress.XtraGrid.Columns.GridColumn[] {
            this.coluser_id,
            this.colid_slujbi,
            this.colfio,
            this.collogin,
            this.colpassword,
            this.colid_doljnost,
            this.coltel,
            this.colemail,
            this.coldata_prin,
            this.coldata_uvoln,
            this.colprava_mod_spravochnika,
            this.colimage});
            this.gridView1.GridControl = this.gridControl1;
            this.gridView1.Name = "gridView1";
            this.gridView1.OptionsBehavior.EditingMode = DevExpress.XtraGrid.Views.Grid.GridEditingMode.EditForm;
            this.gridView1.OptionsNavigation.AutoFocusNewRow = true;
            this.gridView1.OptionsView.NewItemRowPosition = DevExpress.XtraGrid.Views.Grid.NewItemRowPosition.Top;
            this.gridView1.InitNewRow += new DevExpress.XtraGrid.Views.Grid.InitNewRowEventHandler(this.gridView1_InitNewRow);
            this.gridView1.RowUpdated += new DevExpress.XtraGrid.Views.Base.RowObjectEventHandler(this.gridView1_RowUpdated);
            // 
            // coluser_id
            // 
            this.coluser_id.FieldName = "user_id";
            this.coluser_id.Name = "coluser_id";
            // 
            // colid_slujbi
            // 
            this.colid_slujbi.Caption = "Подразделение";
            this.colid_slujbi.ColumnEdit = this.repositoryItemLookUpEdit1;
            this.colid_slujbi.FieldName = "id_slujbi";
            this.colid_slujbi.Name = "colid_slujbi";
            this.colid_slujbi.Visible = true;
            this.colid_slujbi.VisibleIndex = 0;
            // 
            // repositoryItemLookUpEdit1
            // 
            this.repositoryItemLookUpEdit1.AutoHeight = false;
            this.repositoryItemLookUpEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemLookUpEdit1.Columns.AddRange(new DevExpress.XtraEditors.Controls.LookUpColumnInfo[] {
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("nazv_slujbi", "Подразделение"),
            new DevExpress.XtraEditors.Controls.LookUpColumnInfo("id_slujbi", "id", 20, DevExpress.Utils.FormatType.None, "", false, DevExpress.Utils.HorzAlignment.Default)});
            this.repositoryItemLookUpEdit1.DataSource = this.sprSlujbaBindingSource;
            this.repositoryItemLookUpEdit1.DisplayMember = "nazv_slujbi";
            this.repositoryItemLookUpEdit1.Name = "repositoryItemLookUpEdit1";
            this.repositoryItemLookUpEdit1.NullText = "";
            this.repositoryItemLookUpEdit1.ValueMember = "id_slujbi";
            this.repositoryItemLookUpEdit1.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.repositoryItemLookUpEdit1_Closed);
            // 
            // sprSlujbaBindingSource
            // 
            this.sprSlujbaBindingSource.DataMember = "sprSlujba";
            this.sprSlujbaBindingSource.DataSource = this.uristDataSet1;
            // 
            // colfio
            // 
            this.colfio.Caption = "ФИО";
            this.colfio.FieldName = "fio";
            this.colfio.Name = "colfio";
            this.colfio.Visible = true;
            this.colfio.VisibleIndex = 1;
            // 
            // collogin
            // 
            this.collogin.Caption = "Логин";
            this.collogin.FieldName = "login";
            this.collogin.Name = "collogin";
            this.collogin.Visible = true;
            this.collogin.VisibleIndex = 2;
            // 
            // colpassword
            // 
            this.colpassword.Caption = "Пароль";
            this.colpassword.FieldName = "password";
            this.colpassword.Name = "colpassword";
            this.colpassword.Visible = true;
            this.colpassword.VisibleIndex = 3;
            // 
            // colid_doljnost
            // 
            this.colid_doljnost.Caption = "Должность";
            this.colid_doljnost.FieldName = "id_doljnost";
            this.colid_doljnost.Name = "colid_doljnost";
            this.colid_doljnost.Visible = true;
            this.colid_doljnost.VisibleIndex = 4;
            // 
            // coltel
            // 
            this.coltel.Caption = "Тел";
            this.coltel.FieldName = "tel";
            this.coltel.Name = "coltel";
            this.coltel.Visible = true;
            this.coltel.VisibleIndex = 5;
            // 
            // colemail
            // 
            this.colemail.Caption = "эмайл";
            this.colemail.FieldName = "email";
            this.colemail.Name = "colemail";
            this.colemail.Visible = true;
            this.colemail.VisibleIndex = 6;
            // 
            // coldata_prin
            // 
            this.coldata_prin.Caption = "Дата прин";
            this.coldata_prin.FieldName = "data_prin";
            this.coldata_prin.Name = "coldata_prin";
            this.coldata_prin.Visible = true;
            this.coldata_prin.VisibleIndex = 7;
            // 
            // coldata_uvoln
            // 
            this.coldata_uvoln.Caption = "Дата уволн";
            this.coldata_uvoln.FieldName = "data_uvoln";
            this.coldata_uvoln.Name = "coldata_uvoln";
            this.coldata_uvoln.Visible = true;
            this.coldata_uvoln.VisibleIndex = 8;
            // 
            // colprava_mod_spravochnika
            // 
            this.colprava_mod_spravochnika.Caption = "Примечание";
            this.colprava_mod_spravochnika.FieldName = "prava_mod_spravochnika";
            this.colprava_mod_spravochnika.Name = "colprava_mod_spravochnika";
            this.colprava_mod_spravochnika.Visible = true;
            this.colprava_mod_spravochnika.VisibleIndex = 9;
            // 
            // colimage
            // 
            this.colimage.Caption = "Фото";
            this.colimage.ColumnEdit = this.repositoryItemImageEdit1;
            this.colimage.FieldName = "image";
            this.colimage.Name = "colimage";
            this.colimage.Visible = true;
            this.colimage.VisibleIndex = 10;
            // 
            // repositoryItemImageEdit1
            // 
            this.repositoryItemImageEdit1.AutoHeight = false;
            this.repositoryItemImageEdit1.Buttons.AddRange(new DevExpress.XtraEditors.Controls.EditorButton[] {
            new DevExpress.XtraEditors.Controls.EditorButton(DevExpress.XtraEditors.Controls.ButtonPredefines.Combo)});
            this.repositoryItemImageEdit1.Name = "repositoryItemImageEdit1";
            this.repositoryItemImageEdit1.Closed += new DevExpress.XtraEditors.Controls.ClosedEventHandler(this.repositoryItemImageEdit1_Closed);
            this.repositoryItemImageEdit1.Click += new System.EventHandler(this.repositoryItemImageEdit1_Click);
            // 
            // users_programmTableAdapter
            // 
            this.users_programmTableAdapter.ClearBeforeFill = true;
            // 
            // sprSlujbaTableAdapter
            // 
            this.sprSlujbaTableAdapter.ClearBeforeFill = true;
            // 
            // ofdInput
            // 
            this.ofdInput.FileOk += new System.ComponentModel.CancelEventHandler(this.ofdInput_FileOk);
            // 
            // Users
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(592, 437);
            this.Controls.Add(this.gridControl1);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.Name = "Users";
            this.Text = "Пользователи";
            this.Load += new System.EventHandler(this.Users_Load);
            ((System.ComponentModel.ISupportInitialize)(this.gridControl1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.usersprogrammBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uristDataSet1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.gridView1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemLookUpEdit1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.sprSlujbaBindingSource)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.repositoryItemImageEdit1)).EndInit();
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraGrid.GridControl gridControl1;
        private DevExpress.XtraGrid.Views.Grid.GridView gridView1;
        private uristDataSet1 uristDataSet1;
        private System.Windows.Forms.BindingSource usersprogrammBindingSource;
        private uristDataSet1TableAdapters.users_programmTableAdapter users_programmTableAdapter;
        private DevExpress.XtraGrid.Columns.GridColumn coluser_id;
        private DevExpress.XtraGrid.Columns.GridColumn colid_slujbi;
        private DevExpress.XtraEditors.Repository.RepositoryItemLookUpEdit repositoryItemLookUpEdit1;
        private DevExpress.XtraGrid.Columns.GridColumn colfio;
        private DevExpress.XtraGrid.Columns.GridColumn collogin;
        private DevExpress.XtraGrid.Columns.GridColumn colpassword;
        private DevExpress.XtraGrid.Columns.GridColumn colid_doljnost;
        private DevExpress.XtraGrid.Columns.GridColumn coltel;
        private DevExpress.XtraGrid.Columns.GridColumn colemail;
        private DevExpress.XtraGrid.Columns.GridColumn coldata_prin;
        private DevExpress.XtraGrid.Columns.GridColumn coldata_uvoln;
        private DevExpress.XtraGrid.Columns.GridColumn colprava_mod_spravochnika;
        private DevExpress.XtraGrid.Columns.GridColumn colimage;
        private System.Windows.Forms.BindingSource sprSlujbaBindingSource;
        private uristDataSet1TableAdapters.sprSlujbaTableAdapter sprSlujbaTableAdapter;
        private DevExpress.XtraEditors.Repository.RepositoryItemImageEdit repositoryItemImageEdit1;
        private System.Windows.Forms.OpenFileDialog ofdInput;
    }
}