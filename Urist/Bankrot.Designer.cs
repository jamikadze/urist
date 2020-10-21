namespace Urist
{
    partial class Bankrot
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
            System.Windows.Forms.Label noteLabel;
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Bankrot));
            this.simpleButton4 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton2 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.memoEdit1 = new DevExpress.XtraEditors.MemoEdit();
            this.deloTableAdapter1 = new Urist.uristDataSet1TableAdapters.deloTableAdapter();
            this.ofdInput = new System.Windows.Forms.OpenFileDialog();
            this.queriesTableAdapter1 = new Urist.uristDataSet1TableAdapters.QueriesTableAdapter();
            this.delo_deb_promTableAdapter1 = new Urist.uristDataSet1TableAdapters.delo_deb_promTableAdapter();
            this.delo_deb_bytTableAdapter1 = new Urist.uristDataSet1TableAdapters.delo_deb_bytTableAdapter();
            noteLabel = new System.Windows.Forms.Label();
            ((System.ComponentModel.ISupportInitialize)(this.memoEdit1.Properties)).BeginInit();
            this.SuspendLayout();
            // 
            // noteLabel
            // 
            noteLabel.AutoSize = true;
            noteLabel.Location = new System.Drawing.Point(12, 9);
            noteLabel.Name = "noteLabel";
            noteLabel.Size = new System.Drawing.Size(192, 13);
            noteLabel.TabIndex = 25;
            noteLabel.Text = "Причина передачи на перемещение:";
            // 
            // simpleButton4
            // 
            this.simpleButton4.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton4.Image")));
            this.simpleButton4.Location = new System.Drawing.Point(344, 454);
            this.simpleButton4.Name = "simpleButton4";
            this.simpleButton4.Size = new System.Drawing.Size(88, 23);
            this.simpleButton4.TabIndex = 30;
            this.simpleButton4.Text = "Закрыть";
            this.simpleButton4.Click += new System.EventHandler(this.simpleButton4_Click);
            // 
            // simpleButton3
            // 
            this.simpleButton3.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton3.Image")));
            this.simpleButton3.Location = new System.Drawing.Point(233, 454);
            this.simpleButton3.Name = "simpleButton3";
            this.simpleButton3.Size = new System.Drawing.Size(97, 23);
            this.simpleButton3.TabIndex = 29;
            this.simpleButton3.Text = "Сохранить";
            this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
            // 
            // simpleButton2
            // 
            this.simpleButton2.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton2.Image")));
            this.simpleButton2.Location = new System.Drawing.Point(126, 454);
            this.simpleButton2.Name = "simpleButton2";
            this.simpleButton2.Size = new System.Drawing.Size(91, 23);
            this.simpleButton2.TabIndex = 28;
            this.simpleButton2.Text = "Прикрепить ";
            this.simpleButton2.Click += new System.EventHandler(this.simpleButton2_Click);
            // 
            // simpleButton1
            // 
            this.simpleButton1.Image = ((System.Drawing.Image)(resources.GetObject("simpleButton1.Image")));
            this.simpleButton1.Location = new System.Drawing.Point(15, 454);
            this.simpleButton1.Name = "simpleButton1";
            this.simpleButton1.Size = new System.Drawing.Size(95, 23);
            this.simpleButton1.TabIndex = 27;
            this.simpleButton1.Text = "Сканировать";
            this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // memoEdit1
            // 
            this.memoEdit1.Anchor = ((System.Windows.Forms.AnchorStyles)((((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Bottom) 
            | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.memoEdit1.Location = new System.Drawing.Point(12, 25);
            this.memoEdit1.Name = "memoEdit1";
            this.memoEdit1.Size = new System.Drawing.Size(569, 423);
            this.memoEdit1.TabIndex = 26;
            // 
            // deloTableAdapter1
            // 
            this.deloTableAdapter1.ClearBeforeFill = true;
            // 
            // delo_deb_promTableAdapter1
            // 
            this.delo_deb_promTableAdapter1.ClearBeforeFill = true;
            // 
            // delo_deb_bytTableAdapter1
            // 
            this.delo_deb_bytTableAdapter1.ClearBeforeFill = true;
            // 
            // Bankrot
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(593, 488);
            this.ControlBox = false;
            this.Controls.Add(this.simpleButton4);
            this.Controls.Add(this.simpleButton3);
            this.Controls.Add(this.simpleButton2);
            this.Controls.Add(this.simpleButton1);
            this.Controls.Add(this.memoEdit1);
            this.Controls.Add(noteLabel);
            this.Name = "Bankrot";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Причина перемещения в папку Банкрот";
            ((System.ComponentModel.ISupportInitialize)(this.memoEdit1.Properties)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private DevExpress.XtraEditors.SimpleButton simpleButton4;
        private DevExpress.XtraEditors.SimpleButton simpleButton3;
        private DevExpress.XtraEditors.SimpleButton simpleButton2;
        private DevExpress.XtraEditors.SimpleButton simpleButton1;
        private DevExpress.XtraEditors.MemoEdit memoEdit1;
        private uristDataSet1TableAdapters.deloTableAdapter deloTableAdapter1;
        private System.Windows.Forms.OpenFileDialog ofdInput;
        private uristDataSet1TableAdapters.QueriesTableAdapter queriesTableAdapter1;
        private uristDataSet1TableAdapters.delo_deb_promTableAdapter delo_deb_promTableAdapter1;
        private uristDataSet1TableAdapters.delo_deb_bytTableAdapter delo_deb_bytTableAdapter1;
    }
}