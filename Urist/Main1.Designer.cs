namespace Urist
{
    partial class Main1
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
            DevExpress.XtraEditors.TileItemElement tileItemElement1 = new DevExpress.XtraEditors.TileItemElement();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Main1));
            DevExpress.XtraEditors.TileItemElement tileItemElement2 = new DevExpress.XtraEditors.TileItemElement();
            DevExpress.XtraEditors.TileItemElement tileItemElement3 = new DevExpress.XtraEditors.TileItemElement();
            this.tileNavPane1 = new DevExpress.XtraBars.Navigation.TileNavPane();
            this.navButton2 = new DevExpress.XtraBars.Navigation.NavButton();
            this.tileNavCategory3 = new DevExpress.XtraBars.Navigation.TileNavCategory();
            this.tileNavUser = new DevExpress.XtraBars.Navigation.TileNavItem();
            this.tileNavByt = new DevExpress.XtraBars.Navigation.TileNavItem();
            this.tileNavProm = new DevExpress.XtraBars.Navigation.TileNavItem();
            this.SuspendLayout();
            // 
            // tileNavPane1
            // 
            this.tileNavPane1.Appearance.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tileNavPane1.Appearance.Options.UseBackColor = true;
            this.tileNavPane1.ButtonPadding = new System.Windows.Forms.Padding(12);
            this.tileNavPane1.Buttons.Add(this.navButton2);
            this.tileNavPane1.Buttons.Add(this.tileNavCategory3);
            // 
            // tileNavCategory1
            // 
            this.tileNavPane1.DefaultCategory.Name = "tileNavCategory1";
            this.tileNavPane1.DefaultCategory.OptionsDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavPane1.DefaultCategory.OwnerCollection = null;
            // 
            // 
            // 
            this.tileNavPane1.DefaultCategory.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            this.tileNavPane1.DefaultCategory.Tile.ItemSize = DevExpress.XtraBars.Navigation.TileBarItemSize.Default;
            this.tileNavPane1.Dock = System.Windows.Forms.DockStyle.Top;
            this.tileNavPane1.Location = new System.Drawing.Point(0, 0);
            this.tileNavPane1.Name = "tileNavPane1";
            this.tileNavPane1.OptionsPrimaryDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavPane1.OptionsSecondaryDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavPane1.Size = new System.Drawing.Size(828, 40);
            this.tileNavPane1.TabIndex = 0;
            this.tileNavPane1.Text = "tileNavPane1";
            // 
            // navButton2
            // 
            this.navButton2.Caption = "Main Menu";
            this.navButton2.IsMain = true;
            this.navButton2.Name = "navButton2";
            this.navButton2.Visible = false;
            // 
            // tileNavCategory3
            // 
            this.tileNavCategory3.Alignment = DevExpress.XtraBars.Navigation.NavButtonAlignment.Left;
            this.tileNavCategory3.Caption = "Администрирование";
            // 
            // tileNavUser
            // 
            this.tileNavUser.Caption = "Пользователи";
            this.tileNavUser.Name = "tileNavUser";
            this.tileNavUser.OptionsDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavUser.OwnerCollection = this.tileNavCategory3.Items;
            // 
            // 
            // 
            this.tileNavUser.Tile.AppearanceItem.Normal.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tileNavUser.Tile.AppearanceItem.Normal.Options.UseBackColor = true;
            this.tileNavUser.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            tileItemElement1.Image = ((System.Drawing.Image)(resources.GetObject("tileItemElement1.Image")));
            tileItemElement1.Text = "Пользователи";
            this.tileNavUser.Tile.Elements.Add(tileItemElement1);
            this.tileNavUser.Tile.ItemSize = DevExpress.XtraBars.Navigation.TileBarItemSize.Default;
            this.tileNavUser.Tile.Name = "tileBarItem1";
            this.tileNavCategory3.Items.AddRange(new DevExpress.XtraBars.Navigation.TileNavItem[] {
            this.tileNavUser,
            this.tileNavByt,
            this.tileNavProm});
            this.tileNavCategory3.Name = "tileNavCategory3";
            this.tileNavCategory3.OptionsDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavCategory3.OwnerCollection = null;
            // 
            // 
            // 
            this.tileNavCategory3.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            this.tileNavCategory3.Tile.ItemSize = DevExpress.XtraBars.Navigation.TileBarItemSize.Default;
            // 
            // tileNavByt
            // 
            this.tileNavByt.Caption = "Импорт быт";
            this.tileNavByt.Name = "tileNavByt";
            this.tileNavByt.OptionsDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavByt.OwnerCollection = this.tileNavCategory3.Items;
            // 
            // 
            // 
            this.tileNavByt.Tile.AppearanceItem.Normal.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tileNavByt.Tile.AppearanceItem.Normal.Options.UseBackColor = true;
            this.tileNavByt.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            tileItemElement2.Image = ((System.Drawing.Image)(resources.GetObject("tileItemElement2.Image")));
            tileItemElement2.Text = "Импорт быт";
            this.tileNavByt.Tile.Elements.Add(tileItemElement2);
            this.tileNavByt.Tile.ItemSize = DevExpress.XtraBars.Navigation.TileBarItemSize.Default;
            this.tileNavByt.Tile.Name = "tileBarItem2";
            this.tileNavByt.Tile.ItemClick += new DevExpress.XtraEditors.TileItemClickEventHandler(this.tileNavItem2_Tile_ItemClick);
            // 
            // tileNavProm
            // 
            this.tileNavProm.Caption = "Имопрт пром";
            this.tileNavProm.Name = "tileNavProm";
            this.tileNavProm.OptionsDropDown.BackColor = System.Drawing.Color.Empty;
            this.tileNavProm.OwnerCollection = this.tileNavCategory3.Items;
            // 
            // 
            // 
            this.tileNavProm.Tile.AppearanceItem.Normal.BackColor = System.Drawing.SystemColors.ActiveCaption;
            this.tileNavProm.Tile.AppearanceItem.Normal.Options.UseBackColor = true;
            this.tileNavProm.Tile.DropDownOptions.BeakColor = System.Drawing.Color.Empty;
            tileItemElement3.Image = ((System.Drawing.Image)(resources.GetObject("tileItemElement3.Image")));
            tileItemElement3.Text = "Имопрт пром";
            this.tileNavProm.Tile.Elements.Add(tileItemElement3);
            this.tileNavProm.Tile.ItemSize = DevExpress.XtraBars.Navigation.TileBarItemSize.Default;
            this.tileNavProm.Tile.Name = "tileBarItem3";
            this.tileNavProm.Tile.ItemClick += new DevExpress.XtraEditors.TileItemClickEventHandler(this.tileNavItem3_Tile_ItemClick_1);
            // 
            // Main1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(828, 465);
            this.Controls.Add(this.tileNavPane1);
            this.Name = "Main1";
            this.Text = "Main1";
            this.WindowState = System.Windows.Forms.FormWindowState.Maximized;
            this.Load += new System.EventHandler(this.Main1_Load);
            this.ResumeLayout(false);

        }

        #endregion

        private DevExpress.XtraBars.Navigation.TileNavPane tileNavPane1;
        private DevExpress.XtraBars.Navigation.NavButton navButton2;
        private DevExpress.XtraBars.Navigation.TileNavCategory tileNavCategory3;
        private DevExpress.XtraBars.Navigation.TileNavItem tileNavByt;
        private DevExpress.XtraBars.Navigation.TileNavItem tileNavProm;
        private DevExpress.XtraBars.Navigation.TileNavItem tileNavUser;


    }
}