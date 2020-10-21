using System;
using System.Text;
using System.Threading;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.InteropServices;
using System.IO;
using TwainLib;
using PdfSharp.Pdf;
using PdfSharp.Drawing;
using Urist;


namespace urist 
{
  public class MainFrame :  DevExpress.XtraEditors.XtraForm, IMessageFilter
  {
      private Panel panel1;
      private Label lblcountscan;
      private TextBox tbcntImage;
    private DevExpress.XtraEditors.DropDownButton dropDownButton1;
    private DevExpress.XtraBars.PopupMenu popupMenu1;
    private DevExpress.XtraBars.BarButtonItem barButtonItem1;
    private DevExpress.XtraBars.BarButtonItem barButtonItem2;
    private DevExpress.XtraBars.BarManager barManager1;
    private DevExpress.XtraBars.BarDockControl barDockControlTop;
    private DevExpress.XtraBars.BarDockControl barDockControlBottom;
    private DevExpress.XtraBars.BarDockControl barDockControlLeft;
    private DevExpress.XtraBars.BarDockControl barDockControlRight;
    private DevExpress.XtraBars.BarButtonItem barButtonItem3;
    private DevExpress.XtraEditors.SimpleButton simpleButton1;
    private DevExpress.XtraEditors.SimpleButton simpleButton3;
    private DevExpress.XtraEditors.SimpleButton simpleButton4;
    private DevExpress.XtraEditors.SimpleButton simpleButton5;
    private Urist.uristDataSet1TableAdapters.QueriesTableAdapter queriesTableAdapter1;
    private DevExpress.XtraEditors.LabelControl labelControl1;
    private DevExpress.XtraEditors.TextEdit textEdit1;
    private IContainer components;

    public MainFrame()
    {
      InitializeComponent();
      tw = new Twain();
      tw.Init(this.Handle);
    }

    protected override void Dispose(bool disposing)
    {
      if (disposing)
      {
        tw.Finish();
        if (components != null)
        {
          components.Dispose();
        }
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
            this.panel1 = new System.Windows.Forms.Panel();
            this.labelControl1 = new DevExpress.XtraEditors.LabelControl();
            this.textEdit1 = new DevExpress.XtraEditors.TextEdit();
            this.barManager1 = new DevExpress.XtraBars.BarManager(this.components);
            this.barDockControlTop = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlBottom = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlLeft = new DevExpress.XtraBars.BarDockControl();
            this.barDockControlRight = new DevExpress.XtraBars.BarDockControl();
            this.barButtonItem1 = new DevExpress.XtraBars.BarButtonItem();
            this.barButtonItem2 = new DevExpress.XtraBars.BarButtonItem();
            this.barButtonItem3 = new DevExpress.XtraBars.BarButtonItem();
            this.simpleButton5 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton4 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton3 = new DevExpress.XtraEditors.SimpleButton();
            this.simpleButton1 = new DevExpress.XtraEditors.SimpleButton();
            this.dropDownButton1 = new DevExpress.XtraEditors.DropDownButton();
            this.popupMenu1 = new DevExpress.XtraBars.PopupMenu(this.components);
            this.tbcntImage = new System.Windows.Forms.TextBox();
            this.lblcountscan = new System.Windows.Forms.Label();
            this.queriesTableAdapter1 = new Urist.uristDataSet1TableAdapters.QueriesTableAdapter();
            this.panel1.SuspendLayout();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupMenu1)).BeginInit();
            this.SuspendLayout();
            // 
            // panel1
            // 
            this.panel1.Anchor = ((System.Windows.Forms.AnchorStyles)(((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Left) 
            | System.Windows.Forms.AnchorStyles.Right)));
            this.panel1.Controls.Add(this.labelControl1);
            this.panel1.Controls.Add(this.textEdit1);
            this.panel1.Controls.Add(this.simpleButton5);
            this.panel1.Controls.Add(this.simpleButton4);
            this.panel1.Controls.Add(this.simpleButton3);
            this.panel1.Controls.Add(this.simpleButton1);
            this.panel1.Controls.Add(this.dropDownButton1);
            this.panel1.Controls.Add(this.tbcntImage);
            this.panel1.Controls.Add(this.lblcountscan);
            this.panel1.Location = new System.Drawing.Point(0, 0);
            this.panel1.Name = "panel1";
            this.panel1.Size = new System.Drawing.Size(792, 89);
            this.panel1.TabIndex = 1;
            // 
            // labelControl1
            // 
            this.labelControl1.Location = new System.Drawing.Point(372, 54);
            this.labelControl1.Name = "labelControl1";
            this.labelControl1.Size = new System.Drawing.Size(65, 13);
            this.labelControl1.TabIndex = 15;
            this.labelControl1.Text = "Примечание:";
            // 
            // textEdit1
            // 
            this.textEdit1.Location = new System.Drawing.Point(439, 51);
            this.textEdit1.MenuManager = this.barManager1;
            this.textEdit1.Name = "textEdit1";
            this.textEdit1.Size = new System.Drawing.Size(341, 20);
            this.textEdit1.TabIndex = 14;
            // 
            // barManager1
            // 
            this.barManager1.DockControls.Add(this.barDockControlTop);
            this.barManager1.DockControls.Add(this.barDockControlBottom);
            this.barManager1.DockControls.Add(this.barDockControlLeft);
            this.barManager1.DockControls.Add(this.barDockControlRight);
            this.barManager1.Form = this;
            this.barManager1.Items.AddRange(new DevExpress.XtraBars.BarItem[] {
            this.barButtonItem1,
            this.barButtonItem2,
            this.barButtonItem3});
            this.barManager1.MaxItemId = 3;
            // 
            // barDockControlTop
            // 
            this.barDockControlTop.CausesValidation = false;
            this.barDockControlTop.Dock = System.Windows.Forms.DockStyle.Top;
            this.barDockControlTop.Location = new System.Drawing.Point(0, 0);
            this.barDockControlTop.Size = new System.Drawing.Size(792, 0);
            // 
            // barDockControlBottom
            // 
            this.barDockControlBottom.CausesValidation = false;
            this.barDockControlBottom.Dock = System.Windows.Forms.DockStyle.Bottom;
            this.barDockControlBottom.Location = new System.Drawing.Point(0, 553);
            this.barDockControlBottom.Size = new System.Drawing.Size(792, 0);
            // 
            // barDockControlLeft
            // 
            this.barDockControlLeft.CausesValidation = false;
            this.barDockControlLeft.Dock = System.Windows.Forms.DockStyle.Left;
            this.barDockControlLeft.Location = new System.Drawing.Point(0, 0);
            this.barDockControlLeft.Size = new System.Drawing.Size(0, 553);
            // 
            // barDockControlRight
            // 
            this.barDockControlRight.CausesValidation = false;
            this.barDockControlRight.Dock = System.Windows.Forms.DockStyle.Right;
            this.barDockControlRight.Location = new System.Drawing.Point(792, 0);
            this.barDockControlRight.Size = new System.Drawing.Size(0, 553);
            // 
            // barButtonItem1
            // 
            this.barButtonItem1.Caption = "Выбрать сканер...";
            this.barButtonItem1.Id = 0;
            this.barButtonItem1.Name = "barButtonItem1";
            this.barButtonItem1.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem1_ItemClick);
            // 
            // barButtonItem2
            // 
            this.barButtonItem2.Caption = "Сканировать";
            this.barButtonItem2.Id = 1;
            this.barButtonItem2.Name = "barButtonItem2";
            this.barButtonItem2.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem2_ItemClick);
            // 
            // barButtonItem3
            // 
            this.barButtonItem3.Caption = "Выход";
            this.barButtonItem3.Id = 2;
            this.barButtonItem3.Name = "barButtonItem3";
            this.barButtonItem3.ItemClick += new DevExpress.XtraBars.ItemClickEventHandler(this.barButtonItem3_ItemClick);
            // 
            // simpleButton5
            // 
            this.simpleButton5.Location = new System.Drawing.Point(282, 48);
            this.simpleButton5.Name = "simpleButton5";
            this.simpleButton5.Size = new System.Drawing.Size(84, 23);
            this.simpleButton5.TabIndex = 13;
            this.simpleButton5.Text = "Отмена";
            this.simpleButton5.Click += new System.EventHandler(this.simpleButton5_Click);
            // 
            // simpleButton4
            // 
            this.simpleButton4.Location = new System.Drawing.Point(192, 48);
            this.simpleButton4.Name = "simpleButton4";
            this.simpleButton4.Size = new System.Drawing.Size(84, 23);
            this.simpleButton4.TabIndex = 12;
            this.simpleButton4.Text = "Удалить";
            this.simpleButton4.Click += new System.EventHandler(this.simpleButton4_Click);
            // 
            // simpleButton3
            // 
            this.simpleButton3.Location = new System.Drawing.Point(102, 48);
            this.simpleButton3.Name = "simpleButton3";
            this.simpleButton3.Size = new System.Drawing.Size(84, 23);
            this.simpleButton3.TabIndex = 11;
            this.simpleButton3.Text = "Сохранить ";
            this.simpleButton3.Click += new System.EventHandler(this.simpleButton3_Click);
            // 
            // simpleButton1
            // 
            this.simpleButton1.Location = new System.Drawing.Point(12, 48);
            this.simpleButton1.Name = "simpleButton1";
            this.simpleButton1.Size = new System.Drawing.Size(83, 23);
            this.simpleButton1.TabIndex = 9;
            this.simpleButton1.Text = "Сканировать";
            this.simpleButton1.Click += new System.EventHandler(this.simpleButton1_Click);
            // 
            // dropDownButton1
            // 
            this.dropDownButton1.DropDownControl = this.popupMenu1;
            this.dropDownButton1.Location = new System.Drawing.Point(12, 8);
            this.dropDownButton1.Name = "dropDownButton1";
            this.dropDownButton1.Size = new System.Drawing.Size(113, 24);
            this.dropDownButton1.TabIndex = 8;
            this.dropDownButton1.Text = "Выбрать сканер";
            // 
            // popupMenu1
            // 
            this.popupMenu1.LinksPersistInfo.AddRange(new DevExpress.XtraBars.LinkPersistInfo[] {
            new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem1),
            new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem2),
            new DevExpress.XtraBars.LinkPersistInfo(this.barButtonItem3)});
            this.popupMenu1.Manager = this.barManager1;
            this.popupMenu1.Name = "popupMenu1";
            // 
            // tbcntImage
            // 
            this.tbcntImage.Enabled = false;
            this.tbcntImage.Font = new System.Drawing.Font("Microsoft Sans Serif", 18F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(204)));
            this.tbcntImage.Location = new System.Drawing.Point(605, -4);
            this.tbcntImage.MaxLength = 1;
            this.tbcntImage.Name = "tbcntImage";
            this.tbcntImage.Size = new System.Drawing.Size(39, 35);
            this.tbcntImage.TabIndex = 4;
            this.tbcntImage.TextAlign = System.Windows.Forms.HorizontalAlignment.Center;
            this.tbcntImage.Visible = false;
            // 
            // lblcountscan
            // 
            this.lblcountscan.AutoSize = true;
            this.lblcountscan.Location = new System.Drawing.Point(650, 11);
            this.lblcountscan.Name = "lblcountscan";
            this.lblcountscan.Size = new System.Drawing.Size(137, 13);
            this.lblcountscan.TabIndex = 2;
            this.lblcountscan.Text = "Количество изображений";
            this.lblcountscan.Visible = false;
            // 
            // MainFrame
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 14);
            this.ClientSize = new System.Drawing.Size(792, 553);
            this.ControlBox = false;
            this.Controls.Add(this.panel1);
            this.Controls.Add(this.barDockControlLeft);
            this.Controls.Add(this.barDockControlRight);
            this.Controls.Add(this.barDockControlBottom);
            this.Controls.Add(this.barDockControlTop);
            this.IsMdiContainer = true;
            this.LookAndFeel.SkinName = "London Liquid Sky";
            this.Name = "MainFrame";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Сканировать";
            this.Load += new System.EventHandler(this.MainFrame_Load);
            this.panel1.ResumeLayout(false);
            this.panel1.PerformLayout();
            ((System.ComponentModel.ISupportInitialize)(this.textEdit1.Properties)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.barManager1)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.popupMenu1)).EndInit();
            this.ResumeLayout(false);

    }
    #endregion



    private void menuItemExit_Click(object sender, System.EventArgs e)
    {
      Close();
    }

    private void menuItemScan_Click(object sender, System.EventArgs e)
    {
     
    }

    private void menuItemSelSrc_Click(object sender, System.EventArgs e)
    {
      tw.Select();
    }


    bool IMessageFilter.PreFilterMessage(ref Message m)
    {
      TwainCommand cmd = tw.PassMessage(ref m);
      if (cmd == TwainCommand.Not)
        return false;

      switch (cmd)
      {
        case TwainCommand.CloseRequest:
          {
            EndingScan();
            tw.CloseSrc();
            break;
          }
        case TwainCommand.CloseOk:
          {
            EndingScan();
            tw.CloseSrc();
            break;
          }
        case TwainCommand.DeviceEvent:
          {
            break;
          }
        case TwainCommand.TransferReady:
          {
            ArrayList pics = tw.TransferPictures();
            EndingScan();
            tw.CloseSrc();
            picnumber++;
            for (int i = 0; i < pics.Count; i++)
            {
              IntPtr img = (IntPtr)pics[i];
              PicForm newpic = new PicForm(img);
              newpic.MdiParent = this;
              int picnum = i + 1;
              newpic.Text = "Изображение " + picnumber.ToString();
              if (newpic.SaveAs_() != "")
                newpic.Show();
            }
            tbcntImage.Text = picnumber.ToString();
            foreach (Form f in MdiChildren)
            {
              if (f is PicForm)
              {
                //MessageBox.Show(f.Text);
              }
            }

            break;
          }
      }

      return true;
    }

    private void EndingScan()
    {
      if (msgfilter)
      {
        Application.RemoveMessageFilter(this);
        msgfilter = false;
        this.Enabled = true;
        this.Activate();
      }
    }

    private bool msgfilter;
    private Twain tw;
    private int picnumber = 0;








    //[STAThread]
    //static void Main()
    //{
    //  if (Twain.ScreenBitDepth < 15)
    //  {
    //    MessageBox.Show("Need high/true-color video mode!", "Screen Bit Depth", MessageBoxButtons.OK, MessageBoxIcon.Information);
    //    return;
    //  }

    //  MainFrame mf = new MainFrame();
    //  Application.Run(mf);
    //}

    private void MainFrame_Load(object sender, EventArgs e)
    {

    }

    private void button1_Click(object sender, EventArgs e)
    {
    }

    private void btnScan_Click(object sender, EventArgs e)
    {
      menuItemScan_Click(sender, e);
    }

    private void btnSaveAll_Click(object sender, EventArgs e)
    {
      foreach (Form f in MdiChildren)
      {
        if (f is PicForm)
        {
          //f.Close();
        }
      }
    }

    private void btnDele_Click(object sender, EventArgs e)
    {
      string nameW = Global.isActiveForm;
      foreach (Form f in MdiChildren)
      {
        if (f is PicForm && f.Text == nameW)
        {
          //MessageBox.Show(f.Text);
          f.Close();
        }
      }
    }

    [DllImport("user32.dll")]
    static extern IntPtr GetForegroundWindow();
    [DllImport("user32.dll")]
    static extern int GetWindowText(IntPtr hWnd, StringBuilder text, int count);

    private string GetActiveWindowTitle()
    {
      const int nChars = 256;
      StringBuilder Buff = new StringBuilder(nChars);
      IntPtr handle = GetForegroundWindow();

      if (GetWindowText(handle, Buff, nChars) > 0)
      {
        return Buff.ToString();
      }
      return null;
    }

    private void btnClose_Click(object sender, EventArgs e)
    {
      Close();
    }

    private void btnSave_Click(object sender, EventArgs e)
    {

    }

    private void barButtonItem1_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
        tw.Select();
    }

    private void barButtonItem2_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
        if (!msgfilter)
        {
            this.Enabled = false;
            msgfilter = true;
            Application.AddMessageFilter(this);
        }
        tw.Acquire();
    }

    private void barButtonItem3_ItemClick(object sender, DevExpress.XtraBars.ItemClickEventArgs e)
    {
        Close();
    }

    private void simpleButton1_Click(object sender, EventArgs e)
    {
        if (!msgfilter)
        {
            this.Enabled = false;
            msgfilter = true;
            Application.AddMessageFilter(this);
        }
        tw.Acquire();
    }

    private void simpleButton3_Click(object sender, EventArgs e)
    {   PdfDocument doc = new PdfDocument();
       
        string dec = Path.GetTempPath() + Path.GetRandomFileName() + ".pdf";

        foreach (Form f in MdiChildren)
        {
      
            if (f is PicForm)
            {
               
                PicForm pp = f as PicForm;
                string str = pp.pBox.ImageLocation;
           
                    
                    PdfPage page = doc.AddPage();
                    XGraphics xgr = XGraphics.FromPdfPage(page);
                    XImage img = XImage.FromFile(str);
                    xgr.DrawImage(img, 0, 0, 210*2.7, 297*2.7);
                
                   
                   
              
               
            }
        

          
        }
        doc.Save(dec);
        string filename = Path.GetFileName(dec);
        byte[] outlen;
        using (var fs = File.Open(dec, FileMode.Open, FileAccess.Read, FileShare.None))
        {

          outlen = new byte[fs.Length];
           fs.Read(outlen, 0, Convert.ToInt32(fs.Length));

        }
        if (outlen.Length != 0)
        {
          if(Globals.click==1)
          { 
            this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, filename, null, Globals.id_user,1,null);
          }
          if (Globals.click == 2)
          {
              this.queriesTableAdapter1.insert_file(Globals.id_delo, outlen, filename, null, Globals.id_user, 1, Globals.id_sud);
          }
          if (Globals.click == 3)
          {
              this.queriesTableAdapter1.insert_file(Globals.id_delo1, outlen, filename, null, Globals.id_user, Convert.ToInt32(Globals.status_sud), Globals.id_sud);
          }
        }
        
        doc.Close();
       
        MessageBox.Show("Данные успешно сохранены");
        Close();
    }

    private void simpleButton4_Click(object sender, EventArgs e)
    {
        string nameW = Global.isActiveForm;
        foreach (Form f in MdiChildren)
        {
            if (f is PicForm && f.Text == nameW)
            {
                //MessageBox.Show(f.Text);
                f.Close();
            }
        }
    }

    private void simpleButton5_Click(object sender, EventArgs e)
    {
        Close();
    }

    private void simpleButton2_Click(object sender, EventArgs e)
    {

    }

  } // class MainFrame

} // namespace TwainGui
