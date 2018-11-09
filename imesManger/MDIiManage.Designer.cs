namespace imesManger
{
    partial class MDIiManage
    {
        /// <summary>
        /// 必需的设计器变量。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 清理所有正在使用的资源。
        /// </summary>
        /// <param name="disposing">如果应释放托管资源，为 true；否则为 false。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows 窗体设计器生成的代码

        /// <summary>
        /// 设计器支持所需的方法 - 不要
        /// 使用代码编辑器修改此方法的内容。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MDIiManage));
            this.menuStrip = new System.Windows.Forms.MenuStrip();
            this.toolStripMenuItemManage = new System.Windows.Forms.ToolStripMenuItem();
            this.acquireAToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.manufactureMToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.manageToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.importDataIToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.offToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem5 = new System.Windows.Forms.ToolStripSeparator();
            this.exitEToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.resourceRToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.produceToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.indentorIToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem2 = new System.Windows.Forms.ToolStripSeparator();
            this.preCodeRToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.viewMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.toolBarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.statusBarToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem4 = new System.Windows.Forms.ToolStripSeparator();
            this.dataBaseDToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.createDatabaseCToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.windowsMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.cascadeToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tileVerticalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tileHorizontalToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.closeAllToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.arrangeIconsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpMenu = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStrip = new System.Windows.Forms.ToolStrip();
            this.toolStripButtonAcquire = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonManufacture = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator1 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButtonManage = new System.Windows.Forms.ToolStripButton();
            this.toolStripSeparator2 = new System.Windows.Forms.ToolStripSeparator();
            this.toolStripButtonImport = new System.Windows.Forms.ToolStripButton();
            this.toolStripButtonOffline = new System.Windows.Forms.ToolStripButton();
            this.statusStrip = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelPrompt = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolTip = new System.Windows.Forms.ToolTip(this.components);
            this.toolStripStatusLabel2 = new System.Windows.Forms.ToolStripStatusLabel();
            this.toolStripStatusLabelDate = new System.Windows.Forms.ToolStripStatusLabel();
            this.menuStrip.SuspendLayout();
            this.toolStrip.SuspendLayout();
            this.statusStrip.SuspendLayout();
            this.SuspendLayout();
            // 
            // menuStrip
            // 
            this.menuStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripMenuItemManage,
            this.resourceRToolStripMenuItem,
            this.viewMenu,
            this.toolsMenu,
            this.windowsMenu,
            this.helpMenu});
            this.menuStrip.Location = new System.Drawing.Point(0, 0);
            this.menuStrip.MdiWindowListItem = this.windowsMenu;
            this.menuStrip.Name = "menuStrip";
            this.menuStrip.Size = new System.Drawing.Size(984, 25);
            this.menuStrip.TabIndex = 0;
            this.menuStrip.Text = "MenuStrip";
            // 
            // toolStripMenuItemManage
            // 
            this.toolStripMenuItemManage.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.acquireAToolStripMenuItem,
            this.manufactureMToolStripMenuItem,
            this.toolStripMenuItem1,
            this.manageToolStripMenuItem,
            this.toolStripMenuItem3,
            this.importDataIToolStripMenuItem,
            this.offToolStripMenuItem,
            this.toolStripMenuItem5,
            this.exitEToolStripMenuItem});
            this.toolStripMenuItemManage.Name = "toolStripMenuItemManage";
            this.toolStripMenuItemManage.Size = new System.Drawing.Size(88, 21);
            this.toolStripMenuItemManage.Text = "Manage(&M)";
            // 
            // acquireAToolStripMenuItem
            // 
            this.acquireAToolStripMenuItem.Image = global::imesManger.Properties.Resources._20130719053511949_easyicon_net_32;
            this.acquireAToolStripMenuItem.Name = "acquireAToolStripMenuItem";
            this.acquireAToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.acquireAToolStripMenuItem.Text = "Acquire(&A)";
            this.acquireAToolStripMenuItem.Visible = false;
            this.acquireAToolStripMenuItem.Click += new System.EventHandler(this.acquireAToolStripMenuItem_Click);
            // 
            // manufactureMToolStripMenuItem
            // 
            this.manufactureMToolStripMenuItem.Image = global::imesManger.Properties.Resources._2013071905351923_easyicon_net_32;
            this.manufactureMToolStripMenuItem.Name = "manufactureMToolStripMenuItem";
            this.manufactureMToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.manufactureMToolStripMenuItem.Text = "Used IMEI Quantity(&M) ";
            this.manufactureMToolStripMenuItem.Visible = false;
            this.manufactureMToolStripMenuItem.Click += new System.EventHandler(this.manufactureMToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(210, 6);
            this.toolStripMenuItem1.Visible = false;
            // 
            // manageToolStripMenuItem
            // 
            this.manageToolStripMenuItem.Image = global::imesManger.Properties.Resources._20130719053501383_easyicon_net_32;
            this.manageToolStripMenuItem.Name = "manageToolStripMenuItem";
            this.manageToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.manageToolStripMenuItem.Text = "Forcast(&F)";
            this.manageToolStripMenuItem.Visible = false;
            this.manageToolStripMenuItem.Click += new System.EventHandler(this.manageToolStripMenuItem_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(210, 6);
            this.toolStripMenuItem3.Visible = false;
            // 
            // importDataIToolStripMenuItem
            // 
            this.importDataIToolStripMenuItem.Image = ((System.Drawing.Image)(resources.GetObject("importDataIToolStripMenuItem.Image")));
            this.importDataIToolStripMenuItem.Name = "importDataIToolStripMenuItem";
            this.importDataIToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.importDataIToolStripMenuItem.Text = "Import Data(&I)";
            this.importDataIToolStripMenuItem.Visible = false;
            this.importDataIToolStripMenuItem.Click += new System.EventHandler(this.importDataIToolStripMenuItem_Click);
            // 
            // offToolStripMenuItem
            // 
            this.offToolStripMenuItem.Image = global::imesManger.Properties.Resources._20130802014223509_easyicon_net_32;
            this.offToolStripMenuItem.Name = "offToolStripMenuItem";
            this.offToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.offToolStripMenuItem.Text = "Import data Offline(&O)";
            this.offToolStripMenuItem.Click += new System.EventHandler(this.offToolStripMenuItem_Click);
            // 
            // toolStripMenuItem5
            // 
            this.toolStripMenuItem5.Name = "toolStripMenuItem5";
            this.toolStripMenuItem5.Size = new System.Drawing.Size(210, 6);
            this.toolStripMenuItem5.Visible = false;
            // 
            // exitEToolStripMenuItem
            // 
            this.exitEToolStripMenuItem.Name = "exitEToolStripMenuItem";
            this.exitEToolStripMenuItem.Size = new System.Drawing.Size(213, 22);
            this.exitEToolStripMenuItem.Text = "Exit(&X)";
            this.exitEToolStripMenuItem.Click += new System.EventHandler(this.exitEToolStripMenuItem_Click);
            // 
            // resourceRToolStripMenuItem
            // 
            this.resourceRToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.produceToolStripMenuItem,
            this.indentorIToolStripMenuItem,
            this.toolStripMenuItem2,
            this.preCodeRToolStripMenuItem});
            this.resourceRToolStripMenuItem.Name = "resourceRToolStripMenuItem";
            this.resourceRToolStripMenuItem.Size = new System.Drawing.Size(90, 21);
            this.resourceRToolStripMenuItem.Text = "Resource(&R)";
            this.resourceRToolStripMenuItem.Visible = false;
            // 
            // produceToolStripMenuItem
            // 
            this.produceToolStripMenuItem.Name = "produceToolStripMenuItem";
            this.produceToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.produceToolStripMenuItem.Text = "Product(&P)";
            this.produceToolStripMenuItem.Click += new System.EventHandler(this.produceToolStripMenuItem_Click);
            // 
            // indentorIToolStripMenuItem
            // 
            this.indentorIToolStripMenuItem.Name = "indentorIToolStripMenuItem";
            this.indentorIToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.indentorIToolStripMenuItem.Text = "Customer(&C)";
            this.indentorIToolStripMenuItem.Click += new System.EventHandler(this.indentorIToolStripMenuItem_Click);
            // 
            // toolStripMenuItem2
            // 
            this.toolStripMenuItem2.Name = "toolStripMenuItem2";
            this.toolStripMenuItem2.Size = new System.Drawing.Size(145, 6);
            // 
            // preCodeRToolStripMenuItem
            // 
            this.preCodeRToolStripMenuItem.Name = "preCodeRToolStripMenuItem";
            this.preCodeRToolStripMenuItem.Size = new System.Drawing.Size(148, 22);
            this.preCodeRToolStripMenuItem.Text = "TAC(&T)";
            this.preCodeRToolStripMenuItem.Click += new System.EventHandler(this.preCodeRToolStripMenuItem_Click);
            // 
            // viewMenu
            // 
            this.viewMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolBarToolStripMenuItem,
            this.statusBarToolStripMenuItem});
            this.viewMenu.Name = "viewMenu";
            this.viewMenu.Size = new System.Drawing.Size(69, 21);
            this.viewMenu.Text = "Views(&V)";
            // 
            // toolBarToolStripMenuItem
            // 
            this.toolBarToolStripMenuItem.Checked = true;
            this.toolBarToolStripMenuItem.CheckOnClick = true;
            this.toolBarToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.toolBarToolStripMenuItem.Name = "toolBarToolStripMenuItem";
            this.toolBarToolStripMenuItem.Size = new System.Drawing.Size(137, 22);
            this.toolBarToolStripMenuItem.Text = "ToolBar(&T)";
            this.toolBarToolStripMenuItem.Click += new System.EventHandler(this.ToolBarToolStripMenuItem_Click);
            // 
            // statusBarToolStripMenuItem
            // 
            this.statusBarToolStripMenuItem.Checked = true;
            this.statusBarToolStripMenuItem.CheckOnClick = true;
            this.statusBarToolStripMenuItem.CheckState = System.Windows.Forms.CheckState.Checked;
            this.statusBarToolStripMenuItem.Name = "statusBarToolStripMenuItem";
            this.statusBarToolStripMenuItem.Size = new System.Drawing.Size(137, 22);
            this.statusBarToolStripMenuItem.Text = "Status(&S)";
            this.statusBarToolStripMenuItem.Click += new System.EventHandler(this.StatusBarToolStripMenuItem_Click);
            // 
            // toolsMenu
            // 
            this.toolsMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.optionsToolStripMenuItem,
            this.toolStripMenuItem4,
            this.dataBaseDToolStripMenuItem,
            this.createDatabaseCToolStripMenuItem});
            this.toolsMenu.Name = "toolsMenu";
            this.toolsMenu.Size = new System.Drawing.Size(67, 21);
            this.toolsMenu.Text = "Tools(&T)";
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.optionsToolStripMenuItem.Text = "Options(&O)";
            this.optionsToolStripMenuItem.Click += new System.EventHandler(this.optionsToolStripMenuItem_Click);
            // 
            // toolStripMenuItem4
            // 
            this.toolStripMenuItem4.Name = "toolStripMenuItem4";
            this.toolStripMenuItem4.Size = new System.Drawing.Size(186, 6);
            this.toolStripMenuItem4.Visible = false;
            // 
            // dataBaseDToolStripMenuItem
            // 
            this.dataBaseDToolStripMenuItem.Name = "dataBaseDToolStripMenuItem";
            this.dataBaseDToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.dataBaseDToolStripMenuItem.Text = "Set Database(&D)";
            this.dataBaseDToolStripMenuItem.Visible = false;
            this.dataBaseDToolStripMenuItem.Click += new System.EventHandler(this.dataBaseDToolStripMenuItem_Click);
            // 
            // createDatabaseCToolStripMenuItem
            // 
            this.createDatabaseCToolStripMenuItem.Name = "createDatabaseCToolStripMenuItem";
            this.createDatabaseCToolStripMenuItem.Size = new System.Drawing.Size(189, 22);
            this.createDatabaseCToolStripMenuItem.Text = "Create Database(&C)";
            this.createDatabaseCToolStripMenuItem.Visible = false;
            this.createDatabaseCToolStripMenuItem.Click += new System.EventHandler(this.createDatabaseCToolStripMenuItem_Click);
            // 
            // windowsMenu
            // 
            this.windowsMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.cascadeToolStripMenuItem,
            this.tileVerticalToolStripMenuItem,
            this.tileHorizontalToolStripMenuItem,
            this.closeAllToolStripMenuItem,
            this.arrangeIconsToolStripMenuItem});
            this.windowsMenu.Name = "windowsMenu";
            this.windowsMenu.Size = new System.Drawing.Size(93, 21);
            this.windowsMenu.Text = "Windows(&W)";
            // 
            // cascadeToolStripMenuItem
            // 
            this.cascadeToolStripMenuItem.Name = "cascadeToolStripMenuItem";
            this.cascadeToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.cascadeToolStripMenuItem.Text = "Cascade(&C)";
            this.cascadeToolStripMenuItem.Click += new System.EventHandler(this.CascadeToolStripMenuItem_Click);
            // 
            // tileVerticalToolStripMenuItem
            // 
            this.tileVerticalToolStripMenuItem.Name = "tileVerticalToolStripMenuItem";
            this.tileVerticalToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.tileVerticalToolStripMenuItem.Text = "Tile Vertical(&V)";
            this.tileVerticalToolStripMenuItem.Click += new System.EventHandler(this.TileVerticalToolStripMenuItem_Click);
            // 
            // tileHorizontalToolStripMenuItem
            // 
            this.tileHorizontalToolStripMenuItem.Name = "tileHorizontalToolStripMenuItem";
            this.tileHorizontalToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.tileHorizontalToolStripMenuItem.Text = "Tile Horizontal(&H)";
            this.tileHorizontalToolStripMenuItem.Click += new System.EventHandler(this.TileHorizontalToolStripMenuItem_Click);
            // 
            // closeAllToolStripMenuItem
            // 
            this.closeAllToolStripMenuItem.Name = "closeAllToolStripMenuItem";
            this.closeAllToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.closeAllToolStripMenuItem.Text = "Close All(&L)";
            this.closeAllToolStripMenuItem.Click += new System.EventHandler(this.CloseAllToolStripMenuItem_Click);
            // 
            // arrangeIconsToolStripMenuItem
            // 
            this.arrangeIconsToolStripMenuItem.Name = "arrangeIconsToolStripMenuItem";
            this.arrangeIconsToolStripMenuItem.Size = new System.Drawing.Size(177, 22);
            this.arrangeIconsToolStripMenuItem.Text = "Arrange Icons(&A)";
            this.arrangeIconsToolStripMenuItem.Click += new System.EventHandler(this.ArrangeIconsToolStripMenuItem_Click);
            // 
            // helpMenu
            // 
            this.helpMenu.Alignment = System.Windows.Forms.ToolStripItemAlignment.Right;
            this.helpMenu.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.aboutToolStripMenuItem});
            this.helpMenu.Name = "helpMenu";
            this.helpMenu.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.helpMenu.Size = new System.Drawing.Size(64, 21);
            this.helpMenu.Text = "Help(&H)";
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(153, 22);
            this.aboutToolStripMenuItem.Text = "About(&A) ... ...";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // toolStrip
            // 
            this.toolStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripButtonAcquire,
            this.toolStripButtonManufacture,
            this.toolStripSeparator1,
            this.toolStripButtonManage,
            this.toolStripSeparator2,
            this.toolStripButtonImport,
            this.toolStripButtonOffline});
            this.toolStrip.Location = new System.Drawing.Point(0, 25);
            this.toolStrip.Name = "toolStrip";
            this.toolStrip.Size = new System.Drawing.Size(984, 25);
            this.toolStrip.TabIndex = 1;
            this.toolStrip.Text = "ToolStrip";
            // 
            // toolStripButtonAcquire
            // 
            this.toolStripButtonAcquire.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonAcquire.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonAcquire.Image")));
            this.toolStripButtonAcquire.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonAcquire.Name = "toolStripButtonAcquire";
            this.toolStripButtonAcquire.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonAcquire.Text = "Acquire";
            this.toolStripButtonAcquire.Visible = false;
            this.toolStripButtonAcquire.Click += new System.EventHandler(this.toolStripButtonAcquire_Click);
            // 
            // toolStripButtonManufacture
            // 
            this.toolStripButtonManufacture.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonManufacture.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonManufacture.Image")));
            this.toolStripButtonManufacture.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonManufacture.Name = "toolStripButtonManufacture";
            this.toolStripButtonManufacture.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonManufacture.Text = "Manufacture";
            this.toolStripButtonManufacture.Visible = false;
            this.toolStripButtonManufacture.Click += new System.EventHandler(this.toolStripButtonManufacture_Click);
            // 
            // toolStripSeparator1
            // 
            this.toolStripSeparator1.Name = "toolStripSeparator1";
            this.toolStripSeparator1.Size = new System.Drawing.Size(6, 25);
            this.toolStripSeparator1.Visible = false;
            // 
            // toolStripButtonManage
            // 
            this.toolStripButtonManage.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonManage.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonManage.Image")));
            this.toolStripButtonManage.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonManage.Name = "toolStripButtonManage";
            this.toolStripButtonManage.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonManage.Text = "Forecast";
            this.toolStripButtonManage.Visible = false;
            this.toolStripButtonManage.Click += new System.EventHandler(this.toolStripButtonManage_Click);
            // 
            // toolStripSeparator2
            // 
            this.toolStripSeparator2.Name = "toolStripSeparator2";
            this.toolStripSeparator2.Size = new System.Drawing.Size(6, 25);
            this.toolStripSeparator2.Visible = false;
            // 
            // toolStripButtonImport
            // 
            this.toolStripButtonImport.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonImport.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonImport.Image")));
            this.toolStripButtonImport.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonImport.Name = "toolStripButtonImport";
            this.toolStripButtonImport.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonImport.Text = "Import data";
            this.toolStripButtonImport.Visible = false;
            this.toolStripButtonImport.Click += new System.EventHandler(this.toolStripButtonImport_Click);
            // 
            // toolStripButtonOffline
            // 
            this.toolStripButtonOffline.DisplayStyle = System.Windows.Forms.ToolStripItemDisplayStyle.Image;
            this.toolStripButtonOffline.Image = ((System.Drawing.Image)(resources.GetObject("toolStripButtonOffline.Image")));
            this.toolStripButtonOffline.ImageTransparentColor = System.Drawing.Color.Magenta;
            this.toolStripButtonOffline.Name = "toolStripButtonOffline";
            this.toolStripButtonOffline.Size = new System.Drawing.Size(23, 22);
            this.toolStripButtonOffline.Text = "Import data Offline";
            this.toolStripButtonOffline.Click += new System.EventHandler(this.toolStripButtonOffline_Click);
            // 
            // statusStrip
            // 
            this.statusStrip.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel,
            this.toolStripStatusLabelPrompt,
            this.toolStripStatusLabel2,
            this.toolStripStatusLabelDate});
            this.statusStrip.Location = new System.Drawing.Point(0, 740);
            this.statusStrip.Name = "statusStrip";
            this.statusStrip.Size = new System.Drawing.Size(984, 22);
            this.statusStrip.TabIndex = 2;
            this.statusStrip.Text = "StatusStrip";
            // 
            // toolStripStatusLabel
            // 
            this.toolStripStatusLabel.Name = "toolStripStatusLabel";
            this.toolStripStatusLabel.Size = new System.Drawing.Size(0, 17);
            // 
            // toolStripStatusLabelPrompt
            // 
            this.toolStripStatusLabelPrompt.Name = "toolStripStatusLabelPrompt";
            this.toolStripStatusLabelPrompt.Size = new System.Drawing.Size(173, 17);
            this.toolStripStatusLabelPrompt.Text = "Open a upload file and start";
            // 
            // toolStripStatusLabel2
            // 
            this.toolStripStatusLabel2.Name = "toolStripStatusLabel2";
            this.toolStripStatusLabel2.Size = new System.Drawing.Size(780, 17);
            this.toolStripStatusLabel2.Spring = true;
            // 
            // toolStripStatusLabelDate
            // 
            this.toolStripStatusLabelDate.Name = "toolStripStatusLabelDate";
            this.toolStripStatusLabelDate.Size = new System.Drawing.Size(16, 17);
            this.toolStripStatusLabelDate.Text = "  ";
            // 
            // MDIiManage
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(984, 762);
            this.Controls.Add(this.statusStrip);
            this.Controls.Add(this.toolStrip);
            this.Controls.Add(this.menuStrip);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.IsMdiContainer = true;
            this.MainMenuStrip = this.menuStrip;
            this.Name = "MDIiManage";
            this.Text = "IMEI Forecast ";
            this.Load += new System.EventHandler(this.MDIiManage_Load);
            this.menuStrip.ResumeLayout(false);
            this.menuStrip.PerformLayout();
            this.toolStrip.ResumeLayout(false);
            this.toolStrip.PerformLayout();
            this.statusStrip.ResumeLayout(false);
            this.statusStrip.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }
        #endregion


        private System.Windows.Forms.MenuStrip menuStrip;
        private System.Windows.Forms.ToolStrip toolStrip;
        private System.Windows.Forms.StatusStrip statusStrip;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tileHorizontalToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem viewMenu;
        private System.Windows.Forms.ToolStripMenuItem toolBarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem statusBarToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem toolsMenu;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem windowsMenu;
        private System.Windows.Forms.ToolStripMenuItem cascadeToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tileVerticalToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem closeAllToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem arrangeIconsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpMenu;
        private System.Windows.Forms.ToolTip toolTip;
        private System.Windows.Forms.ToolStripMenuItem toolStripMenuItemManage;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem exitEToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem resourceRToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem produceToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem indentorIToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem2;
        private System.Windows.Forms.ToolStripMenuItem preCodeRToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem acquireAToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem manufactureMToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem manageToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripButton toolStripButtonAcquire;
        private System.Windows.Forms.ToolStripButton toolStripButtonManufacture;
        private System.Windows.Forms.ToolStripButton toolStripButtonManage;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator1;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem4;
        private System.Windows.Forms.ToolStripMenuItem dataBaseDToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem importDataIToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem5;
        private System.Windows.Forms.ToolStripSeparator toolStripSeparator2;
        private System.Windows.Forms.ToolStripButton toolStripButtonImport;
        private System.Windows.Forms.ToolStripMenuItem createDatabaseCToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem offToolStripMenuItem;
        private System.Windows.Forms.ToolStripButton toolStripButtonOffline;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelPrompt;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabel2;
        private System.Windows.Forms.ToolStripStatusLabel toolStripStatusLabelDate;
    }
}



