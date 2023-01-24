namespace SelectAndCalculate
{
    partial class frmMain
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
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(frmMain));
            this.notMain = new System.Windows.Forms.NotifyIcon(this.components);
            this.cmsMain = new System.Windows.Forms.ContextMenuStrip(this.components);
            this.openToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.txtHotKey = new System.Windows.Forms.TextBox();
            this.label3 = new System.Windows.Forms.Label();
            this.msMain = new System.Windows.Forms.MenuStrip();
            this.fileToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.exitToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.optionsToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.runAtWindowsStartupToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem3 = new System.Windows.Forms.ToolStripSeparator();
            this.copyToClipboardResultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pasteResultToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.downloadSoftwareToolStripMenuItem = new com._4dots_software.download_software.DownloadSoftwareToolStripMenuItem();
            this.helpToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.helpGuideToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.pleaseDonateToolStripMenuItem1 = new System.Windows.Forms.ToolStripMenuItem();
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.checkForNewVersionEachWeekToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.tiHelpFeedback = new System.Windows.Forms.ToolStripMenuItem();
            this.checkForNewVersionToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.toolStripMenuItem1 = new System.Windows.Forms.ToolStripSeparator();
            this.followUsOnTwitterToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.visit4dotsSoftwareWebsiteToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.youtubeChannelToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.aboutToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.btnClose = new System.Windows.Forms.Button();
            this.label4 = new System.Windows.Forms.Label();
            this.txtTestBox = new System.Windows.Forms.TextBox();
            this.minimizeToWindowsSystemTrayToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
            this.cmsMain.SuspendLayout();
            this.msMain.SuspendLayout();
            this.SuspendLayout();
            // 
            // notMain
            // 
            this.notMain.ContextMenuStrip = this.cmsMain;
            this.notMain.Icon = ((System.Drawing.Icon)(resources.GetObject("notMain.Icon")));
            this.notMain.Text = "Select and Calculate";
            this.notMain.Visible = true;
            this.notMain.MouseDoubleClick += new System.Windows.Forms.MouseEventHandler(this.notMain_MouseDoubleClick);
            // 
            // cmsMain
            // 
            this.cmsMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.openToolStripMenuItem,
            this.exitToolStripMenuItem1});
            this.cmsMain.Name = "cmsMain";
            this.cmsMain.Size = new System.Drawing.Size(104, 48);
            // 
            // openToolStripMenuItem
            // 
            this.openToolStripMenuItem.Name = "openToolStripMenuItem";
            this.openToolStripMenuItem.Size = new System.Drawing.Size(103, 22);
            this.openToolStripMenuItem.Text = "&Open";
            this.openToolStripMenuItem.Click += new System.EventHandler(this.openToolStripMenuItem_Click);
            // 
            // exitToolStripMenuItem1
            // 
            this.exitToolStripMenuItem1.Image = global::SelectAndCalculate.Properties.Resources.exit;
            this.exitToolStripMenuItem1.Name = "exitToolStripMenuItem1";
            this.exitToolStripMenuItem1.Size = new System.Drawing.Size(103, 22);
            this.exitToolStripMenuItem1.Text = "&Exit";
            this.exitToolStripMenuItem1.Click += new System.EventHandler(this.exitToolStripMenuItem1_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.BackColor = System.Drawing.Color.Transparent;
            this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label1.ForeColor = System.Drawing.Color.DarkBlue;
            this.label1.Location = new System.Drawing.Point(121, 87);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(110, 16);
            this.label1.TabIndex = 0;
            this.label1.Text = "Shortcut Keys :";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.BackColor = System.Drawing.Color.Transparent;
            this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label2.ForeColor = System.Drawing.Color.Black;
            this.label2.Location = new System.Drawing.Point(270, 87);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(156, 16);
            this.label2.TabIndex = 1;
            this.label2.Text = "Control + C , Control +";
            // 
            // txtHotKey
            // 
            this.txtHotKey.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.txtHotKey.Location = new System.Drawing.Point(432, 84);
            this.txtHotKey.MaxLength = 1;
            this.txtHotKey.Name = "txtHotKey";
            this.txtHotKey.Size = new System.Drawing.Size(96, 22);
            this.txtHotKey.TabIndex = 2;
            this.txtHotKey.Click += new System.EventHandler(this.txtHotKey_Click);
            this.txtHotKey.Enter += new System.EventHandler(this.txtHotKey_Enter);
            this.txtHotKey.KeyDown += new System.Windows.Forms.KeyEventHandler(this.txtHotKey_KeyDown);
            this.txtHotKey.KeyPress += new System.Windows.Forms.KeyPressEventHandler(this.txtHotKey_KeyPress);
            this.txtHotKey.Validating += new System.ComponentModel.CancelEventHandler(this.txtHotKey_Validating);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.BackColor = System.Drawing.Color.Transparent;
            this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label3.ForeColor = System.Drawing.Color.DarkBlue;
            this.label3.Location = new System.Drawing.Point(21, 43);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(643, 16);
            this.label3.TabIndex = 3;
            this.label3.Text = "Type Math Expression anywhere, Select with Keyboard and Press Shortcut Keys to Ca" +
    "lculate";
            // 
            // msMain
            // 
            this.msMain.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.fileToolStripMenuItem,
            this.optionsToolStripMenuItem,
            this.downloadSoftwareToolStripMenuItem,
            this.helpToolStripMenuItem});
            this.msMain.Location = new System.Drawing.Point(0, 0);
            this.msMain.Name = "msMain";
            this.msMain.Size = new System.Drawing.Size(686, 24);
            this.msMain.TabIndex = 4;
            this.msMain.Text = "menuStrip1";
            // 
            // fileToolStripMenuItem
            // 
            this.fileToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.exitToolStripMenuItem});
            this.fileToolStripMenuItem.Name = "fileToolStripMenuItem";
            this.fileToolStripMenuItem.Size = new System.Drawing.Size(37, 20);
            this.fileToolStripMenuItem.Text = "&File";
            // 
            // exitToolStripMenuItem
            // 
            this.exitToolStripMenuItem.Image = global::SelectAndCalculate.Properties.Resources.exit;
            this.exitToolStripMenuItem.Name = "exitToolStripMenuItem";
            this.exitToolStripMenuItem.Size = new System.Drawing.Size(180, 22);
            this.exitToolStripMenuItem.Text = "&Exit";
            this.exitToolStripMenuItem.Click += new System.EventHandler(this.exitToolStripMenuItem_Click);
            // 
            // optionsToolStripMenuItem
            // 
            this.optionsToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.minimizeToWindowsSystemTrayToolStripMenuItem,
            this.runAtWindowsStartupToolStripMenuItem,
            this.toolStripMenuItem3,
            this.copyToClipboardResultToolStripMenuItem,
            this.pasteResultToolStripMenuItem});
            this.optionsToolStripMenuItem.Name = "optionsToolStripMenuItem";
            this.optionsToolStripMenuItem.Size = new System.Drawing.Size(61, 20);
            this.optionsToolStripMenuItem.Text = "&Options";
            // 
            // runAtWindowsStartupToolStripMenuItem
            // 
            this.runAtWindowsStartupToolStripMenuItem.Name = "runAtWindowsStartupToolStripMenuItem";
            this.runAtWindowsStartupToolStripMenuItem.Size = new System.Drawing.Size(254, 22);
            this.runAtWindowsStartupToolStripMenuItem.Text = "&Run at Windows Startup";
            this.runAtWindowsStartupToolStripMenuItem.Click += new System.EventHandler(this.runAtWindowsStartupToolStripMenuItem_Click);
            // 
            // toolStripMenuItem3
            // 
            this.toolStripMenuItem3.Name = "toolStripMenuItem3";
            this.toolStripMenuItem3.Size = new System.Drawing.Size(251, 6);
            // 
            // copyToClipboardResultToolStripMenuItem
            // 
            this.copyToClipboardResultToolStripMenuItem.Name = "copyToClipboardResultToolStripMenuItem";
            this.copyToClipboardResultToolStripMenuItem.Size = new System.Drawing.Size(254, 22);
            this.copyToClipboardResultToolStripMenuItem.Text = "Copy to Clipboard Result";
            this.copyToClipboardResultToolStripMenuItem.Click += new System.EventHandler(this.copyToClipboardResultToolStripMenuItem_Click);
            // 
            // pasteResultToolStripMenuItem
            // 
            this.pasteResultToolStripMenuItem.Name = "pasteResultToolStripMenuItem";
            this.pasteResultToolStripMenuItem.Size = new System.Drawing.Size(254, 22);
            this.pasteResultToolStripMenuItem.Text = "&Paste Result";
            this.pasteResultToolStripMenuItem.Click += new System.EventHandler(this.pasteResultToolStripMenuItem_Click);
            // 
            // downloadSoftwareToolStripMenuItem
            // 
            this.downloadSoftwareToolStripMenuItem.Name = "downloadSoftwareToolStripMenuItem";
            this.downloadSoftwareToolStripMenuItem.Size = new System.Drawing.Size(122, 20);
            this.downloadSoftwareToolStripMenuItem.Text = "&Download Software";
            // 
            // helpToolStripMenuItem
            // 
            this.helpToolStripMenuItem.DropDownItems.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.helpGuideToolStripMenuItem,
            this.pleaseDonateToolStripMenuItem1,
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem,
            this.checkForNewVersionEachWeekToolStripMenuItem,
            this.tiHelpFeedback,
            this.checkForNewVersionToolStripMenuItem,
            this.toolStripMenuItem1,
            this.followUsOnTwitterToolStripMenuItem,
            this.visit4dotsSoftwareWebsiteToolStripMenuItem,
            this.youtubeChannelToolStripMenuItem,
            this.aboutToolStripMenuItem});
            this.helpToolStripMenuItem.Name = "helpToolStripMenuItem";
            this.helpToolStripMenuItem.Size = new System.Drawing.Size(44, 20);
            this.helpToolStripMenuItem.Text = "&Help";
            // 
            // helpGuideToolStripMenuItem
            // 
            this.helpGuideToolStripMenuItem.Image = global::SelectAndCalculate.Properties.Resources.help2;
            this.helpGuideToolStripMenuItem.Name = "helpGuideToolStripMenuItem";
            this.helpGuideToolStripMenuItem.ShortcutKeys = System.Windows.Forms.Keys.F1;
            this.helpGuideToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.helpGuideToolStripMenuItem.Text = "Help - Users Manual";
            this.helpGuideToolStripMenuItem.Click += new System.EventHandler(this.helpGuideToolStripMenuItem_Click);
            // 
            // pleaseDonateToolStripMenuItem1
            // 
            this.pleaseDonateToolStripMenuItem1.BackColor = System.Drawing.Color.Gold;
            this.pleaseDonateToolStripMenuItem1.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.pleaseDonateToolStripMenuItem1.Name = "pleaseDonateToolStripMenuItem1";
            this.pleaseDonateToolStripMenuItem1.Size = new System.Drawing.Size(274, 22);
            this.pleaseDonateToolStripMenuItem1.Text = "&Donate via Paypal !";
            this.pleaseDonateToolStripMenuItem1.Click += new System.EventHandler(this.pleaseDonateToolStripMenuItem1_Click);
            // 
            // dotsSoftwarePRODUCTCATALOGToolStripMenuItem
            // 
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem.Font = new System.Drawing.Font("Segoe UI", 9F, System.Drawing.FontStyle.Bold);
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem.ForeColor = System.Drawing.Color.DarkBlue;
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem.Name = "dotsSoftwarePRODUCTCATALOGToolStripMenuItem";
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem.Text = "4dots Software PRODUCT CATALOG";
            this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem.Click += new System.EventHandler(this.dotsSoftwarePRODUCTCATALOGToolStripMenuItem_Click);
            // 
            // checkForNewVersionEachWeekToolStripMenuItem
            // 
            this.checkForNewVersionEachWeekToolStripMenuItem.CheckOnClick = true;
            this.checkForNewVersionEachWeekToolStripMenuItem.Name = "checkForNewVersionEachWeekToolStripMenuItem";
            this.checkForNewVersionEachWeekToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.checkForNewVersionEachWeekToolStripMenuItem.Text = "Check for new Version each Week";
            this.checkForNewVersionEachWeekToolStripMenuItem.Click += new System.EventHandler(this.checkForNewVersionEachWeekToolStripMenuItem_Click);
            // 
            // tiHelpFeedback
            // 
            this.tiHelpFeedback.Name = "tiHelpFeedback";
            this.tiHelpFeedback.Size = new System.Drawing.Size(274, 22);
            this.tiHelpFeedback.Text = "&Report a Bug or ask for a new Feature";
            this.tiHelpFeedback.Click += new System.EventHandler(this.tiHelpFeedback_Click);
            // 
            // checkForNewVersionToolStripMenuItem
            // 
            this.checkForNewVersionToolStripMenuItem.Name = "checkForNewVersionToolStripMenuItem";
            this.checkForNewVersionToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.checkForNewVersionToolStripMenuItem.Text = "Check For New Version";
            this.checkForNewVersionToolStripMenuItem.Click += new System.EventHandler(this.checkForNewVersionToolStripMenuItem_Click);
            // 
            // toolStripMenuItem1
            // 
            this.toolStripMenuItem1.Name = "toolStripMenuItem1";
            this.toolStripMenuItem1.Size = new System.Drawing.Size(271, 6);
            // 
            // followUsOnTwitterToolStripMenuItem
            // 
            this.followUsOnTwitterToolStripMenuItem.Image = global::SelectAndCalculate.Properties.Resources.twitter_24;
            this.followUsOnTwitterToolStripMenuItem.Name = "followUsOnTwitterToolStripMenuItem";
            this.followUsOnTwitterToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.followUsOnTwitterToolStripMenuItem.Text = "Follow us on &Twitter";
            this.followUsOnTwitterToolStripMenuItem.Click += new System.EventHandler(this.followUsOnTwitterToolStripMenuItem_Click);
            // 
            // visit4dotsSoftwareWebsiteToolStripMenuItem
            // 
            this.visit4dotsSoftwareWebsiteToolStripMenuItem.Image = global::SelectAndCalculate.Properties.Resources.earth;
            this.visit4dotsSoftwareWebsiteToolStripMenuItem.Name = "visit4dotsSoftwareWebsiteToolStripMenuItem";
            this.visit4dotsSoftwareWebsiteToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.visit4dotsSoftwareWebsiteToolStripMenuItem.Text = "&Visit 4dots Software Website";
            this.visit4dotsSoftwareWebsiteToolStripMenuItem.Click += new System.EventHandler(this.visit4dotsSoftwareWebsiteToolStripMenuItem_Click);
            // 
            // youtubeChannelToolStripMenuItem
            // 
            this.youtubeChannelToolStripMenuItem.Image = global::SelectAndCalculate.Properties.Resources.youtube_32;
            this.youtubeChannelToolStripMenuItem.Name = "youtubeChannelToolStripMenuItem";
            this.youtubeChannelToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.youtubeChannelToolStripMenuItem.Text = "&Youtube Channel";
            this.youtubeChannelToolStripMenuItem.Click += new System.EventHandler(this.youtubeChannelToolStripMenuItem_Click);
            // 
            // aboutToolStripMenuItem
            // 
            this.aboutToolStripMenuItem.Image = global::SelectAndCalculate.Properties.Resources.information;
            this.aboutToolStripMenuItem.Name = "aboutToolStripMenuItem";
            this.aboutToolStripMenuItem.Size = new System.Drawing.Size(274, 22);
            this.aboutToolStripMenuItem.Text = "&About";
            this.aboutToolStripMenuItem.Click += new System.EventHandler(this.aboutToolStripMenuItem_Click);
            // 
            // btnClose
            // 
            this.btnClose.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Bottom | System.Windows.Forms.AnchorStyles.Left)));
            this.btnClose.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F);
            this.btnClose.Image = global::SelectAndCalculate.Properties.Resources.check;
            this.btnClose.ImageAlign = System.Drawing.ContentAlignment.MiddleLeft;
            this.btnClose.ImeMode = System.Windows.Forms.ImeMode.NoControl;
            this.btnClose.Location = new System.Drawing.Point(273, 184);
            this.btnClose.Name = "btnClose";
            this.btnClose.Size = new System.Drawing.Size(135, 29);
            this.btnClose.TabIndex = 1;
            this.btnClose.Text = "&OK";
            this.btnClose.UseVisualStyleBackColor = true;
            this.btnClose.Click += new System.EventHandler(this.btnClose_Click);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.BackColor = System.Drawing.Color.Transparent;
            this.label4.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.label4.ForeColor = System.Drawing.Color.DarkBlue;
            this.label4.Location = new System.Drawing.Point(121, 134);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(77, 16);
            this.label4.TabIndex = 6;
            this.label4.Text = "Test Box :";
            // 
            // txtTestBox
            // 
            this.txtTestBox.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(161)));
            this.txtTestBox.Location = new System.Drawing.Point(273, 131);
            this.txtTestBox.MaxLength = 32000;
            this.txtTestBox.Name = "txtTestBox";
            this.txtTestBox.Size = new System.Drawing.Size(255, 22);
            this.txtTestBox.TabIndex = 0;
            this.txtTestBox.Text = "1+2";
            // 
            // minimizeToWindowsSystemTrayToolStripMenuItem
            // 
            this.minimizeToWindowsSystemTrayToolStripMenuItem.CheckOnClick = true;
            this.minimizeToWindowsSystemTrayToolStripMenuItem.Name = "minimizeToWindowsSystemTrayToolStripMenuItem";
            this.minimizeToWindowsSystemTrayToolStripMenuItem.Size = new System.Drawing.Size(254, 22);
            this.minimizeToWindowsSystemTrayToolStripMenuItem.Text = "&Minimize to Windows System Tray";
            this.minimizeToWindowsSystemTrayToolStripMenuItem.Click += new System.EventHandler(this.minimizeToWindowsSystemTrayToolStripMenuItem_Click);
            // 
            // frmMain
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.ClientSize = new System.Drawing.Size(686, 237);
            this.Controls.Add(this.txtTestBox);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.btnClose);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.txtHotKey);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.msMain);
            this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
            this.MainMenuStrip = this.msMain;
            this.MaximizeBox = false;
            this.Name = "frmMain";
            this.ShowInTaskbar = true;
            this.Text = "Select and Calculate";
            this.Load += new System.EventHandler(this.frmMain_Load);
            this.Resize += new System.EventHandler(this.frmMain_Resize);
            this.cmsMain.ResumeLayout(false);
            this.msMain.ResumeLayout(false);
            this.msMain.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.NotifyIcon notMain;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.TextBox txtHotKey;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.MenuStrip msMain;
        private System.Windows.Forms.ToolStripMenuItem fileToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem helpGuideToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pleaseDonateToolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem dotsSoftwarePRODUCTCATALOGToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem checkForNewVersionEachWeekToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem tiHelpFeedback;
        private System.Windows.Forms.ToolStripMenuItem checkForNewVersionToolStripMenuItem;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem1;
        private System.Windows.Forms.ToolStripMenuItem followUsOnTwitterToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem visit4dotsSoftwareWebsiteToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem youtubeChannelToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem aboutToolStripMenuItem;
        private System.Windows.Forms.Button btnClose;
        private com._4dots_software.download_software.DownloadSoftwareToolStripMenuItem downloadSoftwareToolStripMenuItem;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.TextBox txtTestBox;
        private System.Windows.Forms.ToolStripMenuItem optionsToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem runAtWindowsStartupToolStripMenuItem;
        private System.Windows.Forms.ContextMenuStrip cmsMain;
        private System.Windows.Forms.ToolStripMenuItem openToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem exitToolStripMenuItem1;
        private System.Windows.Forms.ToolStripSeparator toolStripMenuItem3;
        private System.Windows.Forms.ToolStripMenuItem copyToClipboardResultToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem pasteResultToolStripMenuItem;
        private System.Windows.Forms.ToolStripMenuItem minimizeToWindowsSystemTrayToolStripMenuItem;
    }
}
