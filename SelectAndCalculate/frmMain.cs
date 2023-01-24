using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;
using Jace;
using System.Diagnostics;
using System.Linq;

namespace SelectAndCalculate
{
    public partial class frmMain : SelectAndCalculate.CustomForm
    {
        KeyboardHook keyboardHook = new KeyboardHook();

        private int HotKeyChar = -1;

        public static bool CheckSetHookTimer=false;

        public static frmMain Instance = null;

        HookingProtector HookingProtector = new HookingProtector();

        public frmMain()
        {
            InitializeComponent();

            ///Properties.Settings.Default.Initialized = false;

            if (Properties.Settings.Default.Initialized && Properties.Settings.Default.MinimizeToTray)
            {
                this.Visible = false;
                WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }

            Instance = this;
        }

        private void frmMain_Load(object sender, EventArgs e)
        {
            keyboardHook.KeyDown += new KeyEventHandler(keyboardHook_KeyDown);
            keyboardHook.KeyUp += new KeyEventHandler(keyboardHook_KeyUp);
            keyboardHook.KeyPress += new KeyPressEventHandler(keyboardHook_KeyPress);

            txtHotKey.Text = Properties.Settings.Default.ShortcutKeyString;

            HotKeyChar = Properties.Settings.Default.ShortcutKeyString[0];

            if (Properties.Settings.Default.CheckWeek)
            {
                UpdateHelper.InitializeCheckVersionWeek();
            }

            checkForNewVersionEachWeekToolStripMenuItem.Checked = Properties.Settings.Default.CheckWeek;
            minimizeToWindowsSystemTrayToolStripMenuItem.Checked = Properties.Settings.Default.MinimizeToTray;
            copyToClipboardResultToolStripMenuItem.Checked = !Properties.Settings.Default.PasteResult;
            pasteResultToolStripMenuItem.Checked = Properties.Settings.Default.PasteResult;

            keyboardHook.Start();

            if (!Properties.Settings.Default.Initialized)
            {
                RunAtWndowsStartupManager.RunAtWindowsStartup = true;

                frmMessageCheckbox fm = new frmMessageCheckbox();
                fm.Show(this);

            }

            runAtWindowsStartupToolStripMenuItem.Checked = RunAtWndowsStartupManager.RunAtWindowsStartup;

            if (Properties.Settings.Default.Initialized && Properties.Settings.Default.MinimizeToTray)
            {
                btnClose_Click(null, null);
            }

            SetTitle();

            if (Properties.Settings.Default.Initialized && Properties.Settings.Default.MinimizeToTray)
            {
                this.Visible = false;
                this.Hide();
                WindowState = FormWindowState.Minimized;
                this.ShowInTaskbar = false;
            }
            else
            {
                this.ShowInTaskbar = true;
            }

            if (!Properties.Settings.Default.Initialized)
            {
                Properties.Settings.Default.Initialized = true;
                Properties.Settings.Default.Save();
            }

            for (int k = 0; k < Module.args.Length; k++)
            {
                if (Module.args[k].Trim() == "/minimized")
                {
                    this.WindowState = FormWindowState.Minimized;
                }
                else if (Module.args[k].Trim() == "/novisible")
                {
                    this.Visible = false;

                    if (Properties.Settings.Default.MinimizeToTray)
                    {
                        notMain.Visible = true;
                    }
                }                
            }

            HookingProtector.Setup();

        }

        private bool ControlCIsPressed = false;

        void keyboardHook_KeyPress(object sender, KeyPressEventArgs e)
        {
            
        }

        void keyboardHook_KeyUp(object sender, KeyEventArgs e)
        {
            if ((char)e.KeyCode == 'C')
            {
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    ControlCIsPressed = true;                    
                }
            }           
        }

        void keyboardHook_KeyDown(object sender, KeyEventArgs e)
        {
            /*
            if (e.KeyData==Keys.F14)
            {
                frmMain.CheckSetHookTimer = false;

                return;
            }
            */

            if ((e.KeyValue==Properties.Settings.Default.ShortcutKey) && ControlCIsPressed && ((Control.ModifierKeys & Keys.Control) == Keys.Control))            
            {                
                e.Handled = true;                
                e.SuppressKeyPress = true;

                Calculate();
            }

            if ((char)e.KeyCode == 'C')
            {
                if ((Control.ModifierKeys & Keys.Control) == Keys.Control)
                {
                    ControlCIsPressed = true;
                }
                else
                {
                    ControlCIsPressed = false;
                }
            }
            else
            {
                ControlCIsPressed = false;
            }            
        }

        private void Calculate()
        {
            try
            {
                string str = Clipboard.GetText();

                CalculationEngine engine = new CalculationEngine();
                engine.AddFunction("sum", Sum);
                engine.AddFunction("count", Count);
                engine.AddFunction("in", In);
                engine.AddFunction("stdev", StandardDeviation);
                engine.AddFunction("var", Variance);
                
                double result = engine.Calculate(str);

                Clipboard.Clear();
                Clipboard.SetText(result.ToString());

                if (Properties.Settings.Default.PasteResult)
                {
                    KeyboardSimulator.SimulateStandardShortcut(StandardShortcut.Paste);
                }

            }
            catch
            {
                Clipboard.Clear();
                Clipboard.SetText("Invalid Math Expression");

                if (Properties.Settings.Default.PasteResult)
                {
                    KeyboardSimulator.SimulateStandardShortcut(StandardShortcut.Paste);
                }

                return;
            }
        }


        #region Calculation Engine Functions

        double Sum(params double[] a)
        {
            double sum = 0;

            for (int k=0;k<a.Length;k++)
            {
                sum += a[k];
            }

            return sum;
        }

        double Count(params double[] a)
        {
            return a.Length;
        }

        double In(params double[] a)
        {
            for (int k=1;k<a.Length;k++)
            {
                if (a[k]==a[0])
                {
                    return 1;
                }
            }

            return 0;
        }

        double Variance(params double[] nums)
        {
            if (nums.Length > 1)
            {

                // Get the average of the values
                double avg = nums.Average();

                // Now figure out how far each point is from the mean
                // So we subtract from the number the average
                // Then raise it to the power of 2
                double sumOfSquares = 0.0;

                foreach (double num in nums)
                {
                    sumOfSquares += Math.Pow((num - avg), 2.0);
                }

                // Finally divide it by n - 1 (for standard deviation variance)
                // Or use length without subtracting one ( for population standard deviation variance)
                return sumOfSquares / (double)(nums.Length - 1);
            }
            else { return 0.0; }
        }
        double StandardDeviation(params double[] values)
        {
            double standardDeviation = 0;

            if (values.Length>0)
            {
                // Compute the average.     
                double avg = values.Average();

                // Perform the Sum of (value-avg)_2_2.      
                double sum = values.Sum(d => Math.Pow(d - avg, 2));

                // Put it all together.      
                standardDeviation = Math.Sqrt((sum) / (values.Count() - 1));
            }

            return standardDeviation;
        }

        double Abs(params double[] a)
        {
            return Math.Abs(a[0]);
        }

        #endregion


        private void txtHotKey_KeyPress(object sender, KeyPressEventArgs e)
        {
            
            
        }

        #region Help Menu

        private void helpGuideToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start(Module.HelpURL);
        }

        private void pleaseDonateToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/donate.php");
        }

        private void dotsSoftwarePRODUCTCATALOGToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("http://www.4dots-software.com/downloads/4dots-Software-PRODUCT-CATALOG.pdf");
        }

        private void checkForNewVersionEachWeekToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.CheckWeek = checkForNewVersionEachWeekToolStripMenuItem.Checked;
            Properties.Settings.Default.Save();
        }

        private void tiHelpFeedback_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.4dots-software.com/support/bugfeature.php?app=" + System.Web.HttpUtility.UrlEncode(Module.ShortApplicationTitle));
        }

        private void checkForNewVersionToolStripMenuItem_Click(object sender, EventArgs e)
        {
            UpdateHelper.CheckVersion(false);
        }

        private void followUsOnTwitterToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Process.Start("https://www.twitter.com/4dotsSoftware");
        }

        private void visit4dotsSoftwareWebsiteToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.4dots-software.com");
        }

        private void youtubeChannelToolStripMenuItem_Click(object sender, EventArgs e)
        {
            System.Diagnostics.Process.Start("https://www.youtube.com/channel/UCovA-lld9Q79l08K-V1QEng");
        }

        private void aboutToolStripMenuItem_Click(object sender, EventArgs e)
        {
            frmAbout f = new frmAbout();
            f.ShowDialog();
        }

        private void exitToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        #endregion

        private void minimizeToWindowsSystemTrayToolStripMenuItem_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.MinimizeToTray = minimizeToWindowsSystemTrayToolStripMenuItem.Checked;
            Properties.Settings.Default.Save();
        }

        private void notMain_MouseDoubleClick(object sender, MouseEventArgs e)
        {            
            this.Visible = true;
            this.Show();
            this.BringToFront();
            this.WindowState = FormWindowState.Normal;
            this.ShowInTaskbar = true;
            this.CenterToScreen();
        }

        private void btnClose_Click(object sender, EventArgs e)
        {
            Properties.Settings.Default.ShortcutKeyString = txtHotKey.Text;

            HotKeyChar = Properties.Settings.Default.ShortcutKeyString[0];

            Properties.Settings.Default.Save();

            this.Visible = !Properties.Settings.Default.MinimizeToTray;

            if (Properties.Settings.Default.MinimizeToTray)
            {
                this.Hide();
                WindowState = FormWindowState.Minimized;
            }

            this.notMain.Visible = Properties.Settings.Default.MinimizeToTray;
        }

        private void txtHotKey_Enter(object sender, EventArgs e)
        {            
            txtHotKey.SelectAll();
        }

        private void frmMain_Resize(object sender, EventArgs e)
        {
            if (WindowState==FormWindowState.Minimized)
            {
                if (Properties.Settings.Default.MinimizeToTray)
                {
                    notMain.Visible = true;
                    this.Visible = false;
                }
            }
        }

        private void txtHotKey_Click(object sender, EventArgs e)
        {
            txtHotKey_Enter(null, null);
        }

        private void txtHotKey_Validating(object sender, CancelEventArgs e)
        {
            
        }

        private void txtHotKey_KeyDown(object sender, KeyEventArgs e)
        {
            int vkCode = (int)e.KeyData;

            int if13 = (int)Keys.F13;
            int if14 = (int)Keys.F14;
            int if15 = (int)Keys.F15;
            int if16 = (int)Keys.F16;
            int if17 = (int)Keys.F17;

            if ((vkCode == if13) || (vkCode == if14) || (vkCode == if15) || (vkCode == if16) || (vkCode == if17))
            {
                e.Handled = true;

                return;
            }

            Properties.Settings.Default.ShortcutKey = e.KeyValue;

            Properties.Settings.Default.ShortcutKeyString = e.KeyCode.ToString();

            Properties.Settings.Default.Save();
            
            txtHotKey.Text = e.KeyCode.ToString();

            
        }

        private void exitToolStripMenuItem1_Click(object sender, EventArgs e)
        {
            Application.Exit();
        }

        private void openToolStripMenuItem_Click(object sender, EventArgs e)
        {
            notMain_MouseDoubleClick(null, null);
        }

        private void runAtWindowsStartupToolStripMenuItem_Click(object sender, EventArgs e)
        {
            RunAtWndowsStartupManager.RunAtWindowsStartup = !runAtWindowsStartupToolStripMenuItem.Checked;

            runAtWindowsStartupToolStripMenuItem.Checked = RunAtWndowsStartupManager.RunAtWindowsStartup;

        }

        bool FreeForPersonalUse = false;
        bool FreeForPersonalAndCommercialUse = true;

        private void SetTitle()
        {
            string str = "";
                        
            if (FreeForPersonalUse)
            {
                str += " - " + TranslateHelper.Translate("Free for Personal Use Only - Please Donate !");
            }
            else if (FreeForPersonalAndCommercialUse)
            {
                str += " - " + TranslateHelper.Translate("Free for Personal and Commercial Use - Please Donate !");
            }

            this.Text = Module.ApplicationTitle + str.ToUpper();
        }

        private void copyToClipboardResultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            copyToClipboardResultToolStripMenuItem.Checked = !copyToClipboardResultToolStripMenuItem.Checked;
            Properties.Settings.Default.PasteResult = !copyToClipboardResultToolStripMenuItem.Checked;
            pasteResultToolStripMenuItem.Checked = !copyToClipboardResultToolStripMenuItem.Checked;

            Properties.Settings.Default.Save();
        }

        private void pasteResultToolStripMenuItem_Click(object sender, EventArgs e)
        {
            pasteResultToolStripMenuItem.Checked = !pasteResultToolStripMenuItem.Checked;
            Properties.Settings.Default.PasteResult = pasteResultToolStripMenuItem.Checked;
            copyToClipboardResultToolStripMenuItem.Checked = !pasteResultToolStripMenuItem.Checked;

            Properties.Settings.Default.Save();
        }
    }
}
