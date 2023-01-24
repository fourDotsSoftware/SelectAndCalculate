using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Text;
using System.Windows.Forms;

namespace SelectAndCalculate
{
    public partial class frmMessageCheckbox : CustomForm
    {
        public frmMessageCheckbox()
        {
            InitializeComponent();
        }

        public frmMessageCheckbox(string title,string lblDesc):this(title,lblDesc,"")
        {

        }

        public frmMessageCheckbox(string title,string lblDesc,string chkDesc)
        {
            InitializeComponent();

            this.Text = title;

            this.lblOption.Text = lblDesc;

            if (chkDesc != String.Empty)
            {
                this.chkOption.Text = chkDesc;
            }
        }

        private void btnOK_Click(object sender, EventArgs e)
        {
            this.Close();

            if (chkOption.Checked)
            {
                this.DialogResult = DialogResult.Yes;
            }
            else
            {
                this.DialogResult = DialogResult.No;
            }
        }

        public static void ShowStartupGuideMessage()
        {
            /*
            if (Properties.Settings.Default.ShowStartupGuideMsg)
            {
                frmMessageCheckbox fm = new frmMessageCheckbox(
                    TranslateHelper.Translate("Guide"),
                    TranslateHelper.Translate("Step 1") + ". " + TranslateHelper.Translate("Enter Key to Disable in Key Textbox or Press the [Select] Button") + "\r\n\r\n" +
                    TranslateHelper.Translate("Step 2") + ". " + TranslateHelper.Translate("Press the [Add Key] Button and specify additional Settings") + "\r\n\r\n" +
                    TranslateHelper.Translate("Step 3") + ". " + TranslateHelper.Translate("Press the [Disable Keys] Button to Disable Selected Keys"));

                if (fm.ShowDialog(frmMain.Instance) == DialogResult.Yes)
                {
                    Properties.Settings.Default.ShowStartupGuideMsg = false;

                    Properties.Settings.Default.Save();
                }
            }*/
        }

        public static void ShowTimeOutValueRegistry()
        {
            /*
            if (Properties.Settings.Default.ShowTimeOutRegistryMsg)
            {
                frmMessageCheckbox fm = new frmMessageCheckbox(TranslateHelper.Translate("Registry Changes"),
                    TranslateHelper.Translate("The first time that the application is run for each user, a change to the Windows Registry is being made to ensure that disabling works. If at first, disabling does not work for you try to restart the computer to refresh Registry settings.")
                    ); ;

                if (fm.ShowDialog(frmMain.Instance) == DialogResult.Yes)
                {
                    Properties.Settings.Default.ShowTimeOutRegistryMsg = false;

                    Properties.Settings.Default.Save();
                }
            }*/
        }

        private void frmMessageCheckbox_Load(object sender, EventArgs e)
        {
            lblOption.Text =
TranslateHelper.Translate("The first time you run the application this settings window is being shown.\n\nHowever, next times the application will run hidden, minimized to the Windows system tray, unless you uncheck the option 'Minimize to System Tray'.");

            this.Text = TranslateHelper.Translate("Warning");
        }
    }
}

