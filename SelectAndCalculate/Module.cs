using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.Net;

namespace SelectAndCalculate
{
    class Module
    {
        public static System.Drawing.Color BlueForeColor = System.Drawing.Color.FromArgb(52, 89, 152);

        public static string ApplicationName = "Select and Calculate";
        public static string Version = "1.0";

        public static string Ver = "2";

        public static string ShortApplicationTitle = ApplicationName + " V" + Version;
        public static string ApplicationTitle = ShortApplicationTitle + " - 4dots Software";                
        
        public static string DownloadURL = "https://www.4dots-software.com/d/SelectAndCalculate/";
        public static string HelpURL = "https://www.4dots-software.com/select-and-calculate/how-to-use.php";
        public static string ProductWebpageURL = "https://www.4dots-software.com/select-and-calculate/";
        public static string BuyURL = "https://www.4dots-software.com/store/buy-select-and-calculate.php";
        public static string VersionURL = "http://cssspritestool.com/versions/select-and-calculate.txt";

        public static int NumberOfFilesProcessed = 0;

        public static string TipText = TranslateHelper.Translate("Great application to select math expression and calculate");                   

        public static string SelectedLanguage = "";
        
        public static string[] args = null;
        

        [DllImport("user32.dll")]
        public static extern int SendMessage(IntPtr hWnd, int wMsg, int wParam,
        int lParam);

        [DllImport("user32.dll")]
        public static extern bool LockWindowUpdate(IntPtr hWndLock);

        public static void WaitNMSeconds(int mseconds)
        {
            if (mseconds < 1) return;
            DateTime _desired = DateTime.Now.AddMilliseconds(mseconds);
            while (DateTime.Now < _desired)
            {
                System.Windows.Forms.Application.DoEvents();
            }
        }

        public static void ShowMessage(string msg)
        {
            MessageBox.Show(TranslateHelper.Translate(msg));

        }

        public static DialogResult ShowQuestionDialog(string msg, string caption)
        {
            return MessageBox.Show(TranslateHelper.Translate(msg), TranslateHelper.Translate(caption), MessageBoxButtons.YesNo, MessageBoxIcon.Question,MessageBoxDefaultButton.Button2);
        }


        public static void ShowError(Exception ex)
        {
            ShowError("Error", ex);
        }

        public static void ShowError(string msg)
        {
            
                try
                {
                    frmError f = new frmError("Error", msg);
                    f.ShowDialog();
                }
                catch
                {

                }
                //MessageBox.Show(msg, "Error", MessageBoxButtons.OK, MessageBoxIcon.Error);
            
                }

        public static void ShowError(string msg, Exception ex)
        {
            //ShowError(msg + "\n\n" + ex.Message);
            ShowError(msg + "\n\n" + ex.ToString());
        }

        public static void ShowError(string msg, string exstr)
        {
            ShowError(msg + "\n\n" + exstr);
        }

        

        public static DialogResult ShowQuestionDialogYesFocus(string msg, string caption)
        {
            return MessageBox.Show(TranslateHelper.Translate(msg), TranslateHelper.Translate(caption), MessageBoxButtons.YesNo, MessageBoxIcon.Question, MessageBoxDefaultButton.Button1);
        }
        

        

        public static string DownloadPage(string uri)
        {
            try
            {
                WebClient client = new WebClient();

                Stream data = client.OpenRead(uri);
                StreamReader reader = new StreamReader(data);
                string s = reader.ReadToEnd();
                data.Close();
                reader.Close();
                return s;
            }
            catch (Exception ex)
            {
                return "Error";
            }
        }        

        public static bool RunAdminAction(string args)
        {
            try
            {
                System.Diagnostics.Process pr = new System.Diagnostics.Process();
                pr.StartInfo.FileName = Application.StartupPath + "\\4dotsAdminActions.exe";
                pr.StartInfo.Arguments = args;
                pr.StartInfo.WindowStyle = System.Diagnostics.ProcessWindowStyle.Hidden;
                pr.Start();

                System.Threading.Thread.Sleep(300);

                while (!pr.HasExited)
                {
                    Application.DoEvents();
                }

                if (pr.ExitCode != 0)
                {
                    return false;
                }
                else
                {
                    return true;
                }
            }
            catch
            {
                return false;
            }
        }

        /*
        public static bool DeleteApplicationSettingsFile()
        {
            try
            {
                string settingsFile = System.Configuration.ConfigurationManager.OpenExeConfiguration(System.Configuration.ConfigurationUserLevel.PerUserRoamingAndLocal).FilePath;

                Properties.Settings.Default.Upgrade();
                Properties.Settings.Default.Save();

                System.IO.FileInfo fi = new FileInfo(settingsFile);
                fi.Attributes = System.IO.FileAttributes.Normal;
                fi.Delete();

                return true;
            }
            catch (Exception ex)
            {
                return false;
            }
        }*/

    }

    class TranslateHelper
    {
        public static string Translate(string str)
        {
            return str;
        }
    }
}