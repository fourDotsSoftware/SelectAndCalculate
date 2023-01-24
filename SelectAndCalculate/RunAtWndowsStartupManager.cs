using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Win32;
using System.Windows.Forms;

namespace SelectAndCalculate
{
    public class RunAtWndowsStartupManager
    {        
        public static bool RunAtWindowsStartup
        {
            get
            {
                RegistryKey key = Registry.CurrentUser;

                try
                {
                    key = key.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);

                    return (key.GetValue("SelectAndCalculate") != null);
                }
                catch
                {
                    return false;
                }                
            }
            set
            {
                bool enable = value;

                RegistryKey key = Registry.CurrentUser;

                try
                {
                    key = key.OpenSubKey(@"Software\Microsoft\Windows\CurrentVersion\Run", true);

                    if (key == null)
                    {
                        Module.ShowMessage("Error. Could not Save if Application will start automatically with Windows");
                        //return;

                        return;
                    }

                    if (enable)
                    {
                        if (key.GetValue("SelectAndCalculate") == null)
                        {
                            key.SetValue("SelectAndCalculate", "\"" + Application.StartupPath + "\\SelectAndCalculate.exe\" /hide");
                        }
                    }
                    else
                    {
                        if (key.GetValue("SelectAndCalculate") != null)
                        {
                            key.DeleteValue("SelectAndCalculate");
                        }
                    }

                }
                catch (Exception ex)
                {
                    Module.ShowMessage("Error. Could not Save if Application will start automatically with Windows");
                    //return;

                    return;
                }
            }
        }
    }
}
