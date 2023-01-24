using System;
using System.Collections.Generic;
using System.Text;
using System.Windows.Forms;
using System.Runtime.InteropServices;
using System.ComponentModel;
using System.Diagnostics;

namespace SelectAndCalculate
{
    public class HookingProtector
    {
        public bool timSetHookEnabled = true;

        BackgroundWorker bwTimSetHook = new BackgroundWorker();

        public void Setup()
        {            
            bwTimSetHook.DoWork += BwTimSetHook_DoWork;
            bwTimSetHook.RunWorkerAsync();
        }

        private void BwTimSetHook_DoWork(object sender, DoWorkEventArgs e)
        {
            while (true)
            {
                if (timSetHookEnabled)
                {
                    timSetHook_Tick(null, null);

                    System.Threading.Thread.Sleep(1000);
                }
            }
        }

        private bool InTimSetHook = false;

        [DllImport("user32.dll")]
        public static extern IntPtr GetWindowThreadProcessId(IntPtr hWnd, out uint ProcessId);

        [DllImport("user32.dll")]
        private static extern IntPtr GetForegroundWindow();

        private void timSetHook_Tick(object sender, EventArgs e)
        {
            //return;

            if (InTimSetHook) return;

            try
            {
                if (CPUUsageMeter.CPUUsageHigh) return;

                InTimSetHook = true;

                KeyboardHook.CheckSetHookTimer = true;

                IntPtr cw = GetForegroundWindow();

                if (cw.Equals(IntPtr.Zero)) return;

                uint procid;

                GetWindowThreadProcessId(cw, out procid);

                bool suc = System.Diagnostics.Process.GetProcessById((int)procid).WaitForInputIdle(1000);

                if (suc && System.Diagnostics.Process.GetProcessById((int)procid).Responding)
                {
                    SendKeys.SendWait("{F14}");

                    System.Threading.Thread.Sleep(3000);

                    suc = System.Diagnostics.Process.GetProcessById((int)procid).WaitForInputIdle(1000);

                    if (!(suc && System.Diagnostics.Process.GetProcessById((int)procid).Responding)) return;

                    if (KeyboardHook.CheckSetHookTimer)
                    {
                        if (CPUUsageMeter.CPUUsageHigh) return;

                        try
                        {
                            //"Hook exited. Reinstalling hooks"

                            //3frmMain.Instance.SaveSizeLocation();

                            string svisible = "";
                            string sminimize = "";

                            if (frmMain.Instance.WindowState == FormWindowState.Minimized)
                            {
                                sminimize = " /minimized";
                            }

                            if (!frmMain.Instance.Visible)
                            {
                                svisible = " /novisible";
                            }

                            /*
                            string sstarted = "";

                            if (IsStarted)
                            {
                                sstarted = " /restart";
                            }
                            */

                            //Module.WriteToLog("Starting Application");
                            //Module.WriteToLog("\"" + Application.ExecutablePath + "\" /redisable /hide");

                            System.Diagnostics.Process proc = new Process();
                            proc.StartInfo.FileName = Application.ExecutablePath;
                            proc.StartInfo.Arguments = sminimize + svisible;// + sstarted;
                            proc.StartInfo.UseShellExecute = true;
                            proc.Start();

                            while (true)
                            {
                                try
                                {
                                    var time = proc.StartTime;
                                    break;
                                }
                                catch (Exception) { }
                            }

                            System.Threading.Thread.Sleep(300);

                            //Application.Exit();
                            Environment.Exit(0);
                        }
                        finally
                        {

                        }
                    }
                }
            }
            catch (Exception exm)
            {
                Console.WriteLine(exm.ToString());
            }
            finally
            {
                InTimSetHook = false;
            }
        }

    }
}
