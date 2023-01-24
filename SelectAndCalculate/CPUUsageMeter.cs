using System;
using System.Collections.Generic;
using System.Text;
using System.Diagnostics;

namespace SelectAndCalculate
{
    public class CPUUsageMeter
    {
        private static PerformanceCounter cpuCounter = new PerformanceCounter("Processor", "% Processor Time", "_Total");
        public static bool CPUUsageHigh
        {
            get
            {
                
                int usage = (int)cpuCounter.NextValue();
                usage = (int)cpuCounter.NextValue();

                bool isNotResponding = !System.Diagnostics.Process.GetCurrentProcess().Responding;

                bool HighUsage= (usage >= 80);

                return isNotResponding || HighUsage;
                
            }
        }
    }
}
