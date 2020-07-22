using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace FrameWork.Utility
{
    public static class ProcessUtility
    {
        public static List<string> GetProcessNames()
            => Process.GetProcesses().Select(p => p.ProcessName).ToList();

        public static Process GetProcess(string name)
            => Process.GetProcesses().Where(p => p.ProcessName.Contains(name)).FirstOrDefault();

        public static void KillProcess(string name)
        {
            Process p = GetProcess(name);
            p?.Kill();
            p?.WaitForExit();
        }
    }
}
