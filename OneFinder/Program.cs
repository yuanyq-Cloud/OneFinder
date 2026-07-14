using System;
using System.Threading;
using System.Windows.Forms;

namespace OneFinder
{
    internal static class Program
    {
        private const string MutexName = "Local\\OneFinder-SingleInstance";

        [STAThread]
        static void Main()
        {
            using var mutex = new Mutex(initiallyOwned: true, MutexName, out bool createdNew);
            if (!createdNew)
            {
                return;
            }

            ApplicationConfiguration.Initialize();

            // Initialize localization before any UI is created.
            // Reads HKCU\Software\OneFinder\Language (written by the bootstrapper).
            // Falls back to HKLM, then zh-CN.
            Loc.Initialize();

            Application.Run(new MainForm());
        }
    }
}