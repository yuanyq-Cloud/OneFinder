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
            Application.Run(new MainForm());
        }
    }
}
