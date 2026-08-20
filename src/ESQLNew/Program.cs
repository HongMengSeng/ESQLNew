using System;
using System.Windows.Forms;
using ESQLNew.Core;

namespace ESQLNew
{
    internal static class Program
    {
        [STAThread]
        private static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            AppPaths.MigrateFromLegacy();
            Application.Run(new MainForm());
        }
    }
}