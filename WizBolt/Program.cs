using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace WizBolt
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main(string[] Arguments)                      //string[] args
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            if (Arguments.Length > 0)
            {
                Application.Run(new WizBoltMainFrame(Arguments[0]));
            }
            else
            {
                Application.Run(new WizBoltMainFrame());
            }
        }
    }
}
