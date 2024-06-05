using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace EDSommaireByINDEX
{
    internal static class Program
    {
        /// <summary>
        /// Point d'entrée principal de l'application.
        /// </summary>
        [STAThread]
        static void Main(string[] args)
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Form1 form1 = new Form1();
            if (args.Length != 0)
            {
                if (args[0] == "/auto")
                {
                    
                    Application.Run(form1);
                    form1.checkBox1.Checked = true;

                }
                else if(args[0] == "/config")
                {
                    form1.checkBox1.Checked = false;
                }

            }
            
            try
            {
                Application.Run(form1);
            }
                catch (Exception ex) { }

        }
    }
}
