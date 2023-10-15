using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;
using SmcBasics;

namespace smcPaketversand
{
    static class Program
    {
        [STAThread]
        static void Main(string[] args)
        {
            
            string dokumentPfad = "";
            string docnr = "";
            if (args.Length > 0)
            {
                dokumentPfad = args[0];
                if (args.Length > 1)
                {
                    docnr = args[1];
                }
            }
            Application.Run(new frmPaketVersand(dokumentPfad, docnr));

        }
    }
}
