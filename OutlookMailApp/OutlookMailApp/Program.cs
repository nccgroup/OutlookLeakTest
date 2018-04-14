// Released as open source by NCC Group Plc - https://www.nccgroup.trust/
// Developed by Soroush Dalili (@irsdl)
// Released under AGPL see LICENSE for more information
using System;
using System.Collections.Generic;
using System.Linq;
using System.Windows.Forms;

namespace OutlookMailApp
{
    static class Program
    {
        /// <summary>
        /// The main entry point for the application.
        /// </summary>
        [STAThread]
        static void Main()
        {
            Application.EnableVisualStyles();
            Application.SetCompatibleTextRenderingDefault(false);
            Application.Run(new FormOutlookMailApp());
        }
    }
}
