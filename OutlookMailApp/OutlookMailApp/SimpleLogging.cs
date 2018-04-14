// Released as open source by NCC Group Plc - https://www.nccgroup.trust/
// Developed by Soroush Dalili (@irsdl)
// Released under AGPL see LICENSE for more information
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;

namespace OutlookMailApp
{
    public class SimpleLogging
    {
        private string sPathName;
        private static readonly object _syncObject = new object();
        public SimpleLogging(String sPathName)
        {
            this.sPathName = sPathName;
        }

        public void log(string sMsg)
        {
            lock (_syncObject)
            {
                StreamWriter sw = new StreamWriter(sPathName, true);
                sw.WriteLine(sMsg);
                sw.Flush();
                sw.Close();
            }
        }
    }
}
