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
    public class TimestampLogging
    {
        private string sErrorTime;
        private string sPathName;
        private static readonly object _syncObject = new object();
        public TimestampLogging(String sPathName)
        {
            this.sPathName = sPathName;
            sErrorTime = DateTime.Now.ToString("yyyyMMdd");
        }

        public void log(string sMsg)
        {
            lock (_syncObject)
            {
                StreamWriter sw = new StreamWriter(sPathName + sErrorTime, true);
                string sLogFormat = DateTime.Now.ToShortDateString().ToString() + " " + DateTime.Now.ToLongTimeString().ToString() + " ==> ";
                sw.WriteLine(sLogFormat + sMsg);
                sw.Flush();
                sw.Close();
            }
        }
    }
}
