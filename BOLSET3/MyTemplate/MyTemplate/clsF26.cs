using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;

namespace BOLSET3IQ
{
    public class clsF26
    {
        public RetStructF26 F26(string strFilePath, string F26RtnNo, string F26RtnArgs)
        {
            RetStructF26 objStructF26 = new RetStructF26();
            objStructF26.Status = "F";

            try
            {
                switch (F26RtnNo)
                {
                    case "C75":
                        break;
                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("F26 - " + ex.Message.ToString());
            }
            return objStructF26;
        }
    }
}