using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
//using System.Windows.Forms;

namespace BOLSET4IQ
{
    public class clsF11
    {
        public RetStructF11 F11(string strRoutineNo, string strArg, string strTiffname)
        {
            RetStructF11 objStructF11 = new RetStructF11();
            objStructF11.Status = "F";

            try
            {
                switch (strRoutineNo)
                {
                    case "C75":
                        break;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F11 - " + ex.Message.ToString());
            }
            return objStructF11;
        }
    }
}