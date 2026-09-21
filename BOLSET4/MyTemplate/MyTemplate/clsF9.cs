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
    public class clsF9
    {
        public RetStructF9 F9(string strRoutineNo, string strArg)
        {
            RetStructF9 objStructF9 = new RetStructF9();
            objStructF9.Status = "F";

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
                //MessageBox.Show("F9 - " + ex.Message.ToString());
            }
            return objStructF9;
        }
    }
}