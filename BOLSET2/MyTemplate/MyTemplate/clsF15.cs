using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;

namespace BOLSET2IQ
{
    public class clsF15
    {
        public RetStructF15 F15(string strRoutineNo, string strArg, string StrFileName, string lngPageNo)
        {
            RetStructF15 objStructF15 = new RetStructF15();
            objStructF15.Status = "F";

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
                MessageBox.Show("F15 - " + ex.Message.ToString());
            }
            return objStructF15;
        }
    }
}