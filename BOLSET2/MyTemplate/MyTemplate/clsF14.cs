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
    public class clsF14
    {
        public RetStructF14 F14(int intPageNo, int X1, int Y1, int X2, int Y2, string strRoutineNo, string strArg, string strTiffPath, string strInputSoImage)
        {
            RetStructF14 objStructF14 = new RetStructF14();
            objStructF14.Status = "F";

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
                MessageBox.Show("F14 - " + ex.Message.ToString());
            }
            return objStructF14;
        }
    }
}