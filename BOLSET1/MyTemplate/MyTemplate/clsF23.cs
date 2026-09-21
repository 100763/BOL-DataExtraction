using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;

namespace BOLSET1IQ
{
    public class clsF23
    {
        public RetStructF23 F23(string strData, string strRoutineNo, string strArg)
        {
            RetStructF23 objStructF23 = new RetStructF23();
            objStructF23.Status = "F";

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
                MessageBox.Show("F23 - " + ex.Message.ToString());
            }
            return objStructF23;
        }
    }
}