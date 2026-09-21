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
    public class clsF6
    {
        public RetStructF6 F6(clsCnCWord objWord, string strRoutineNo, string strArg)
        {
            RetStructF6 objRetStructF6 = new RetStructF6();
            objRetStructF6.Status = "F";

            try
            {
                switch (strRoutineNo)
                {
                    case "C01":
                        objRetStructF6.Status = "S";

                        break;
                    case "C02":
                        objRetStructF6.Status = "S";

                        break;
                    case "C03":
                        objRetStructF6.Status = "S";

                        break;
                    case "C04":
                        objRetStructF6.Status = "S";

                        break;
                    case "C05":
                        objRetStructF6.Status = "S";

                        break;
                    case "C06":
                        objRetStructF6.Status = "S";

                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("F6 - " + ex.Message.ToString());
            }
            return objRetStructF6;
        }
    }
}