using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text;
using System.Text.RegularExpressions;
using iQPUBLIC;
using System.Data.SqlClient;
//using System.Windows.Forms;
//using BOL_SET4;

namespace BOLSET4IQ
{
    public class clsF7
    {

        public RetStructF7 F7(string strData, string strRoutineNo, string strArgs)
        {
            RetStructF7 objRetStructF7 = new RetStructF7();

            try
            {
                switch (strRoutineNo)
                {

                    //case "C80": // BillOfLading

                    //    if (objBOLF7.BillOfLading(strData) == true)
                    //    {
                    //        objRetStructF7.Status = "S";
                    //    }
                    //    else
                    //    {
                    //        objRetStructF7.Status = "F";
                    //        objRetStructF7.ErrorMsg = "Error Message";
                    //    }

                    //    break;

                    //case "C81": // PurchaseOrder

                    //    if (objBOLF7.PurchaseOrder(strData) == true)
                    //    {
                    //        objRetStructF7.Status = "S";
                    //    }
                    //    else
                    //    {
                    //        objRetStructF7.Status = "F";
                    //        objRetStructF7.ErrorMsg = "Error Message";
                    //    }

                    //    break;

                        
                    //case "C82"://ShipperNumber
                    //    if (objBOLF7.ShipperNumber(strData) == true)
                    //    {
                    //        objRetStructF7.Status = "S";
                    //    }
                    //    else
                    //    {
                    //        objRetStructF7.Status = "F";
                    //        objRetStructF7.ErrorMsg = "Error Message";
                    //    }
                    //    break;
                    
                }
            }
            catch (Exception ex)
            {
                objRetStructF7.Status = "F";
                //MessageBox.Show("Error in F7\\n\\n" + ex.Message.ToString());
            }
            objRetStructF7.Status = "S";
            return objRetStructF7;
        }

      
    }
}