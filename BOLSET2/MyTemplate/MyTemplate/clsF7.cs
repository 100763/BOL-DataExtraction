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
using System.Windows.Forms;
//using BOL_SET1;

namespace BOLSET2IQ
{
    public class clsF7
    {
        //BOL_SET2.clsF7 objBOLF7;

        public RetStructF7 F7(string strData, string strRoutineNo, string strArgs)
        {
            RetStructF7 objRetStructF7 = new RetStructF7();

           // objBOLF7 = new BOL_SET2.clsF7();

            try
            {
                switch (strRoutineNo)
                {

                    case "C80": // BillOfLading

                        //if (objBOLF7.BillOfLading(strData) == true)
                        //{
                            objRetStructF7.Status = "S";
                        //}
                        //else
                        //{
                        //    objRetStructF7.Status = "F";
                        //    objRetStructF7.ErrorMsg = "Invalid Bill of Lading Number.";
                        //}

                        break;

                    case "C81": // PurchaseOrder

                        //if (objBOLF7.PurchaseOrder(strData) == true)
                        //{
                            objRetStructF7.Status = "S";
                        //}
                        //else
                        //{
                        //    objRetStructF7.Status = "F";
                        //    objRetStructF7.ErrorMsg = "Invalid Purchase Order Number";
                        //}

                        break;

                        
                    case "C82"://ShipperNumber
                        //if (objBOLF7.ShipperNumber(strData) == true)
                        //{
                            objRetStructF7.Status = "S";
                        //}
                        //else
                        //{
                        //    objRetStructF7.Status = "F";
                        //    objRetStructF7.ErrorMsg = "Invalid Shipper Number";
                        //}
                        break;
                    
                }
            }
            catch (Exception ex)
            {
                objRetStructF7.Status = "F";
                MessageBox.Show("Error in F7\\n\\n" + ex.Message.ToString());
            }

            return objRetStructF7;
        }

      
    }
}