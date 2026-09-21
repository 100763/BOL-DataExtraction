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
    public class clsF8
    {
        public RetStructF8 F8(string strRoutineNo, string strArg)
        {
            RetStructF8 objStructF8 = new RetStructF8();
            objStructF8.Status = "F";
            objStructF8.htF8Errors = new Hashtable();
            objStructF8.ErrorMsg = new List<string>(); 

            try
            {
                switch (strRoutineNo)
                {
                    case "C75":

                        objStructF8 = Validate_F8();

                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("F8 - " + ex.Message.ToString());
            }
            return objStructF8;
        }


        private RetStructF8 Validate_F8()
        {
            RetStructF8 objStructF8 = new RetStructF8();
            objStructF8.Status = "F";
            objStructF8.htF8Errors = new Hashtable();
            objStructF8.ErrorMsg = new List<string>();
            try
            {
                DataTable dt = iQPUBLIC.Common.GetDataTable();

                string netamount = iQPUBLIC.Common.GetHeaderDataForCell(1, dt);
                string vatamount = iQPUBLIC.Common.GetHeaderDataForCell(2, dt);
                string totalamount = iQPUBLIC.Common.GetHeaderDataForCell(4, dt);

                if (Double.Parse(totalamount) == Math.Round((Double.Parse(netamount) + Double.Parse(vatamount)), 2))
                {
                    objStructF8.Status = "S";

                }
                else
                {
                    objStructF8.Status = "F";
                    string errMsg = "Total Amount = Net Amount + Vat Amout";
                    objStructF8.ErrorMsg.Add(errMsg);
                    objStructF8.htF8Errors.Add(iQPUBLIC.Common.GetHeaderKeyNameForCell(1, dt), iQPUBLIC.Common.GetHeaderKeyValueForCell(1, dt));
                    objStructF8.htF8Errors.Add(iQPUBLIC.Common.GetHeaderKeyNameForCell(2, dt), iQPUBLIC.Common.GetHeaderKeyValueForCell(2, dt));
                    objStructF8.htF8Errors.Add(iQPUBLIC.Common.GetHeaderKeyNameForCell(4, dt), iQPUBLIC.Common.GetHeaderKeyValueForCell(4, dt));
                }



            }
            catch (Exception ex)

            {
                MessageBox.Show("Validate_F8: " + ex.Message);
            
            }



            return objStructF8;
        }
    }
}