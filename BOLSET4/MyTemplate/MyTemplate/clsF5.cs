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
    public class clsF5
    {
        
        public RetStructF5 F5(clsCnCWord objWord, string strRoutineNo, string strArg)
        {
            RetStructF5 objRetStructF5 = new RetStructF5();
            objRetStructF5.Status = "F";


            


            try
            {
                switch (strRoutineNo)
                {


                    //case "C80"://BillOfLading


                    //    objWord.strWord= objBOLF5.BillOfLading(objWord.strWord, strArg);

                    //    objRetStructF5.Status = "S";
                    //    objRetStructF5.Words = objWord;

                    //    break;


                    //case "C81"://PurchaseOrder

                    //    objWord.strWord = objBOLF5.PurchaseOrder(objWord.strWord, strArg);

                    //    objRetStructF5.Status = "S";
                    //    objRetStructF5.Words = objWord;

                    //    break;

                    //case "C82"://ShipperNumber

                    //    objWord.strWord = objBOLF5.ShipperNumber(objWord.strWord, strArg);
                    //    objRetStructF5.Status = "S";

                    //    objRetStructF5.Words = objWord;

                    //    break;

                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F5 - " + ex.Message.ToString());
            }

            return objRetStructF5;
        }
    }
}