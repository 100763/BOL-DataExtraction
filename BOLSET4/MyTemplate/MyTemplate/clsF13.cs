using iQPUBLIC;
using System;
//using System.Windows.Forms;

namespace BOLSET4IQ
{
    public class clsF13
    {
        public RetStructF13 F13(string strRoutineNo, string strArgs, int iX1, int iY1, int iX2, int iY2, string strTiffPath, string strInputPath, string strDBFName, string CurrFieldNumber, int intPageNo)
        {
            RetStructF13 objStructF13 = new RetStructF13();
            objStructF13.Status = "S";

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
                //MessageBox.Show("F13 - " + ex.Message.ToString());
            }
            return objStructF13;
        }
    }
}