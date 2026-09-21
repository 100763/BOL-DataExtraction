using iQPUBLIC;
using System;
//using System.Windows.Forms;

namespace BOLSET4IQ
{
    public class clsF10
    {
        public RetStructF10 F10(string strData, string strRoutineNo, string strArg)
        {
            RetStructF10 objStructF10 = new RetStructF10();
            objStructF10.Status = "F";

            
            try
            {
                switch (strRoutineNo)
                {
                    case "C76":

                        objStructF10.ContentsOfField= strData.Trim().ToLower();
                        objStructF10.Status = "S";

                        break;

                    case "C75":

                        clsCNCSBR objCNCSBR = new clsCNCSBR();

                        RetStructIQSBR008 obj008 = objCNCSBR.IQSBR008(strData, "E");

                        if (obj008.Status == "S")
                        {
                            objStructF10.ContentsOfField = obj008.Month + obj008.Day + obj008.Year;
                            objStructF10.Status = "S";
                        }
                        else
                        {
                            objStructF10.Status = "F";
                        }


                        break;



                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F10 - " + ex.Message.ToString());
            }
            return objStructF10;
        }
    }
}