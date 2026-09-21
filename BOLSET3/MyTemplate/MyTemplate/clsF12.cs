using iQPUBLIC;
using System;
using System.Windows.Forms;

namespace BOLSET3IQ
{
    public class clsF12
    {
        public RetStructF12 F12(string F12RtnNo, string F12RtnArgs, string strTiffname)
        {
            RetStructF12 objStructF12 = new RetStructF12();
            objStructF12.Status = "F";

            try
            {

            }
            catch (Exception ex)
            {
              //  MessageBox.Show("F12 - " + ex.Message.ToString());
            }
            return objStructF12;
        }
    }
}