using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using System.IO;

namespace BOLSET3IQ
{
    public class clsF29
    {

        public RetStructF29 F29(string strUserName, string strIQMode, string strBatchName,bool bReject,string strRejectReason,string strRejectReasonComments,string strProcessedXMLFile, string strRoutineNo, string strArg, DocStatastics obj)
        {
            RetStructF29 objStructF29 = new RetStructF29();
            objStructF29.Status = "F";

            try
            {
                switch (strRoutineNo)
                {
                    case "C75":

                        objStructF29.Status = "S";
                       // objStructF28 = RetrieveTiff(strUserName);

                        break;
                }
            }
            catch (Exception ex)
            {
              //  MessageBox.Show("F29 - " + ex.Message.ToString());
            }
            return objStructF29;
        }


    

    }
}
