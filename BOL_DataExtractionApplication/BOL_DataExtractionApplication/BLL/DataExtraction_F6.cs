using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
//using BOL_SET1;
using iQPUBLIC;
using iQDataProvider;
using System.Windows.Forms;

namespace BOLDocumentDataExtraction.BLL
{
    //public class DataExtraction_F6
    //{
    //    public static BOL_SET1.clsF6 objBOLF6;
    //    public static RetStructF6 retStructF6;

    //    public RetStructF6 F6(clsCnCWord objWord, string strRoutineNo, string strArg)
    //    {
    //        retStructF6.Status = "F";
            
    //        try
    //        {
    //            switch (strRoutineNo)
    //            {
    //                case "C001"://BillOfLading

    //                    retStructF6 = objBOLF6.BillOfLading(objWord, strArg);
    //                    break;

    //                case "C002"://PurchaseOrder

    //                    retStructF6 = objBOLF6.PurchaseOrder(objWord, strArg);
    //                    break;

    //                case "C003"://ShipperNumber

    //                    retStructF6 = objBOLF6.ShipperNumber(objWord, strArg);
    //                    break;

    //                case "C04":

    //                    retStructF6.Status = "S";
    //                    break;

    //                default:

    //                    retStructF6.Status = "S";
    //                    break;

    //            }
    //        }
    //        catch (Exception ex)
    //        {
    //            MessageBox.Show("F6 - " + ex.Message.ToString());
    //        }
    //        return retStructF6;
    //    }
    //}
}
