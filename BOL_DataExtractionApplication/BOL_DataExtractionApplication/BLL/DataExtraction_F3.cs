using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
//using BOL_SET1;
//using BOL_SET2;

namespace BOLDocumentDataExtraction.BLL
{
    //public class DataExtraction_F3
    //{
    //    public static BOL_SET1.clsF3 objBOLF3;
    //    public static BOL_SET2.clsF3 objBOL2F3;
    //    public static RetStructF3 retStructF3;

    //    //public static BOL_SET2.clsF3.USAddLine1CityStateZip objFullAddress;

    //    private static string BOLNumber = string.Empty;
    //    public static bool isAddressFound;


    //    public static RetStructF3 F3(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms, string fieldName)
    //    {
    //        retStructF3.Status = "F";
    //        retStructF3.NoOfieldsSuspects = 0;

    //        try
    //        {
    //            switch (fieldName)
    //            {
    //                case "C001":

    //                    retStructF3 = objBOLF3.BillOfLading(ObjMetaData, intCurrPageNumber,"");

    //                    if (retStructF3.Status == "S")
    //                    {
    //                        BOLNumber = retStructF3.Words[0].strWord;
    //                    }
    //                    else
    //                    {
    //                        BOLNumber = "";
    //                    }

    //                    break;

    //                case "C002":

    //                    retStructF3 = objBOLF3.PurchaseOrder (ObjMetaData, intCurrPageNumber,"");

    //                    break;

    //                case "C003":

    //                    retStructF3 = objBOLF3.ShipperNumber(ObjMetaData, intCurrPageNumber,"");

    //                    break;


    //                case "C004":

    //                    retStructF3 = objBOL2F3.ShipperName(ObjMetaData, intCurrPageNumber, BOLNumber);

    //                    break;

    //                case "C005":

    //                    retStructF3 = objBOL2F3.ShipperAddress(ObjMetaData, intCurrPageNumber, BOLNumber);

    //                    break;


    //                case "C006":

                       
    //                    retStructF3 = objBOL2F3.ShipperAddressLine1(ObjMetaData, intCurrPageNumber);

    //                    break;

    //                case "C007":

    //                    retStructF3 = objBOL2F3.ShipperCity(ObjMetaData, intCurrPageNumber);

    //                    break;

    //                case "C008":

    //                    retStructF3 = objBOL2F3.ShipperState(ObjMetaData, intCurrPageNumber);

    //                    break;

    //                case "C009":

    //                    retStructF3 = objBOL2F3.ShipperZip(ObjMetaData, intCurrPageNumber);

    //                    break;

    //                default:

    //                    break;


    //            }

    //        }
    //        catch (Exception ex)
    //        {
    //            MessageBox.Show("F3: " + ex.Message);

    //        }
    //        return retStructF3;
    //    }

    //    //private static RetStructF3 PurchaseOrder(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms)
    //    //{
    //    //    retStructF3.Status = "F";
    //    //    retStructF3.NoOfieldsSuspects = 0;

    //    //    try
    //    //    {
    //    //        retStructF3 = objBOLF3.PurchaseOrder(ObjMetaData, intCurrPageNumber);

    //    //        //if (retStructF3.Status == "S")
    //    //        //{
    //    //        //    string ReturnPurchaseOrder = string.Empty;
    //    //        //    ReturnPurchaseOrder = objBOLF5.PurchaseOrder(retStructF3.Words[0].strWord.Trim(), "");

    //    //        //    //if (objBOLF7.PurchaseOrder(ReturnPurchaseOrder) == false)
    //    //        //    //{
    //    //        //    //    ReturnPurchaseOrder = "";
    //    //        //    //}
    //    //        //}

    //    //    }
    //    //    catch (Exception ex)
    //    //    {
    //    //        MessageBox.Show("PurchaseOrder: " + ex.Message);

    //    //    }
    //    //    return retStructF3;
    //    //}

    //    //private static RetStructF3 ShipperNumber(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms)
    //    //{

    //    //    retStructF3.Status = "F";
    //    //    retStructF3.NoOfieldsSuspects = 0;

    //    //    try
    //    //    {
    //    //        retStructF3 = objBOLF3.ShipperNumber(ObjMetaData, intCurrPageNumber);

    //    //        //if (retStructF3.Status == "S")
    //    //        //{
    //    //        //    string ReturnShipperNumber = string.Empty;
    //    //        //    ReturnShipperNumber = objBOLF5.ShipperNumber(retStructF3.Words[0].strWord.Trim(), "");

    //    //        //    //if (objBOLF7.ShipperNumber(ReturnShipperNumber) == false)
    //    //        //    //{
    //    //        //    //    ReturnShipperNumber = "";
    //    //        //    //}
    //    //        //}

    //    //    }
    //    //    catch (Exception ex)
    //    //    {
    //    //        MessageBox.Show("ShipperNumber: " + ex.Message);

    //    //    }
    //    //    return retStructF3;
    //    //}
    //}
}
