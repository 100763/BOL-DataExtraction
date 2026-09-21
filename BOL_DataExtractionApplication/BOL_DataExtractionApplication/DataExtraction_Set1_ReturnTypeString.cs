using System;

using System.Collections.Generic;
using System.Linq;
using System.Text;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using BOL_SET1;

namespace SAIABOLDocumentDataExtraction
{
    
    public class DataExtraction_Set1
    {
        public static BOL_SET1.clsF3 objBOLF3;
        public static BOL_SET1.clsF5 objBOLF5;
        public static BOL_SET1.clsF7 objBOLF7;

        //static BOL_SET1.clsF3 objBOLF3 = new clsF3();

        public static string Synonyms = string.Empty;
        public static string BillOfLading(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms)
        {
            string ReturnBillOfLading = string.Empty;
            try
            {
                RetStructF3 retStructF3 = objBOLF3.BillOfLading(ObjMetaData, intCurrPageNumber);

                if (retStructF3.Status == "S")
                {
                    ReturnBillOfLading = objBOLF5.BillOfLading(retStructF3.Words[0].strWord.Trim(), "");

                    //if (objBOLF7.BillOfLading(ReturnBillOfLading) == false)
                    //{
                    //    ReturnBillOfLading = "";
                    //}
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("BillOfLading: " + ex.Message);

            }
            return ReturnBillOfLading;
        }

        public static string PurchaseOrder(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms)
        {
            string ReturnPurchaseOrder = string.Empty;
            try
            {
                RetStructF3 retStructF3 = objBOLF3.PurchaseOrder(ObjMetaData, intCurrPageNumber);
                if (retStructF3.Status == "S")
                {
                    ReturnPurchaseOrder = objBOLF5.PurchaseOrder(retStructF3.Words[0].strWord.Trim(), "");

                    //if (objBOLF7.PurchaseOrder(ReturnPurchaseOrder) == false)
                    //{
                    //    ReturnPurchaseOrder = "";
                    //}
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("PurchaseOrder: " + ex.Message);

            }
            return ReturnPurchaseOrder;
        }

        public static string ShipperNumber(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms)
        {
            string ReturnShipperNumber = string.Empty;
            try
            {
                RetStructF3 retStructF3 = objBOLF3.ShipperNumber(ObjMetaData, intCurrPageNumber);
                if (retStructF3.Status == "S")
                {
                    ReturnShipperNumber = objBOLF5.ShipperNumber(retStructF3.Words[0].strWord.Trim(), "");

                    //if (objBOLF7.ShipperNumber(ReturnShipperNumber) == false)
                    //{
                    //    ReturnShipperNumber = "";
                    //}
                }

            }
            catch (Exception ex)
            {
                MessageBox.Show("ShipperNumber: " + ex.Message);

            }
            return ReturnShipperNumber;
        }


    }
}
