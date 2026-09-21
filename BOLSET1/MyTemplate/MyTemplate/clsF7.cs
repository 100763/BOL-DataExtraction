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
using BOLSET1IQ;
using System.Linq;

namespace BOLSET1IQ
{
    public class clsF7
    {
       // BOL_SET1.clsF7 objBOLF7;
        //clsCNCSBR Objsbr = new clsCNCSBR();
        //Regex regex = new Regex("^([a-zA-Z]*[-,.#/_]?[0-9]+[-,.#/_]?[a-zA-Z]*[-,.#/_]?)*$");
        //Regex regex1 = new Regex("^([a-zA-Z]*[-,./#_]?[a-zA-Z]*[-,./#_]?[0-9]+[-,./#_]?)*$");
        //Regex InvalidDate = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.]*(Map|MAP|mAP|MAp)(\d{4}|(\d{3}[a-zA-Z])|(\d{1}[a-zA-Z]\d{2}))");
        //Regex Time = new Regex(@"^(?:[01]?[0-9]|2[0-3]):[0-5][0-9]$");

        public RetStructF7 F7(string strData, string strRoutineNo, string strArgs)
        {
           // RetStructF7 objRetStructF7 = new RetStructF7();

           // objBOLF7 = new BOL_SET1.clsF7();

            try
            {
                //strData = Regex.Replace(strData, @"[^0-9a-zA-Z]+", "");
                switch (strRoutineNo)
                {                   

                    case "CKM213": // BillOfLading

                        if (BillOfLading(strData) == true)
                        {
                          Module1. objRetStructF7.Status = "S";
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Bill of Lading Number.";
                        }

                        break;

                    case "CKM214": // PurchaseOrder

                        if (PurchaseOrder(strData) == true)
                        {
                            Module1.objRetStructF7.Status = "S";
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Purchase Order Number";
                        }

                        break;

                        
                    case "CKM215"://ShipperNumber
                        if (ShipperNumber(strData) == true)
                        {
                            Module1.objRetStructF7.Status = "S";
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Shipper Number";
                        }
                        break;

                    case "CKM218"://HazMat Telephone
                        if (HazMatTel(strData) == true)
                        {
                            Module1.objRetStructF7.Status = "S";
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid HazMat Telephone Number";
                        }
                        break;

                    case "CKM213_1": // Load#

                        if (LoadID(strData) == true)
                        {
                            Module1.objRetStructF7.Status = "S";
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Bill of Lading Number.";
                        }

                        break;

                }
            }
            catch (Exception ex)
            {
                Module1.objRetStructF7.Status = "F";
               // MessageBox.Show("Error in F7\\n\\n" + ex.Message.ToString());
            }

            return Module1.objRetStructF7;
        }

        private bool BillOfLading(string BillOfLading)
        {
            bool bReturn = false;
            try
            {
                if (BillOfLading.Trim() == "")
                {
                    bReturn = true;
                }
                else if (BillOfLading.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    BillOfLading = BillOfLading.Replace(" ", "");
                    if (Module1.SpecialCharacters.Contains(BillOfLading[0].ToString()))
                    {
                        BillOfLading = BillOfLading.Remove(0);
                    }
                    RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(BillOfLading, "E");
                    if (Obj008.Status == "S")
                    {
                        if (BillOfLading.All(char.IsLetterOrDigit))
                        {
                            Obj008.Status = "F";
                        }
                    }
                    if ((Module1.regexVal.IsMatch(BillOfLading) || Module1.regex1Val.IsMatch(BillOfLading)) && (!Module1.InvalidDate.IsMatch(BillOfLading)) && (!Module1.Time.IsMatch(BillOfLading)) && Obj008.Status == "F" && BillOfLading.Length >= 4)
                    {
                        bReturn = true;
                    }
                    else
                    {

                        bReturn = false;
                    }

                }

            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return bReturn;
        }
        private bool PurchaseOrder(string PurchaseOrder)
        {
            bool bReturn = false;
            try
            {
                if (PurchaseOrder.Trim() == "")
                {
                    bReturn = true;
                }
                else if (PurchaseOrder.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    string PO = PurchaseOrder.Replace(" ", "");
                    if (Module1.SpecialCharacters.Contains(PurchaseOrder[0].ToString()))
                    {
                        PurchaseOrder = PurchaseOrder.Remove(0);
                    }
                    RetStructIQSBR008 Obj008 =Module1.ObjSbr.IQSBR008(PO, "E");
                    if (Obj008.Status == "S")
                    {
                        if (PO.All(char.IsLetterOrDigit))
                        {
                            Obj008.Status = "F";
                        }
                    }
                    if ((Module1. regexVal.IsMatch(PO) ||Module1. regex1Val.IsMatch(PO)) && (!Module1. InvalidDate.IsMatch(PO)) && (!Module1.Time.IsMatch(PO)) && Obj008.Status == "F" && PO.Length >= 4)
                    {
                        bReturn = true;
                    }
                    else
                    {
                        bReturn = false;
                    }

                }

            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return bReturn;
        }
        private bool ShipperNumber(string ShipperNumber)
        {
            bool bReturn = false;
            try
            {
                if (ShipperNumber.Trim() == "")
                {
                    bReturn = true;
                }
                else if (ShipperNumber.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    //if (Module1.SpecialCharacters.Contains(ShipperNumber[0].ToString()))
                    //{
                    //    ShipperNumber = ShipperNumber.Remove(0);
                    //}
                    if(ShipperNumber.Contains("@@@"))
                    {
                        ShipperNumber = ShipperNumber.Replace("@@@", "");
                    }
                    RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(ShipperNumber, "E");
                    if (Obj008.Status == "S")
                    {
                        if (ShipperNumber.All(char.IsLetterOrDigit))
                        {
                            Obj008.Status = "F";
                        }
                    }
                    //if ((Module1.regex.IsMatch(ShipperNumber) || Module1.regex1.IsMatch(ShipperNumber)) && (!Module1.InvalidDate.IsMatch(ShipperNumber)) && (!Module1.Time.IsMatch(ShipperNumber)) && Obj008.Status == "F" && ShipperNumber.Length > 3)
                    if ((Module1.regexVal.IsMatch(ShipperNumber) || Module1.regex1Val.IsMatch(ShipperNumber)) && (!Module1.InvalidDate.IsMatch(ShipperNumber)) && (!Module1.Time.IsMatch(ShipperNumber)) && Obj008.Status == "F" && ShipperNumber.Length >= 4)
                    {
                        bReturn = true;
                    }
                    else
                    {
                        bReturn = false;
                    }
                }
            }

            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return bReturn;
        }
        private bool HazMatTel(string HazmatTel)
        {
            bool bReturn = false;
            try
            {
                if (HazmatTel.Trim() == "")
                {
                    bReturn = true;
                }
                else if (HazmatTel.Trim() == "*****")
                {
                    bReturn = true;
                    Module1.EMerTel = HazmatTel;
                }
                else
                {                   
                    if (Module1.RegTelwithMob.IsMatch(HazmatTel))
                    {
                        bReturn = true;
                        Module1.EMerTel = HazmatTel;
                    }
                    else
                    {
                        bReturn = false;
                    }
                }
            }

            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return bReturn;
        }
        private bool LoadID(string LoadID)
        {
            bool bReturn = false;
            try
            {
                if (LoadID.Trim() == "")
                {
                    bReturn = true;
                }
                else if (LoadID.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    LoadID = LoadID.Replace(" ", "");
                    if (Module1.SpecialCharacters.Contains(LoadID[0].ToString()))
                    {
                        LoadID = LoadID.Remove(0);
                    }
                    RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(LoadID, "E");
                    if (Obj008.Status == "S")
                    {
                        if (LoadID.All(char.IsLetterOrDigit))
                        {
                            Obj008.Status = "F";
                        }
                    }
                    if ((Module1.regexVal.IsMatch(LoadID) || Module1.regex1Val.IsMatch(LoadID)) && (!Module1.InvalidDate.IsMatch(LoadID)) && (!Module1.Time.IsMatch(LoadID)) && Obj008.Status == "F" && LoadID.Length >= 4)
                    {
                        bReturn = true;
                    }
                    else
                    {

                        bReturn = false;
                    }

                }

            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return bReturn;
        }
    }
}