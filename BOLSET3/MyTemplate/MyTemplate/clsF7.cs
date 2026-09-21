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


namespace BOLSET3IQ
{
    public class clsF7
    {
        //BOL_SET3.clsF7 objBOLF7;
       
        public RetStructF7 F7(string strData, string strRoutineNo, string strArgs)
        {
            //RetStructF7 objRetStructF7 = new RetStructF7();

            try
            {
                switch (strRoutineNo)
                {

                    case "CKM216": //COD
                        if (COD(strData))
                        {
                            Module1.objRetStructF7.Status = "S";
                            //objRetStructF7.Status = "S";
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid COD Amount.";
                            
                        }

                        break;

                    case "CKM254": // RADF_MABD

                        if (RADF_MABD(strData))
                        {
                            Module1.objRetStructF7.Status = "S";
                            
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Date.";
                           
                        }

                        break;

                    case "CKM255": // RADF_MABD

                        if (RADF_MABD(strData))
                        {
                            Module1.objRetStructF7.Status = "S";
                            
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Date.";
                           
                        }

                        break;


                    case "CKM220"://QUOTENO
                        if (QUOTENO(strData))
                        {
                            Module1.objRetStructF7.Status = "S";
                            
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Quote NO.";
                            
                        }
                        break;

                    case "CKM250":// Total Weight
                        if (TotalWeight(strData))
                        {
                            Module1.objRetStructF7.Status = "S";
                           
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Total Weight.";
                           
                        }
                        break;

                    case "CKM249":// Total Pieces
                        if (TotalPieces(strData))
                        {
                            Module1.objRetStructF7.Status = "S";
                           
                        }
                        else
                        {
                            Module1.objRetStructF7.Status = "F";
                            Module1.objRetStructF7.ErrorMsg = "Invalid Total Pieces.";
                            
                        }
                        break;
                }
            }
            catch (Exception ex)
            {
                Module1.objRetStructF7.Status = "F";
              
            }

            return Module1.objRetStructF7;
            //return objRetStructF7;
        }

        private bool COD(string COD)
        {
            bool bReturn = false;
            try
            {
                if (COD.Trim() == ""|| COD.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(COD, "E");
                    if ((Module1.regex.IsMatch(COD)|| COD=="0") && iQSBR008.Status == "F" )
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

        private bool RADF_MABD(string RADF_MABD)
        {
            bool bReturn = false;
            try
            {
                if (RADF_MABD.Trim() == "" || RADF_MABD.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(RADF_MABD, "E");
                    if (iQSBR008.Status == "S" || Module1.DMYwithChar.IsMatch(RADF_MABD)
                       || Module1.DMYwithoutSpace.IsMatch(RADF_MABD) || Module1.YMDwithoutSpace.IsMatch(RADF_MABD)
                       || Module1.MDYwithChar.IsMatch(RADF_MABD))
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

        private bool QUOTENO(string QUOTENO)
        {
            bool bReturn = false;
            try
            {
                if (QUOTENO.Trim() == "" || QUOTENO.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(QUOTENO, "E");
                    //if ((Module1.reg1.IsMatch(QUOTENO)|| Module1.reg2.IsMatch(QUOTENO)) && iQSBR008.Status == "F")
                    if (iQSBR008.Status == "F")
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

        private bool TotalWeight(string TotalWeight)
        {
            bool bReturn = false;
            try
            {
                if (TotalWeight.Trim() == "" || TotalWeight.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(TotalWeight, "E");
                    if (Module1.regexTW.IsMatch(TotalWeight) && iQSBR008.Status == "F" && TotalWeight.Length <=7)
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

        private bool TotalPieces(string TotalPieces)
        {
            bool bReturn = false;
            try
            {
                if (TotalPieces.Trim() == "" || TotalPieces.Trim() == "*****")
                {
                    bReturn = true;
                }
                else
                {
                    RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(TotalPieces, "E");
                    if (Module1.regexTW.IsMatch(TotalPieces) && iQSBR008.Status == "F" && TotalPieces.Length <= 5)
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
    }
}