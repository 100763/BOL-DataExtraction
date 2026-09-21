using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using System.Text.RegularExpressions;


namespace BOLSET3IQ
{
    public class clsF5
    {
       // BOL_SET3.clsF5 objBOLF5;
        public RetStructF5 F5(clsCnCWord objWord, string strRoutineNo, string strArg)
        {
            //RetStructF5 objRetStructF5 = new RetStructF5();
            Module1.objRetStructF5.Status = "F";
            
            string strData = "";
            string DD = string.Empty;
            string MM = string.Empty;
            string YY = string.Empty;

            // objBOLF5 = new BOL_SET3.clsF5();
            try
            {
                strData = objWord.strWord;
                Module1.objRetStructF5.Words = objWord;
                
                switch (strRoutineNo)
                {

                    case "CKM216"://COD
                        strData = strData.Trim();
                        if (strData.Trim() != "")
                        {
                            strData = strData.Replace("$","").Replace(".","").Replace(",", "");
                            Module1.objRetStructF5.Status = "S";
                            Module1.objRetStructF5.Words.strWord = strData;

                        }
                        else
                        {
                            Module1.objRetStructF5.Status = "F";
                        }
                        break;

                    case "CKM254"://RADF
                        strData = strData.Trim();
                        if (strData.Trim() != "")
                        {
                            strData = DateFormatDDMMYY(strData);

                            Module1.objRetStructF5.Status = "S";
                            Module1.objRetStructF5.Words.strWord = strData;
                            Module1.RADF_PDT = strData;
                           
                        }
                        else
                        {
                            Module1.objRetStructF5.Status = "F";
                        }

                        break;

                    case "CKM255"://MABD
                        strData = strData.Trim();
                        if (strData.Trim() != "")
                        {
                            strData =DateFormatDDMMYY(strData);

                            Module1.objRetStructF5.Status = "S";
                            Module1.objRetStructF5.Words.strWord = strData;
                            Module1.MABD_PDT = strData;
                        }
                        else
                        {
                            Module1.objRetStructF5.Status = "F";
                        }

                        break;


                    case "CKM220"://Quote No
                        strData = strData.Trim();
                        if (strData.Trim() != "")
                        {
                                strData = strData.ToUpper().Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Replace(":", "").Trim();
                                string Firstletter = strData.Substring(0, 1);
                                if (Firstletter == "#")
                                {
                                    strData = strData.Remove(0, 1);
                                    strData = strData.Replace(".", "");
                                    if (!objWord.Flag.Contains("~"))
                                    {
                                        ////if (!String.IsNullOrEmpty(strData) && char.IsLetter(strData[0]) && strData.Length > 7)
                                        ////{
                                        ////    strData = strData.Substring(0, 8);
                                        ////}
                                        ////else if (!String.IsNullOrEmpty(strData) && !char.IsLetter(strData[0]) && strData.Length >= 7)
                                        ////{
                                        ////    strData = strData.Substring(0, 7);
                                        ////}
                                    }

                                Module1.objRetStructF5.Status = "S";
                                Module1.objRetStructF5.Words.strWord = strData;
                            }
                                else
                                {
                                    strData = strData.Replace(".", "");
                                    if (!objWord.Flag.Contains("~"))
                                    {
                                        ////if (!String.IsNullOrEmpty(strData) && char.IsLetter(strData[0]) && strData.Length > 7)
                                        ////{
                                        ////    strData = strData.Substring(0, 8);
                                        ////}
                                        ////else if (!String.IsNullOrEmpty(strData) && !char.IsLetter(strData[0]) && strData.Length >= 7)
                                        ////{
                                        ////    strData = strData.Substring(0, 7);
                                        ////}
                                    }

                                  Module1.objRetStructF5.Status = "S";
                                  Module1.objRetStructF5.Words.strWord = strData;
                                }
                        }
                        else
                        {
                            Module1.objRetStructF5.Status = "F";
                        }
                        break;

                    case "CKM250"://Total Weight
                        strData = strData.Trim();
                        if (strData.Trim() != "")
                        {
                           
                            strData = strData.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS", "").Replace("LB", "");
                            strData = Module1.RoundOff(strData);
                            if (Module1.regexTW.IsMatch(strData))
                            {
                                Module1.objRetStructF5.Status = "S";
                                Module1.objRetStructF5.Words.strWord = strData;
                            }
                            else
                            {
                                Module1.objRetStructF5.Status = "F";
                            }
                        }
                        else
                        {
                            Module1.objRetStructF5.Status = "F";
                        }
                        break;

                    case "CKM249"://Total Pieces
                        strData = strData.Trim();
                        if (strData.Trim() != "")
                        {
                            
                            strData = strData.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS", "").Replace("LB", "");
                            strData = Module1.RoundOff(strData);
                            if (Module1.regexTW.IsMatch(strData))
                            {
                                Module1.objRetStructF5.Status = "S";
                                Module1.objRetStructF5.Words.strWord = strData;
                            }
                            else
                            {
                                Module1.objRetStructF5.Status = "F";
                            }
                        }
                        else
                        {
                            Module1.objRetStructF5.Status = "F";
                        }
                        break;

                }
            }
            catch (Exception ex)
            {
               // MessageBox.Show("F5 - " + ex.Message.ToString());
            }

            return Module1.objRetStructF5;
            
        }
        private string DateFormatDDMMYY(string strData) // Format MMDDYY
        {
            string DD = string.Empty;
            string MM = string.Empty;
            string YY = string.Empty;
            strData = strData.Replace("  "," ");
            RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(strData, "E");
            if (iQSBR008.Status == "S")
            {
                strData = IQSBR008DAteFormat(iQSBR008,strData);
            }
            else
            {
                if (Module1.YMDwithoutSpace.IsMatch(strData))
                {
                    Match Dates = Module1.YMDwithoutSpace.Match(strData);
                    string date = Dates.ToString();
                    YY = date.Substring(0, 4);
                    YY = YY.Substring(YY.Length - 2);
                    MM = date.Substring(4, 2);
                    if (MM.Length == 1)
                    {
                        MM = "0" + "" + MM;
                    }
                    DD = date.Substring(date.Length - 2);
                    if (DD.Length == 1)
                    {
                        DD = "0" + "" + DD;
                    }
                    strData = MM + "" + DD + "" + YY;
                    //Module1.MABD_PDT = strData;
                }
                //else if (Module1.DMYwithChar.IsMatch(strData))
                //{
                //    Match Dates = Module1.DMYwithChar.Match(strData);
                //    string date = Dates.ToString();
                    
                //    // strData = strData.Substring(0, date.Length);
                //    strData = date;
                //    strData = IQSBR008DAteFormat(iQSBR008, strData);
                //}
                else if (Module1.MDYwithChar.IsMatch(strData))
                {
                    Match Dates = Module1.MDYwithChar.Match(strData);
                    string date = Dates.ToString();
                    //strData = strData.Substring(0, date.Length);
                    strData = date;
                    strData = IQSBR008DAteFormat(iQSBR008, strData);
                }
                else if (Module1.MDYwithComma.IsMatch(strData))
                {
                    Match Dates = Module1.MDYwithComma.Match(strData);
                    string date = Dates.ToString();
                    //strData = strData.Substring(0, date.Length);
                    strData = date;
                    strData = strData.Replace(",", "-").Replace(" ","-");
                    strData = IQSBR008DAteFormat(iQSBR008, strData);
                }
                else if (Module1.DMMMYwithChar.IsMatch(strData))
                {
                    Match Dates = Module1.MDYwithComma.Match(strData);
                    string date = Dates.ToString();
                    //strData = strData.Substring(0, date.Length);
                    strData = date;
                    strData = strData.Replace("/", "-").Replace(",", "-").Replace(".", "-").Replace(" ", "-");
                    strData = IQSBR008DAteFormat(iQSBR008, strData);
                }
                else if (Module1.YMD.IsMatch(strData))
                {
                    Match Dates = Module1.YMD.Match(strData);
                    string date = Dates.ToString();
                    
                    strData = date;
                    strData = strData.Replace("/", "").Replace(",", "").Replace(".", "").Replace(" ", "").Replace("-", "");
                    YY = strData.Substring(0, 4);
                    YY = YY.Substring(YY.Length - 2);
                    MM = strData.Substring(4, 2);
                    if (MM.Length == 1)
                    {
                        MM = "0" + "" + MM;
                    }
                    DD = strData.Substring(strData.Length - 2);
                    if (DD.Length == 1)
                    {
                        DD = "0" + "" + DD;
                    }
                    strData = MM + "" + DD + "" + YY;
                  
                }
                else
                {
                   strData= "*****";
                }
                
            }

            return strData;
        }

        private string IQSBR008DAteFormat(RetStructIQSBR008 iQSBR008,string strData)
        {
            string DD = string.Empty;
            string MM = string.Empty;
            string YY = string.Empty;

            iQSBR008 = Module1.ObjSbr.IQSBR008(strData, "001");
            if (iQSBR008.Status == "F")
            {
                iQSBR008 = Module1.ObjSbr.IQSBR008(strData, "E");
            }
            if (iQSBR008.Status == "S")
            {
                DD = iQSBR008.Day;
                if (DD.Length == 1)
                {
                    DD = "0" + "" + DD;
                }
                MM = iQSBR008.Month;
                if (MM.Length == 1)
                {
                    MM = "0" + "" + MM;
                }
                YY = iQSBR008.Year.Substring(iQSBR008.Year.Length - 2);
                strData = MM + "" + DD + "" + YY;
            }
            else
            {
                strData = "*****";
            }

            return strData;
        }

    }
}