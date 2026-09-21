using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;
using BOLSET1IQ;
using System.Text.RegularExpressions;
using System.Linq;

namespace BOLSET1IQ
{
    public class clsF5
    {
        //BOL_SET1.clsF5 objBOLF5;
        public RetStructF5 F5(clsCnCWord objWord, string strRoutineNo, string strArg)
        {
            //RetStructF5 objRetStructF5 = new RetStructF5();
            Module1.objRetStructF5.Status = "F";
           // objRetStructF5.Status = "F";

            //objBOLF5 = new BOL_SET1.clsF5();

            try
            {
                switch (strRoutineNo)
                {

                    case "CKM213"://BillOfLading
                        // objWord.strWord= BillOfLading(objWord.strWord, strArg);
                        objWord.strWord = FormatValue(objWord.strWord);
                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;

                        break;


                    case "CKM214"://PurchaseOrder

                        objWord.strWord = FormatValue(objWord.strWord);
                        if (objWord.strWord.ToUpper().StartsWith("PO") || objWord.strWord.ToUpper().StartsWith("PO#") || objWord.strWord.ToUpper().StartsWith("PO-"))
                        {
                            objWord.strWord = objWord.strWord.Replace("PO#", "");
                            objWord.strWord = objWord.strWord.Replace("PO-", "");
                            objWord.strWord = objWord.strWord.Replace("PO", "");

                        }

                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;

                        break;

                    case "CKM215"://ShipperNumber

                        objWord.strWord = FormatValue(objWord.strWord);
                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;


                        break;

                    case "CKM218"://HazmatTel

                        objWord.strWord = HazmatTel(objWord.strWord, strArg);
                        Module1.objRetStructF5.Status = "S";

                        Module1.objRetStructF5.Words = objWord;

                        break;

                    case "CKM213_1":
                        objWord.strWord = FormatValue(objWord.strWord);
                        if (objWord.strWord.ToUpper().StartsWith("ID"))
                        {
                            objWord.strWord = objWord.strWord.ToUpper().Replace("ID", "");
                        }
                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;
                        break;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show("F5 - " + ex.Message.ToString());
            }

            return Module1.objRetStructF5;
        }

        private string BillOfLading(string BillOfLading, string correctionWords)
        {
            try
            {
                // BillOfLading = BillOfLading.ToLower();
                if (BillOfLading.Trim() == "")
                {
                    BillOfLading = "";
                }
                else
                {
                   // BillOfLading = BillOfLading.Replace(" ", "");
                    BillOfLading = FormatValue(BillOfLading);
                    //BillOfLading = ReplaceSpace(BillOfLading);
                }

            }
            catch (Exception ex)
            {
               // MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return BillOfLading;
        }
        private string PurchaseOrder(string PurchaseOrder, string correctionWords)
        {
            try
            {
                if (PurchaseOrder.Trim() == "")
                {
                    PurchaseOrder = "*****";
                }
                else
                {
                    if (PurchaseOrder != "*****")
                    {
                       // PurchaseOrder = PurchaseOrder.Replace(" ", "");
                        PurchaseOrder = FormatValue(PurchaseOrder);                                               

                        if (PurchaseOrder.ToUpper().StartsWith("PO") || PurchaseOrder.ToUpper().StartsWith("PO#") || PurchaseOrder.ToUpper().StartsWith("PO-"))
                        {
                            PurchaseOrder = PurchaseOrder.Replace("PO#", "");
                            PurchaseOrder = PurchaseOrder.Replace("PO-", "");
                            PurchaseOrder = PurchaseOrder.Replace("PO", "");

                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return PurchaseOrder;
        }
        private string ShipperNumber(string ShipperNumber, string correctionWords)
        {
            try
            {
                if (ShipperNumber.Trim() == "")
                {
                    ShipperNumber = "";
                }
                else
                {
                    //ShipperNumber = ShipperNumber.Replace(" ", "");
                    ShipperNumber = FormatValue(ShipperNumber);
                    //if (ShipperNumber.Contains(","))
                    //{
                    //    ShipperNumber = ReplaceComma(ShipperNumber);
                    //}
                    //else if (ShipperNumber.EndsWith("~"))
                    //{
                    //    ShipperNumber = ShipperNumber.Remove(ShipperNumber.Length - 1);
                    //}
                    //else if (ShipperNumber[0] == '#')
                    //{
                    //    ShipperNumber = ShipperNumber.Replace("#", "");

                    //}
                    //else if (ShipperNumber.Contains(","))
                    //{
                    //    ShipperNumber = ShipperNumber.Replace(',', '~');
                    //    //BillOfLading = ReplaceComma(BillOfLading);
                    //}
                    

                }
            }
            catch (Exception ex)
            {
                //Error.Log File
            }

            return ShipperNumber;
        }
        private string LoadNo(string LoadNo, string correctionWords)
        {
            try
            {
                // BillOfLading = BillOfLading.ToLower();
                if (LoadNo.Trim() == "")
                {
                    LoadNo = "";
                }
                else
                {
                    LoadNo = LoadNo.Replace(" ", "");
                    LoadNo = FormatValue(LoadNo);
                    if (LoadNo.ToUpper().StartsWith("ID"))
                    {
                        LoadNo = LoadNo.ToUpper().Replace("ID", "");
                    }
                }

            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message.ToString());
                //Error.Log File
            }

            return LoadNo;
        }

        private string FormatValue(string str)
        {
            if (str.Trim() == "")
            {
                str = "*****";
            }
            else
            {
                if (str != "*****")
                {
                    //str = str.Replace(" ", "");                   
                    str = str.Replace("'", "");
                    str = str.Replace(":", "");
                    str = str.Replace(";", "");
                    str = str.Replace("[", "");
                    str = str.Replace("]", "");
                    str = str.Replace("{", "");
                    str = str.Replace("}", "");
                    int countStartBrace = str.Count(c => (c == '('));
                    int countEndBrace= str.Count(c => (c == ')'));

                    if(countStartBrace==0 || countEndBrace==0)
                    {
                        str = str.Replace("(", "");
                        str = str.Replace(")", "");
                    }

                    //if (str.ToUpper().StartsWith("NUMBER"))
                    //{
                    //    str = str.ToUpper().Replace("NUMBER", "");
                    //}
                    //if (str.ToUpper().StartsWith("NO"))
                    //{
                    //    str = str.ToUpper().Replace("NO", "");
                    //}
                    //if(str.ToUpper().StartsWith("ID"))
                    //{
                    //    str = str.ToUpper().Replace("ID", "");
                    //}
                    if (str.ToUpper().StartsWith("NURNBER"))
                    {
                        str = str.ToUpper().Replace("NURNBER", "");
                    }
                    LabelStart: if (Module1.SpecialCharacters.Any(str.StartsWith)&&(!str.StartsWith("(")))
                    {
                        str = str.Remove(0, 1);
                        goto LabelStart;
                    }
                 LabelEnd:   if (Module1.SpecialCharacters.Any(str.EndsWith) && (!str.EndsWith(")")))
                    {
                        str = str.Remove(str.Length-1, 1);
                        goto LabelEnd;
                    }
                    if (str.Contains(","))
                    {
                        str = ReplaceComma(str);
                    }
                    if (str.EndsWith("~"))
                    {
                        str = str.Remove(str.Length - 1);
                    }
                    //if(str.Replace(" ","").All(char.IsDigit))   //replace space if all numeric
                    //{
                    //    str = str.Replace(" ", "");
                    //}
                    //if (!str.Contains("~"))
                    //{
                    //    if ((str.Count(char.IsLetter) == 1) && (!char.IsLetter(str[0])) && (!char.IsLetter(str[str.Length - 1])))
                    //    {
                    //        str = Module1.ReplaceAlphanumericWord(str);
                    //    }
                    //}                   
                   
                }

            }
            return str.ToUpper();
        }
        
        //private string ReplaceSpace(string str)
        //{
        //    string ReplacedStr = string.Empty;
        //    if (str.Contains("~"))
        //    {
        //        List<string> lstReplacedWord = new List<string>();
                
        //        String[] splitStr = str.Split('~');
        //        foreach(string BolStr in splitStr)
        //        {
        //            if(BolStr.Replace(" ", "").All(char.IsDigit))   //replace space if all numeric
        //            {                       
        //                lstReplacedWord.Add(BolStr.Replace(" ", ""));
        //            }
        //            else
        //            {
        //                lstReplacedWord.Add(BolStr);
        //            }
        //        }
        //        ReplacedStr = string.Join("~", lstReplacedWord.ToArray());
        //    }
        //    else
        //    {
        //        if (str.Replace(" ", "").All(char.IsDigit))   //replace space if all numeric
        //        {
        //            ReplacedStr= str.Replace(" ", "");
        //        }
        //        else
        //        {
        //            ReplacedStr = str;
        //        }
        //    }
        //    return ReplacedStr;
        //}

        private string HazmatTel(string HazmatTel, string correctionWords)
        {
            try
            {
                if (HazmatTel.Trim() == "")
                {
                    HazmatTel = "";
                }
                else
                {
                    if (HazmatTel != "*****")
                    {
                        HazmatTel = Module1.FuncReplace(HazmatTel);

                        if (Module1.RegTel.IsMatch(HazmatTel))
                        {
                            Match m = Module1.RegTel.Match(HazmatTel);
                            HazmatTel = m.Value.ToString();
                        }
                        if (Module1.SpecialCharacters.Any(HazmatTel.Contains))
                        {
                            HazmatTel= Regex.Replace(HazmatTel, @"[^0-9a-zA-Z]+", "");
                            //HazmatTel = Module1.RemoveSpecialCharacters(HazmatTel);
                        }
                        // int s = 0;
                        HazmatTel = HazmatTel.Replace(" ", "");
                    }

                    
                }
            }
            catch (Exception ex)
            {
                //Error.Log File
            }

            return HazmatTel;
        }

        private string ReplaceComma(string StrWord)
        {
            //string ReplaceWord = StrWord;
            char ch = ',';
            if(StrWord.EndsWith(","))
            {
                StrWord=StrWord.Remove(StrWord.Length - 1);
            }
            //int CommaCount = StrWord.Count(c => (c == ch));
            //if(CommaCount==1)
            //{
            //    string[] pattern = StrWord.Split(',');
            //    if(Module1.GetPattern(pattern[0].Trim())== Module1.GetPattern(pattern[1].Trim()))
            //    {
            //        StrWord = StrWord.Replace(',', '~');
            //    }
            //}
            //else
            //{
            //    StrWord = StrWord.Replace(',', '~');
            //}
            
            return StrWord;
        }
    }
}