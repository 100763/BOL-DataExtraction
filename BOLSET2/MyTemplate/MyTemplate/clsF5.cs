using Microsoft.VisualBasic;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using iQDataProvider;
using iQPUBLIC;
using System.Windows.Forms;

namespace BOLSET2IQ
{
    public class clsF5
    {
        //BOL_SET2.clsF5 objBOLF5;
        public RetStructF5 F5(clsCnCWord objWord, string strRoutineNo, string strArg)
        {
            //RetStructF5 objRetStructF5 = new RetStructF5();
            Module1.objRetStructF5.Status = "F";
            //objBOLF5 = new BOL_SET2.clsF5();
            try
            {
                switch (strRoutineNo)
                {
                    case "C013": // Freight Terms, City, State, Zip
                        if (!string.IsNullOrEmpty(objWord.strWord))
                        {
                            objWord.strWord = Module1.func_RemoveSpecialCharacter(objWord.strWord);
                        }
                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;
                        break;

                    case "C900": // Address
                        if (!string.IsNullOrEmpty(objWord.strWord))
                        {
                            objWord.strWord = Module1.func_RemoveKeywordsNameAddress(objWord.strWord);
                        }
                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;
                        break;

                    case "C902": // Name

                        //if (!string.IsNullOrEmpty(objWord.strWord))
                        //{
                        //    if (objWord.strWord.Contains(":") == true)
                        //    {
                        //        string strData = objWord.strWord;
                        //        int indexcount = strData.IndexOf(":") + 1;
                        //        objWord.strWord = strData.Substring(indexcount, strData.Length - indexcount);
                        //    }
                        //    else
                        //    {
                        //        objWord.strWord = Module1.func_RemoveKeywordsNameAddress(objWord.strWord);
                        //    }
                        //}

                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;
                        break;

                    case "C901": // Telephone
                        if (!string.IsNullOrEmpty(objWord.strWord))
                        {
                            objWord.strWord = Module1.func_RemoveKeywordsTelephone(objWord.strWord);
                        }
                        Module1.objRetStructF5.Status = "S";
                        Module1.objRetStructF5.Words = objWord;
                        break;
                }
            }
            catch (Exception ex)
            {
                MessageBox.Show("F5 - " + ex.Message.ToString());
            }

            return Module1.objRetStructF5;
        }


    }
}