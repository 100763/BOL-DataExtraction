using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Linq;
using System.Text.RegularExpressions;
using iQDataProvider;
using iQPUBLIC;


namespace BOLSET3IQ
{
    static class Module1
    {


        //public static int AmtNum = 0;


        public static RetStructF3 objRetStructF3 = new RetStructF3();
        public static RetStructF5 objRetStructF5 = new RetStructF5();
        public static RetStructF7 objRetStructF7 = new RetStructF7();
        public static RetStructF27 objStructF27 = new RetStructF27();

        
        public static clsCnCWord MABDDate= new clsCnCWord();
        public static clsCnCBoundingWords ObjBoundingWord = new clsCnCBoundingWords();
        public static clsCNCSBR ObjSbr = new clsCNCSBR();
        public static string RADF_PDT = string.Empty;
        public static string MABD_PDT = string.Empty;
        public static bool Dlv_O_flag = false;
        public static bool isFuturedate = false;
        public static bool FuturePDTRADF = false;
        public static bool FuturePDTMABD = false;

        public static int MX1, MX2, MY1, MY2;

        public static CultureInfo provider = new CultureInfo("en-US");
        public static DateTime ToadysDate = DateTime.Now;
        public static DateTime ImageDate;

        //public static DataTable dtSpecial_Instruct = new DataTable();
        public static DataTable dtAccessorials = new DataTable();
        public static DataTable ComnInstrConatainer = new DataTable();
        public static DataTable Dt_distinctInstruction = new DataTable();
        public static bool pgSearchInstrut = false;
        public static DataTable Dt_IgnoreShipperDetails = new DataTable();
        public static bool IgnoreshipperExist= false;

        public static DataTable Dt_KeywordMerge = new DataTable();

        //public static DataTable dtPickupConsigneDetails = new DataTable();
        //public static DataTable dtSpecial_Instruct_Accessorial = new DataTable();
        //  public static DataTable dtPickup_Consigne_Details = new DataTable();
        public static List<string> SpecialCharacters = new List<string>() { "!", "@", "#", "$", "%", "^", "&", "*", "(", ")", "_", "-", "+", "=", "{", "}", "[", "]", ":", ";", ",", "<", ">", ".", "?", "/", "~", "`" };
        public static List<string> MarkRedSpecialCharacters = new List<string>() { ".", ",", "/", ":", ";", "(", ")" , "\"","!" };
        public static List<string> MarkRedDateSpecialCharacters = new List<string>() { ":", ";", "(", ")", "\"", "!" };
        #region  Regex for All fields

        public static Regex regonlynumeric = new Regex(@"^[0-9*#+/.@,()!~%&-]+$");
        public static Regex onlyAlphawithSpecialChar = new Regex(@"^[a-zA-Z*#+/.@,()!~%&-:;/^]+$");
        //COD regex
        public static Regex regex = new Regex(@"^[\s]*(rs|inr|rs|US\.|£|€|US\$|\$)?[\s]*((([0-9]+)([\.|\,]([0-9]{1,2})))|((([0-9]{2}[\,|\.|'])+[0-9]{3})([\,|\.]([0-9]{1,2}))?)|(((([0-9]{1,2})[\,|\.|'])+([0-9]{3}))([\.|\,]([0-9]{1,2}))?)|((([0-9]{1,2})[\,|\.])+([0-9]{3}))|([0-9]{3}))[\s]*(\£|\€|euro|EUR|Eur|eur)?[\s]*$");

        //Quate No Regex
        public static Regex reg1 = new Regex(@"^(E|L|T|S|#)(\d{7})(?!\d)");
        public static Regex reg2 = new Regex(@"^([1-9](\d{6})(?!\d))");
        //public static Regex regQuote2 = new Regex(@"^([A-Z]|#)(\d{7})(?!\d)");

        // Date Regex
        //public static Regex DMYwithoutSpace = new Regex(@"^(3[01]|[12][0-9]|0?[1-9]|[1-9])((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(1[012]|0?[1-9]))(\d{4})$");
        //public static Regex YMDwithoutSpace = new Regex(@"^(\d{4})((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(1[012]|0?[1-9]))(3[01]|[12][0-9]|0?[1-9]|[1-9])$");
        //public static Regex DMYwithChar = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.]((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(1[012]|0?[1-9]))[,-/.](\d{4})(?:/[A-Za-z])?");
        //public static Regex MDYwithChar = new Regex(@"((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(1[012]|0?[1-9]))[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])[,-/.](\d{4})(?:/[A-Za-z])?");

        //(((19|20)[0-9]{2})|(\d{2}))

        public static Regex DMYwithoutSpace = new Regex(@"^(3[01]|[12][0-9]|0?[1-9]|[1-9])(1[012]|0?[1-9])(\d{4})$");
        //public static Regex YMDwithoutSpace = new Regex(@"^(\d{4})(1[012]|0?[1-9])(3[01]|[12][0-9]|0?[1-9]|[1-9])$");
        public static Regex DMYwithChar = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.](1[012]|0?[1-9])[,-/.](\d{4})(?:/[A-Za-z])?");
        public static Regex MDYwithChar = new Regex(@"(1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])[,-/.](\d{4})(?:/[A-Za-z])?");
        public static Regex YMD = new Regex(@"^(\d{4})[,-/.](1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])$");

        //public static Regex DMYwithoutSpace = new Regex(@"^(3[01]|[12][0-9]|0?[1-9]|[1-9])(1[012]|0?[1-9])(((19|20)[0-9]{2})|(\d{2}))$");
        public static Regex YMDwithoutSpace = new Regex(@"^(((19|20)[0-9]{2})|(\d{2}))(1[012]|0?[1-9])(3[01]|[12][0-9]|0?[1-9]|[1-9])$");
        //public static Regex DMYwithChar = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.](1[012]|0?[1-9])[,-/.](((19|20)[0-9]{2})|(\d{2}))(?:/[A-Za-z])?");
        //public static Regex MDYwithChar = new Regex(@"(1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])[,-/.](((19|20)[0-9]{2})|(\d{2}))(?:/[A-Za-z])?");
        // public static Regex YMD = new Regex(@"^(((19|20)[0-9]{2})|(\d{2}))[,-/.](1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])$");


        public static Regex DMMMYwithoutSpace = new Regex(@"^(3[01]|[12][0-9]|0?[1-9]|[1-9])(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(((19|20)[0-9]{2})|(\d{2}))$");
        public static Regex YMMMDwithoutSpace = new Regex(@"^(((19|20)[0-9]{2})|(\d{2}))(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)(3[01]|[12][0-9]|0?[1-9]|[1-9])$");
        public static Regex DMMMYwithChar = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.](January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[,-/.](((19|20)[0-9]{2})|(\d{2}))(?:/[A-Za-z])?");
        public static Regex MMMDYwithChar = new Regex(@"(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])[,-/.](((19|20)[0-9]{2})|(\d{2}))(?:/[A-Za-z])?");

        public static Regex MDYwithComma = new Regex(@"^((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s(3[01]|[12][0-9]|0?[1-9]|[1-9]),\s((?:[0-9]{2})?[0-9]{2}))");

        // Total Weight Regex
        public static Regex regexTW = new Regex(@"^[\s]*(rs|inr|rs|US\.|£|€|US\$|\$)?[\s]*([+-]?[0-9]{1,3}(?:[0-9]*(?:[.,][0-9]{2})?|(?:,[0-9]{3})*(?:\.[0-9]{0,3})?|(?:\.[0-9]{3})*(?:,[0-9]{2})?))[\s]*(\£|\€|euro|EUR|Eur|eur|lbs|LBS)?[\s]*$");

        // Telephone Regex
        public static Regex RegTel = new Regex(@"(\d{3}\s+\d{3}\s+\d{4}|\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\-\s*\d{3}\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{1}\s*\.\s*\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\.\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\d{3}\s*\-\s*\d{4}|\(\s*\d{3}\s*\)\s*\-\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\.\s*\d{4}|\d{3}\s*\.\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*/\s*\d{3}\s*\-\s*\d{4}|\d{3}\s*\-\s*\d{3}\s*/\s*\d{4}|\d{3}\s*/\s*\d{3}\s*/\s*\d{4})");

        #endregion
        
        // for only numeric or numeric with special characters

        public static bool bFoundSearchKeywordTable = false;
        public static string imageName = "";

        public static clsCnCWord func_GetDefaultWord(int PageNumber)
        {
            clsCnCWord oWord = new clsCnCWord();
            var _with1 = oWord;
            _with1.X1Char = "5";
            _with1.Y1Char = "5";
            _with1.X2Char = "5";
            _with1.Y2Char = "5";
            _with1.Confidence = 90;
            //.intLineNumber = 0
            _with1.LineNo = 1;
            _with1.PageNo = PageNumber;
            _with1.Left = 250;
            _with1.Right = 250;
            _with1.Top = 400;
            _with1.Bottom = 400;
            _with1.strWord = "";
            _with1.ConfString = "";

            return oWord;

        }

        #region Old function
        //public static DataSet func_GetKeywords_old(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        //{
        //    DataSet ds = new DataSet();
        //    //bool bFoundSearchKeywordTable = false;
        //    if (bFoundSearchKeywordTable == false && iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
        //    {
        //        iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET3_keyword");
        //    }
        //    if (!bFoundSearchKeywordTable)
        //    {
        //        //// iQPUBLIC.PublicComponents.htMyVariable = new Hashtable(); commented 16/11/2020
        //        //iQPUBLIC.PublicComponents.htMyVariable = new Hashtable();
        //        DataTable dtkeyword = F3_SearchKeyword(ObjMetaData, intCurrPageNumber);
        //        //DataSet ds = new DataSet();
        //        ds.Tables.Add(dtkeyword);
        //        iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_keyword", ds);
        //        bFoundSearchKeywordTable = true;
        //    }
        //    return ds;
        //}

        //public static DataTable F3_SearchKeyword_old(clsCncMetaData oMeta, int intCurrPageNo)
        //{
        //    var dtKeywordInfo = new DataTable();
        //    dtKeywordInfo = MakeDt();
        //    try
        //    {
        //        // -----------------------------------------INVOICE NUMBER/INVOICE DATE/PURCHASE ORDER NUMBER/INOICE AMOUNT KEYWORD------------------------------------------------------------
        //        var sKeyarray = new string[]
        //            {
        //            "SPECIAL INSTRUCTION", "Comments/Special Instruction", "Shipping Instructions","Consignee Special Instructions","Shipper Special Instructions", "Special Services and Instructions",
        //            "Carrier Special delivery Instructions","Origin Instructions","SpecialInstructions","Special delivery Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction",
        //            "COD AMOUNT", "COD AMT", "COD",
        //            "Due Date","Delivery Date","Estimated Delivery Date","Guaranteed Before","Deliver On","Delivery Requested Date","Requested Delivery Date","Delivery Appointment",
        //            "Deliv From","Delivery Date/Time","Deliver Before Date","To Arrive","Deliver No Earlier Than","Deliver No Later Than","Req Delivery date","Schedule Delivery date","Sch Del Date",
        //            "MABD","MABD Date","Must Arrive By Date","Must Deliver by date","Must Deliver By","Deliver By","ARRIVE BY",
        //            "Quote", "Quote No", "Quote Number","QuoteID","Quote Id","Quote ID Number","RQ","QuoteNumber",
        //            "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","GRANDTOTAL","NetWeight","GrossWeight","TotalAmount","Weight Total","TOTAL WT","TotalWeight","PAGE SUBTOTAL","Weight",
        //            "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
        //            "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
        //            "QTY","Quantity","Pallets", "Total","NO. OF PKGS","No. Packages","Total Skids","Total Pieces","Colis / Package","Package Count","H. Units","HU"};
        //        // -----------------------------------------------------------------------------------------------------------------------------------

        //        Array.Sort(sKeyarray);
        //        var CNC = new ClsCNC();
        //        var conflevel = new List<int>();
        //        conflevel.Clear();
        //        int iPage;
        //        //var loopTo = oMeta.PageCount;
        //        var loopTo = intCurrPageNo;
        //        for (iPage = 1; iPage <= loopTo; iPage++)
        //        {
        //            int h = iQPUBLIC.PublicComponents.ImageHeight;
        //            int w = iQPUBLIC.PublicComponents.ImageWidth;
        //            //clsCnCWord[] Words = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999);
        //            clsCnCLine[] CncLines = CNC.GetLinesFromROI(oMeta.Page, iPage, 0, 0, 9999, 9999);
        //            var wordsArray = new clsCnCWord[0];
        //            var oClsCnc = new ClsCNC();
        //            var oWord = new clsCnCWord();
        //            if (CncLines is object)
        //            {
        //                if (CncLines.Length != 0)
        //                {
        //                    for (int line = 0; line <= CncLines.Length - 1; line++)
        //                    {
        //                        for (int i = 1; i <= CncLines[line].WordCount; i++)
        //                        {
        //                            try
        //                            {
        //                                for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
        //                                {
        //                                    var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
        //                                    if (skeysplit.Length > 1)
        //                                    {
        //                                        if (CheckValidWord(CncLines[line].Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), new string[] { skeysplit[0] }) == true)
        //                                        {
        //                                            // If Words(i).strWord.ToString.ToUpper.Trim.Replace("'", "`").Contains(skeysplit(0).ToUpper.Trim) Then

        //                                            if (i >= CncLines[line].WordCount)
        //                                            {
        //                                            }
        //                                            else
        //                                            {
        //                                                int iIndex;
        //                                                var loopTo3 = skeysplit.Length - 1;
        //                                                for (iIndex = 0; iIndex <= loopTo3; iIndex++)
        //                                                {
        //                                                    if (i + iIndex > CncLines[line].WordCount)
        //                                                        continue;
        //                                                    if (CheckValidWord(CncLines[line].Word[i + iIndex].strWord.ToString().ToUpper().Trim().Replace("'", "`"), new string[] { skeysplit[iIndex] }) == true)
        //                                                    {

        //                                                        // If Words(i + iIndex).strWord.ToString.ToUpper.Trim.Contains(skeysplit(iIndex).Trim.ToUpper()) Then

        //                                                        if (skeysplit.Length - 1 != iIndex)
        //                                                        {
        //                                                            continue;
        //                                                        }
        //                                                        else
        //                                                        {
        //                                                            DataRow drDataRow = dtKeywordInfo.NewRow();
        //                                                            drDataRow["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                                            drDataRow["Page No"] = iPage;
        //                                                            drDataRow["Line No"] = CncLines[line].Word[i].LineNo;
        //                                                            //drDataRow["Word No"] = i + iIndex;
        //                                                            drDataRow["Word No"] = CncLines[line].Word[i].WordNumber;
        //                                                            // drDataRow["LineWordNo"] = CncLines[line].Word[i].WordNumber;
        //                                                            // drDataRow("X1") = Words(i).Top
        //                                                            // drDataRow("Y1") = Words(i).Left
        //                                                            // drDataRow("X2") = Words(i + iIndex).Bottom
        //                                                            // drDataRow("Y2") = Words(i + iIndex).Right
        //                                                            // drDataRow("Word No") = Words(i).WordNumber + iIndex
        //                                                            drDataRow["X1"] = CncLines[line].Word[i].Left;
        //                                                            drDataRow["Y1"] = CncLines[line].Word[i].Top;
        //                                                            drDataRow["X2"] = CncLines[line].Word[i + iIndex].Right;
        //                                                            drDataRow["Y2"] = CncLines[line].Word[i + iIndex].Bottom;
        //                                                            dtKeywordInfo.Rows.Add(drDataRow);
        //                                                        }
        //                                                    }
        //                                                    else
        //                                                    {
        //                                                        break;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }
        //                                    }
        //                                    //else if (Words[i].strWord.ToString().ToUpper().Trim().Replace("'", "`").Contains(sKeyarray[jLoop].Trim().ToUpper()))
        //                                    else if (CheckValidWord(CncLines[line].Word[i].strWord.ToString().ToUpper().Trim().Replace("'", "`"), new string[] { skeysplit[0] }) == true)
        //                                    {
        //                                        DataRow drDataRow;
        //                                        drDataRow = dtKeywordInfo.NewRow();
        //                                        drDataRow["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                        drDataRow["Page No"] = iPage;
        //                                        drDataRow["Line No"] = CncLines[line].Word[i].LineNo;
        //                                        // drDataRow["Word No"] = i;
        //                                        drDataRow["Word No"] = CncLines[line].Word[i].WordNumber;
        //                                        //drDataRow["LineWordNo"] = CncLines[line].Word[i].WordNumber;
        //                                        // drDataRow("X1") = Words(i).Top
        //                                        // drDataRow("Y1") = Words(i).Left
        //                                        // drDataRow("X2") = Words(i).Bottom
        //                                        // drDataRow("Y2") = Words(i).Right
        //                                        // drDataRow("Word No") = Words(i).WordNumber
        //                                        drDataRow["X1"] = CncLines[line].Word[i].Left;
        //                                        drDataRow["Y1"] = CncLines[line].Word[i].Top;
        //                                        drDataRow["X2"] = CncLines[line].Word[i].Right;
        //                                        drDataRow["Y2"] = CncLines[line].Word[i].Bottom;
        //                                        dtKeywordInfo.Rows.Add(drDataRow);
        //                                    }
        //                                }
        //                            }
        //                            catch (Exception ex)
        //                            { }
        //                        }
        //                    }
        //                }
        //            }
        //        }

        //    }
        //    catch (Exception ex)
        //    {
        //        //MsgBox(ex.Message, MsgBoxStyle.Critical, "SearchKeywords");
        //    }

        //    return dtKeywordInfo;
        //}

        //private static bool CheckValidWord_old(string sWord, string[] arrcheckkeywords)
        //{
        //    bool bFlag = false;
        //    int lengthofword;
        //    // Dim AllMatches As MatchCollection = Regex.Matches(sWord, "[a-zA-Z0-9]+", RegexOptions.IgnoreCase)
        //    // For Each SingleMatch As Match In AllMatches
        //    // If Array.IndexOf(arrAllWords, SingleMatch.Value.ToUpper()) >= 0 Then
        //    // bFlag = True
        //    // End If
        //    // Next
        //    var objRestruct056 = new iQPUBLIC.clsCNCSBR();
        //    iQPUBLIC.RetStructIQSBR058 objRestruct058;
        //    if (!string.IsNullOrEmpty(sWord))
        //    {

        //        // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
        //        var arrFinalkeyword = new List<string>();
        //        sWord = sWord.Replace("'", "");
        //        sWord = sWord.Replace("`", "");
        //        if (SpecialCharacters.Any(sWord.Contains))
        //        {
        //            //sWord = RemoveSpecialCharacters(sWord);
        //            sWord = Regex.Replace(sWord, @"[^0-9a-zA-Z]+", "");
        //        }

        //        lengthofword = sWord.Length;

        //        if (lengthofword <= 3)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length <= 3)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword > 3 & lengthofword <= 5)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length > 2 && arrcheckkeywords[iLoop].Length <= 6)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }

        //        }
        //        else if (lengthofword >= 6 & lengthofword <= 8)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 4 && arrcheckkeywords[iLoop].Length <= 10)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 9 & lengthofword <= 11)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 6 && arrcheckkeywords[iLoop].Length <= 14)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 12)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("'", "");
        //                arrcheckkeywords[iLoop] = arrcheckkeywords[iLoop].Replace("`", "");
        //                if (arrcheckkeywords[iLoop].Length >= 8)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }

        //        for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)
        //        {
        //            objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
        //            if (lengthofword <= 3)
        //            {
        //                if (objRestruct058.CharChanged < 1)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            if (lengthofword > 3 && lengthofword <= 5)
        //            {
        //                if (objRestruct058.CharChanged <= 1)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 6 & lengthofword <= 8)
        //            {
        //                if (objRestruct058.CharChanged <= 2)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 9 & lengthofword <= 11)
        //            {
        //                if (objRestruct058.CharChanged <= 3)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 12)
        //            {
        //                if (objRestruct058.CharChanged <= 4)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //        }

        //    }

        //    return bFlag;
        //}
        #endregion


        #region  Keyword Searching Logic
        // new with dictionary
        public static DataSet func_GetKeywords(clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            DataSet ds = new DataSet();
            //bool bFoundSearchKeywordTable = false;
            if (bFoundSearchKeywordTable == false)
            {
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    iQPUBLIC.PublicComponents.htMyVariable.Remove("BOLSET3_keyword");
                }

                DataTable dtkeyword = F3_SearchKeyword(ObjMetaData, intCurrPageNumber);
                //DataSet ds = new DataSet();
                ds.Tables.Add(dtkeyword);
                iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_keyword", ds);
                bFoundSearchKeywordTable = true;
            }

            //if (!bFoundSearchKeywordTable)
            //{
            //    DataTable dtkeyword = F3_SearchKeyword(ObjMetaData, intCurrPageNumber);
            //    //DataSet ds = new DataSet();
            //    ds.Tables.Add(dtkeyword);
            //    iQPUBLIC.PublicComponents.htMyVariable.Add("BOLSET3_keyword", ds);
            //    bFoundSearchKeywordTable = true;
            //}
            return ds;
        }
        public static DataTable F3_SearchKeyword_old14062021(clsCncMetaData oMeta, int intCurrPageNo)
        {
            var dtKeywordInfo = new DataTable();
            dtKeywordInfo = MakeDt();
            try
            {
                // -----------------------------------------INVOICE NUMBER/INVOICE DATE/PURCHASE ORDER NUMBER/INOICE AMOUNT KEYWORD------------------------------------------------------------
                #region Original Keywords
                ////Original Keywords
                var sKeyarray = new string[]
                    {
                    "SPECIAL INSTRUCTION", "Comments/Special Instruction", "Shipping Instructions","Consignee Special Instructions","Shipper Special Instructions", "Special Services and Instructions",
                    "Carrier Special delivery Instructions","Origin Instructions","SpecialInstructions","Special delivery Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction",
                    "COD AMOUNT", "COD AMT", "COD",
                    "Due Date","Delivery Date","Estimated Delivery Date","Guaranteed Before","Deliver On","Delivery Requested Date","Requested Delivery Date","Delivery Appointment",
                    "Deliv From","Delivery Date/Time","Deliver Before Date","To Arrive","Deliver No Earlier Than","Deliver No Later Than","Req Delivery date","Schedule Delivery date","Sch Del Date",
                    "Estimated Delivery","Requested Date/Time","REQ ARV DTE","DLV Date","REQ DATE","ARRIVE DATE","Est. Arrival Date","Delv Date","Del Date","Delivery Dt","Delivery","Appt Date/Time",
                    "Delivery Window","DE L DATE","Appt. Date / Time","Est. Arrival","Request Date","Must deliver only on","Dropoff Date/Time","Do Not Del Before","Do Not Del After","Do Not Deliver Before","Do Not Deliver After",
                    "MABD","MABD Date","Must Arrive By Date","Must Deliver by date","Must Deliver By","Deliver By","ARRIVE BY","MUST DELIVER ON OR BEFORE","MUST DELIVER BETWEEN","Must Arrive By Window","DELIVERY MUST BE MADE BY",
                    "Quote", "Quote No", "Quote Number","QuoteID","Quote Id","Quote ID Number","Quote Reference ID","QuoteNumber","Quot e","Q#",
                    "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","GRANDTOTAL","NetWeight","GrossWeight","TotalAmount","Weight Total","TOTAL WT","TotalWeight","PAGE SUBTOTAL","Weight","Total","Total Gross","Gross",
                    "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                    "Pallet Count","Total Skids","Colis/Package","Package Count","H. Units","HU","No. of Pallets", "Number of Pallets","TOTAL PALLETS",
                    "TOTAL PALLE TS","Handling Units Total","Total Handling Units","H/Us","Handling Units","TOTAL PLTS" //,"TOTAL PKGS","NO. OF PKGS","No. Packages","Total Pieces", "Total","TOTAL PCS","QTY","Pallets"
                    };
                ////Original Keywords
                #endregion


                #region Minimize Keywords
                //Mimnimized Keywords
                ////var sKeyarray = new string[]
                ////    {
                ////    "SPECIAL INSTRUCTION", "Comments/Special Instruction", "Shipping Instructions","Consignee Special Instructions","Shipper Special Instructions",//// "Special Services and Instructions",
                ////    ////"Carrier Special delivery Instructions","Origin Instructions","SpecialInstructions","Special delivery Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction",
                ////    "COD AMOUNT", "COD AMT", "COD",
                ////    "Due Date","Delivery Date","Estimated Delivery Date","Guaranteed Before","Deliver On","Delivery Requested Date","Requested Delivery Date","Delivery Appointment",
                ////    "Deliv From","Delivery Date/Time","Deliver Before Date","To Arrive","Deliver No Earlier Than","Deliver No Later Than","Req Delivery date","Schedule Delivery date","Sch Del Date",
                ////    "Estimated Delivery",////"Requested Date/Time","REQ ARV DTE","DLV Date","REQ DATE","ARRIVE DATE","Est. Arrival Date","Delv Date","Del Date","Delivery Dt","Delivery","Appt Date/Time",
                ////    ////"Delivery Window","DE L DATE","Appt. Date / Time","Est. Arrival","Request Date","Must deliver only on","Dropoff Date/Time","Do Not Del Before","Do Not Del After","Do Not Deliver Before","Do Not Deliver After",
                ////    "MABD","MABD Date","Must Arrive By Date","Must Deliver by date","Must Deliver By","Deliver By","ARRIVE BY","MUST DELIVER ON OR BEFORE",////"MUST DELIVER BETWEEN","Must Arrive By Window","DELIVERY MUST BE MADE BY",
                ////    "Quote", "Quote No", "Quote Number","QuoteID","Quote Id","Quote ID Number",////"Quote Reference ID","QuoteNumber","Quot e","Q#",
                ////    "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","GRANDTOTAL",////"NetWeight","GrossWeight","TotalAmount","Weight Total","TOTAL WT","TotalWeight","PAGE SUBTOTAL","Weight","Total","Total Gross","Gross",
                ////    "DESTINATION INSTRUCTIONS","Delivery Instructions",////"Delivery Notes","Delivery Remarks",
                ////    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                ////    "Pallet Count","Package Count","H. Units","HU","No. of Pallets", "Number of Pallets","TOTAL PALLETS",////"Total Skids","Colis/Package"
                ////   //// "Handling Units Total","Total Handling Units","H/Us","Handling Units","TOTAL PLTS" ,"TOTAL PALLE TS"
                ////    };
                // -----------------------------------------------------------------------------------------------------------------------------------
                //Mimnimized Keywords
                #endregion

                #region  Sorted Keywords
                //////Sorted Keywords
                //var sKeyarray = new string[]
                //{ "Special Services and Instructions","Carrier Special delivery Instructions","Comments/Special Instruction","Consignee Special Instructions","Shipper Special Instructions",
                //    "Special delivery Instructions", "Shipping Instructions","Origin Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction","SPECIAL INSTRUCTION","SpecialInstructions",
                //    "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                //    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                //    "COD AMOUNT","COD AMT","COD",
                //    "Quote ID Number","Quote Reference ID","Quote No","Quote Number","Quote Id","QuoteID","QuoteNumber","Quote",
                //    "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","Weight Total","TOTAL WT","PAGE SUBTOTAL","Total Gross","GRANDTOTAL","NetWeight","GrossWeight",
                //    "TotalAmount","TotalWeight","Weight","Total","Gross",
                //    "Number of Pallets","No. of Pallets","TOTAL PALLE TS","Handling Units Total","Total Handling Units","Pallet Count","Total Skids","Package Count","H. Units","TOTAL PALLETS",
                //    "Handling Units","TOTAL PLTS","H/Us","HU",
                //    "Deliver No Earlier Than","Deliver No Later Than","Appt. Date / Time","Must deliver only on","Do Not Del Before","Do Not Del After","Do Not Deliver Before","Do Not Deliver After",
                //    "Estimated Delivery Date","Delivery Requested Date","Requested Delivery Date","Deliver Before Date","Req Delivery date","Schedule Delivery date","Sch Del Date","REQ ARV DTE","Est. Arrival Date",
                //    "Delivery Date","Guaranteed Before","Deliver On","Delivery Appointment","Deliv From","Delivery Date/ Time","Due Date","To Arrive","Estimated Delivery","Requested Date/ Time","DLV Date",
                //    "REQ DATE","ARRIVE DATE","Delv Date","Del Date","Delivery Dt","Appt Date/Time","Delivery Window","Est. Arrival","Request Date","Dropoff Date/Time",
                //    "DELIVERY MUST BE MADE BY","MUST DELIVER ON OR BEFORE","Must Arrive By Date","Must Arrive By Window","Must Deliver by date","MUST DELIVER BETWEEN",
                //    "Must Deliver By","ARRIVE BY","Deliver By","MABD Date","Delivery","MABD"
                //};
                ////Sorted Keywords
                #region Sorted backup
                ///Sorted backup
                //var sKeyarray = new string[]
                //{ "Special Services and Instructions","Carrier Special delivery Instructions","Comments/Special Instruction","Consignee Special Instructions","Shipper Special Instructions",
                //    "Special delivery Instructions", "Shipping Instructions","Origin Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction","SPECIAL INSTRUCTION","SpecialInstructions",
                //    "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                //    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                //    "COD AMOUNT","COD AMT","COD",
                //    "Quote ID Number","Quote Reference ID","Quote No","Quote Number","Quote Id","QuoteID","QuoteNumber","Quote",
                //    //"Quot e","Q#",
                //    "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","Weight Total","TOTAL WT","PAGE SUBTOTAL","Total Gross","GRANDTOTAL","NetWeight","GrossWeight",
                //    "TotalAmount","TotalWeight","Weight","Total","Gross",
                //    "Number of Pallets","No. of Pallets","TOTAL PALLE TS","Handling Units Total","Total Handling Units","Pallet Count","Total Skids","Package Count","H. Units","TOTAL PALLETS",
                //    "Handling Units","TOTAL PLTS","H/Us","HU",//"Colis/Package",
                //    "Deliver No Earlier Than","Deliver No Later Than","Appt. Date / Time","Must deliver only on","Do Not Del Before","Do Not Del After","Do Not Deliver Before","Do Not Deliver After",
                //    "Estimated Delivery Date","Delivery Requested Date","Requested Delivery Date","Deliver Before Date","Req Delivery date","Schedule Delivery date","Sch Del Date","REQ ARV DTE","Est. Arrival Date",
                //    //"DE L DATE",
                //    "Delivery Date","Guaranteed Before","Deliver On","Delivery Appointment","Deliv From","Delivery Date/ Time","Due Date","To Arrive","Estimated Delivery","Requested Date/ Time","DLV Date",
                //    "REQ DATE","ARRIVE DATE","Delv Date","Del Date","Delivery Dt","Appt Date/Time","Delivery Window","Est. Arrival","Request Date","Dropoff Date/Time",
                //    "DELIVERY MUST BE MADE BY","MUST DELIVER ON OR BEFORE","Must Arrive By Date","Must Arrive By Window","Must Deliver by date","MUST DELIVER BETWEEN",
                //    "Must Deliver By","ARRIVE BY","Deliver By","MABD Date","Delivery","MABD"
                //};
                ///
                #endregion

                #endregion


                 Array.Sort(sKeyarray); //commented 03/06/2021 for keyword break

                DataTable dt = oMeta.DocDictionary;

                System.Data.DataRow[] foundRows;


                int max = Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == intCurrPageNo.ToString()).Max(row => row["Line_No"]));

                for (int iLine = 1; iLine <= max; iLine++)
                {
                    foundRows = dt.Select("[Page_no]=" + intCurrPageNo + " AND [Line_no]=" + iLine);
                    //foreach (DataRow drDataRow in foundRows)
                    int RowCount = foundRows.Length;
                    for (int iWord = 0; iWord < RowCount; iWord++)
                    {
                        // int WordCount= Convert.ToInt32(dt.AsEnumerable().Where(row => row["Page_No"].ToString() == "1" && row["Line_No"].ToString() ==iLine.ToString()).Max(row => row["Word_No"]));

                        int WordNo = Convert.ToInt32(foundRows[iWord]["Word_no"]);
                        //string Word = foundRows[iWord]["word_Ntext"].ToString();
                        string Word = foundRows[iWord]["word_Otext"].ToString();
                        int WordLength = Convert.ToInt32(foundRows[iWord]["Word_Length"]);
                        string WordType = foundRows[iWord]["Word_Type"].ToString();
                        int lengthDiff;
                        // bool mergeQouteflag = false;
                        string QNoCheck = string.Empty;
                        if (!string.IsNullOrEmpty(WordType))
                        {
                            if (WordType.ToUpper() != "N")
                            {
                                if (!Word.ToUpper().Trim().Contains("FREIGHT"))
                                {
                                    if (Word.Length > 7 && Word.Substring(0, 5).ToUpper() == "QUOTE")
                                    {
                                        QNoCheck = Word.ToUpper().Replace("#", "").Replace(":", "").Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Trim();
                                        if ((Module1.reg1.IsMatch(QNoCheck) || Module1.reg2.IsMatch(QNoCheck)))
                                        {
                                            DataRow dr;

                                            dr = dtKeywordInfo.NewRow();
                                            dr["Keyword"] = "QUOTE";//sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                            dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
                                            dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
                                            dr["Word No"] = WordNo;
                                            // dr["LineWordNo"] = WordNo;
                                            dr["X1"] = foundRows[iWord]["X1"];
                                            dr["Y1"] = foundRows[iWord]["Y1"];
                                            dr["X2"] = foundRows[iWord]["X2"];
                                            dr["Y2"] = foundRows[iWord]["Y2"];
                                            dr["NoofWords"] = 1;//Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                            dtKeywordInfo.Rows.Add(dr);

                                            ////// Sorted keyword Logic
                                            //continue;
                                            ////// Sorted keyword Logic
                                        }


                                    }
                                    for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
                                    {
                                        var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
                                        lengthDiff = WordLength - skeysplit[0].Length;
                                        if (skeysplit.Length > 1)
                                        {
                                            if (lengthDiff >= -2 && lengthDiff <= 2)
                                            {
                                                if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                                {
                                                    if (iWord >= RowCount)
                                                    {
                                                    }
                                                    else
                                                    {
                                                        int iIndex;
                                                        var loopTo3 = skeysplit.Length - 1;
                                                        for (iIndex = 1; iIndex <= loopTo3; iIndex++)
                                                        {
                                                            if (iWord + iIndex >= RowCount)
                                                                continue;
                                                            lengthDiff = Convert.ToInt32(foundRows[iWord + iIndex]["Word_Length"]) - skeysplit[iIndex].Length;
                                                            if (lengthDiff >= -2 && lengthDiff <= 2)
                                                            {
                                                                //if (CheckValidWord(foundRows[iWord + iIndex]["word_Ntext"].ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
                                                                if (CheckValidWord(foundRows[iWord + iIndex]["word_Otext"].ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
                                                                {
                                                                    if (skeysplit.Length - 1 != iIndex)
                                                                    {
                                                                        continue;
                                                                    }
                                                                    else
                                                                    {
                                                                        DataRow dr;

                                                                        dr = dtKeywordInfo.NewRow();
                                                                        dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                                        dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
                                                                        dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
                                                                        dr["Word No"] = WordNo;
                                                                        //dr["LineWordNo"] = WordNo;
                                                                        dr["X1"] = foundRows[iWord]["X1"];
                                                                        dr["Y1"] = foundRows[iWord]["Y1"];
                                                                        dr["X2"] = foundRows[iWord]["X2"];
                                                                        dr["Y2"] = foundRows[iWord]["Y2"];
                                                                        dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                                        dtKeywordInfo.Rows.Add(dr);

                                                                        ////New added to break Loop  03/06/2021
                                                                        //iIndex = loopTo3;
                                                                        //jLoop = sKeyarray.Length;
                                                                        //iWord = iWord + loopTo3;
                                                                        //break;
                                                                        ////New added to break Loop
                                                                    }
                                                                }
                                                                else
                                                                {
                                                                    break;
                                                                }
                                                            }
                                                            // For MergkeywordData  
                                                            else if (iIndex == loopTo3)
                                                            {
                                                                string word = foundRows[iWord + iIndex]["word_Ntext"].ToString();
                                                                if (foundRows[iWord + iIndex]["word_Ntext"].ToString().ToUpper().StartsWith(skeysplit[iIndex].ToUpper()))
                                                                {
                                                                    if (foundRows[iWord + iIndex]["word_Ntext"].ToString().ToUpper().Any(char.IsDigit))
                                                                    {
                                                                        DataRow dr;

                                                                        dr = dtKeywordInfo.NewRow();
                                                                        dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                                        dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
                                                                        dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
                                                                        dr["Word No"] = WordNo;
                                                                        //dr["LineWordNo"] = WordNo;
                                                                        dr["X1"] = foundRows[iWord]["X1"];
                                                                        dr["Y1"] = foundRows[iWord]["Y1"];
                                                                        dr["X2"] = foundRows[iWord]["X2"];
                                                                        dr["Y2"] = foundRows[iWord]["Y2"];
                                                                        dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                                        dtKeywordInfo.Rows.Add(dr);

                                                                        ////New added to break Loop  03/06/2021
                                                                        //iIndex = loopTo3;
                                                                        //jLoop = sKeyarray.Length;
                                                                        //iWord = iWord + loopTo3;
                                                                        //break;
                                                                        ////New added to break Loop
                                                                    }
                                                                }
                                                            }
                                                            // For MergkeywordData
                                                            else
                                                            { break; }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        else if (lengthDiff >= -2 && lengthDiff <= 2)
                                        {
                                            if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                            {
                                                DataRow dr;

                                                dr = dtKeywordInfo.NewRow();
                                                dr["Keyw" +
                                                    "ord"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                dr["Page No"] = Convert.ToInt32(foundRows[iWord]["Page_no"]);
                                                dr["Line No"] = Convert.ToInt32(foundRows[iWord]["Line_no"]);
                                                dr["Word No"] = WordNo;
                                                // dr["LineWordNo"] = WordNo;
                                                dr["X1"] = foundRows[iWord]["X1"];
                                                dr["Y1"] = foundRows[iWord]["Y1"];
                                                dr["X2"] = foundRows[iWord]["X2"];
                                                dr["Y2"] = foundRows[iWord]["Y2"];
                                                dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                dtKeywordInfo.Rows.Add(dr);

                                                ////New added to break Loop  03/06/2021
                                                //jLoop = sKeyarray.Length;
                                                //break;
                                                ////New added to break Loop
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                // throw;
                //MessageBox.Show(ex.Message, "SearchKeywords");
            }

            return dtKeywordInfo;
        }

        #region
        public static DataTable F3_SearchKeyword(clsCncMetaData oMeta, int intCurrPageNo)
        {
            var dtKeywordInfo = new DataTable();
            dtKeywordInfo = MakeDt();
            try
            {
                // -----------------------------------------INVOICE NUMBER/INVOICE DATE/PURCHASE ORDER NUMBER/INOICE AMOUNT KEYWORD------------------------------------------------------------
                #region Original Keywords
                ////Original Keywords
                ////var sKeyarray = new string[]
                ////    {
                ////    "SPECIAL INSTRUCTION", "Comments/Special Instruction", "Shipping Instructions","Consignee Special Instructions","Shipper Special Instructions", "Special Services and Instructions",
                ////    "Carrier Special delivery Instructions","Origin Instructions","SpecialInstructions","Special delivery Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction",
                ////    "COD AMOUNT", "COD AMT", "COD",
                ////    "Due Date","Delivery Date","Estimated Delivery Date","Guaranteed Before","Deliver On","Delivery Requested Date","Requested Delivery Date","Delivery Appointment",
                ////    "Deliv From","Delivery Date/Time","Deliver Before Date","To Arrive","Deliver No Earlier Than","Deliver No Later Than","Req Delivery date","Schedule Delivery date","Sch Del Date",
                ////    "Estimated Delivery","Requested Date/Time","REQ ARV DTE","DLV Date","REQ DATE","ARRIVE DATE","Est. Arrival Date","Delv Date","Del Date","Delivery Dt","Delivery","Appt Date/Time",
                ////    "Delivery Window","DE L DATE","Appt. Date / Time","Est. Arrival","Request Date","Must deliver only on","Dropoff Date/Time","Do Not Del Before","Do Not Del After","Do Not Deliver Before","Do Not Deliver After",
                ////    "MABD","MABD Date","Must Arrive By Date","Must Deliver by date","Must Deliver By","Deliver By","ARRIVE BY","MUST DELIVER ON OR BEFORE","MUST DELIVER BETWEEN","Must Arrive By Window","DELIVERY MUST BE MADE BY",
                ////    "Quote", "Quote No", "Quote Number","QuoteID","Quote Id","Quote ID Number","Quote Reference ID","QuoteNumber","Quot e","Q#",
                ////    "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","GRANDTOTAL","NetWeight","GrossWeight","TotalAmount","Weight Total","TOTAL WT","TotalWeight","PAGE SUBTOTAL","Weight","Total","Total Gross","Gross",
                ////    "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                ////    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                ////    "Pallet Count","Total Skids","Colis/Package","Package Count","H. Units","HU","No. of Pallets", "Number of Pallets","TOTAL PALLETS",
                ////    "TOTAL PALLE TS","Handling Units Total","Total Handling Units","H/Us","Handling Units","TOTAL PLTS" //,"TOTAL PKGS","NO. OF PKGS","No. Packages","Total Pieces", "Total","TOTAL PCS","QTY","Pallets"
                ////    };
                ////Original Keywords
                #endregion


                #region  Sorted Keywords
                ////Sorted Keywords
                //var sKeyarray = new string[]
                //{ "Special Services and Instructions","Carrier Special delivery Instructions","Comments/Special Instruction","Consignee Special Instructions","Shipper Special Instructions",
                //    "Special delivery Instructions", "Shipping Instructions","Origin Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction","SPECIAL INSTRUCTION","SpecialInstructions",
                //    "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                //    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                //    //"COD AMOUNT","COD AMT","COD",
                //    "Quote ID Number","Quote Reference ID","Quote No","Quote Number","Quote Id","QuoteID","QuoteNumber","Quote",
                //    //"Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","Weight Total","TOTAL WT","PAGE SUBTOTAL","Total Gross","GRANDTOTAL","NetWeight","GrossWeight",
                //    //"TotalAmount","TotalWeight","Weight","Total","Gross",
                //    "Number of Pallets","No. of Pallets","Handling Units Total","Total Handling Units","Pallet Count","Total Skids","Package Count","H. Units","TOTAL PALLETS",
                //    "Handling Units","TOTAL PLTS","H/Us","HU",
                //    "Deliver No Earlier Than","Deliver No Later Than","Appt. Date / Time","Must deliver only on","Do Not Del Before","Do Not Del After","Deliver Before Date","Do Not Deliver Before","Do Not Deliver After",
                //    "Estimated Delivery Date","Delivery Requested Date","Requested Delivery Date","Req Delivery date","Schedule Delivery date","Sch Del Date","REQ ARV DTE","Est. Arrival Date",
                //    "Delivery Date/Time","Delivery Date","Guaranteed Before","Deliver On","Delivery Appointment","Deliv From","Due Date","To Arrive","Estimated Delivery","Requested Date/ Time","DLV Date",
                //    "REQ DATE","ARRIVE DATE","Delv Date","Del Date","Delivery Dt","Appt Date/Time","Delivery Window","Est. Arrival","Request Date","Dropoff Date/Time",
                //    "DELIVERY MUST BE MADE BY","MUST DELIVER ON OR BEFORE","Must Arrive By Date","Must Arrive By Window","Must Deliver by date","MUST DELIVER BETWEEN",
                //    "Must Deliver By","ARRIVE BY","Deliver By","MABD Date","Delivery","MABD"
                //};


                //Sorted Keywords
                #region Sorted backup
                ///Sorted backup
                //var sKeyarray = new string[]
                //{ "Special Services and Instructions","Carrier Special delivery Instructions","Comments/Special Instruction","Consignee Special Instructions","Shipper Special Instructions",
                //    "Special delivery Instructions", "Shipping Instructions","Origin Instructions","Shipper Instructions","Shipper Instruction","Consignee Instruction","SPECIAL INSTRUCTION","SpecialInstructions",
                //    "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                //    "DESTINATION Accessorials","Delivery Accessorials","Accessorial Required","Accessorials",
                //    "COD AMOUNT","COD AMT","COD",
                //    "Quote ID Number","Quote Reference ID","Quote No","Quote Number","Quote Id","QuoteID","QuoteNumber","Quote",
                //    //"Quot e","Q#",
                //    "Gross Weight","Net Weight","Total Amount","Grand Total","Total Weight","Gross WT","Weight Total","TOTAL WT","PAGE SUBTOTAL","Total Gross","GRANDTOTAL","NetWeight","GrossWeight",
                //    "TotalAmount","TotalWeight","Weight","Total","Gross",
                //    "Number of Pallets","No. of Pallets","TOTAL PALLE TS","Handling Units Total","Total Handling Units","Pallet Count","Total Skids","Package Count","H. Units","TOTAL PALLETS",
                //    "Handling Units","TOTAL PLTS","H/Us","HU",//"Colis/Package",
                //    "Deliver No Earlier Than","Deliver No Later Than","Appt. Date / Time","Must deliver only on","Do Not Del Before","Do Not Del After","Do Not Deliver Before","Do Not Deliver After",
                //    "Estimated Delivery Date","Delivery Requested Date","Requested Delivery Date","Deliver Before Date","Req Delivery date","Schedule Delivery date","Sch Del Date","REQ ARV DTE","Est. Arrival Date",
                //    //"DE L DATE",
                //    "Delivery Date","Guaranteed Before","Deliver On","Delivery Appointment","Deliv From","Delivery Date/ Time","Due Date","To Arrive","Estimated Delivery","Requested Date/ Time","DLV Date",
                //    "REQ DATE","ARRIVE DATE","Delv Date","Del Date","Delivery Dt","Appt Date/Time","Delivery Window","Est. Arrival","Request Date","Dropoff Date/Time",
                //    "DELIVERY MUST BE MADE BY","MUST DELIVER ON OR BEFORE","Must Arrive By Date","Must Arrive By Window","Must Deliver by date","MUST DELIVER BETWEEN",
                //    "Must Deliver By","ARRIVE BY","Deliver By","MABD Date","Delivery","MABD"
                //};
                ///
                #endregion

                #endregion


                #region "KeyWord Removed which F3 not required"
                var sKeyarray = new string[]
                {   "DESTINATION INSTRUCTIONS","Delivery Instructions","Delivery Notes","Delivery Remarks",
                    "Quote ID Number","Quote Reference ID","Quote No","Quote Number","Quote Id","QuoteID","QuoteNumber","Quote",
                    "Number of Pallets","No. of Pallets","Handling Units Total","Total Handling Units","Pallet Count","Total Skids","Package Count","H. Units","TOTAL PALLETS",
                    "Handling Units","TOTAL PLTS","H/Us","HU",
                };
                #endregion
                ////Array.Sort(sKeyarray); //commented 03/06/2021 for keyword break

                int max = oMeta.Page[intCurrPageNo].LineCount;
                for (int iLine = 1; iLine <= max; iLine++)
                {
                    int RowCount = oMeta.Page[intCurrPageNo].Line[iLine].WordCount;
                    for (int iWord = 1; iWord <= RowCount; iWord++)
                    {
                        clsCnCWord cncWord = oMeta.Page[intCurrPageNo].Line[iLine].Word[iWord];
                        // new changes
                        int WordNo = iWord;
                        string Word = oMeta.Page[intCurrPageNo].Line[iLine].Word[iWord].strWord;
                        int WordLength = Word.Length;
                        int lengthDiff;
                        string QNoCheck = string.Empty;
                        //

                        if (!string.IsNullOrEmpty(Word))
                        {
                            if (!Word.All(char.IsNumber))
                            {

                                if (!Word.ToUpper().Trim().Contains("FREIGHT"))
                                {
                                    if (Word.Length > 7 && Word.Substring(0, 5).ToUpper() == "QUOTE")
                                    {
                                        QNoCheck = Word.ToUpper().Replace("#", "").Replace(":", "").Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Trim();
                                        ////if ((Module1.reg1.IsMatch(QNoCheck) || Module1.reg2.IsMatch(QNoCheck)))
                                        ////{
                                            DataRow dr;

                                            dr = dtKeywordInfo.NewRow();
                                            dr["Keyword"] = "QUOTE";//sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                            dr["Page No"] = cncWord.PageNo;
                                            dr["Line No"] = cncWord.LineNo;
                                            dr["Word No"] = WordNo;
                                            //dr["LineWordNo"] = WordNo;
                                            dr["X1"] = cncWord.Left;
                                            dr["Y1"] = cncWord.Top;
                                            dr["X2"] = cncWord.Right;
                                            dr["Y2"] = cncWord.Bottom;
                                            dr["NoofWords"] = 1;
                                            dtKeywordInfo.Rows.Add(dr);

                                            //// Sorted keyword Logic
                                            continue;
                                            //// Sorted keyword Logic
                                       //// }


                                    }
                                    for (int jLoop = 0; jLoop <= sKeyarray.Length - 1; jLoop++)
                                    {
                                        var skeysplit = sKeyarray[jLoop].Trim().Split(' ');
                                        lengthDiff = WordLength - skeysplit[0].Length;
                                        if (skeysplit.Length > 1)
                                        {
                                            if (lengthDiff >= -2 && lengthDiff <= 2)
                                            {
                                                if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                                {
                                                    if (iWord > RowCount)
                                                    {
                                                    }
                                                    else
                                                    {
                                                        int iIndex;
                                                        var loopTo3 = skeysplit.Length - 1;
                                                        for (iIndex = 1; iIndex <= loopTo3; iIndex++)
                                                        {
                                                            if (iWord + iIndex > RowCount)
                                                                continue;

                                                            clsCnCWord NextWord = oMeta.Page[intCurrPageNo].Line[iLine].Word[iWord + iIndex];

                                                            if (NextWord != null)
                                                            {

                                                                lengthDiff = NextWord.strWord.Length - skeysplit[iIndex].Length;
                                                                if (lengthDiff >= -2 && lengthDiff <= 2)
                                                                {
                                                                    //if (CheckValidWord(foundRows[iWord + iIndex]["word_Ntext"].ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
                                                                    if (CheckValidWord(NextWord.strWord.ToString().ToUpper().Trim(), new string[] { skeysplit[iIndex] }) == true)
                                                                    {
                                                                        if (skeysplit.Length - 1 != iIndex)
                                                                        {
                                                                            continue;
                                                                        }
                                                                        else
                                                                        {
                                                                            DataRow dr;

                                                                            dr = dtKeywordInfo.NewRow();
                                                                            dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                                            dr["Page No"] = cncWord.PageNo;
                                                                            dr["Line No"] = cncWord.LineNo;
                                                                            dr["Word No"] = WordNo;
                                                                            //dr["LineWordNo"] = WordNo;
                                                                            dr["X1"] = cncWord.Left;
                                                                            dr["Y1"] = cncWord.Top;
                                                                            dr["X2"] = cncWord.Right;
                                                                            dr["Y2"] = cncWord.Bottom;
                                                                            dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                                            dtKeywordInfo.Rows.Add(dr);

                                                                            //New added to break Loop  03/06/2021
                                                                            iIndex = loopTo3;
                                                                            jLoop = sKeyarray.Length;
                                                                            iWord = iWord + loopTo3;
                                                                            break;
                                                                            //New added to break Loop
                                                                        }
                                                                    }
                                                                    else
                                                                    {
                                                                        break;
                                                                    }
                                                                }
                                                                // For MergkeywordData  
                                                                else if (iIndex == loopTo3)
                                                                {
                                                                    ////string word = foundRows[iWord + iIndex]["word_Ntext"].ToString();
                                                                    if (NextWord.strWord.ToString().ToUpper().StartsWith(skeysplit[iIndex].ToUpper()))
                                                                    {
                                                                        if (NextWord.strWord.ToString().ToUpper().Any(char.IsDigit))
                                                                        {
                                                                            DataRow dr;

                                                                            dr = dtKeywordInfo.NewRow();
                                                                            dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                                            dr["Page No"] = cncWord.PageNo;
                                                                            dr["Line No"] = cncWord.LineNo;
                                                                            dr["Word No"] = WordNo;
                                                                            ////dr["LineWordNo"] = WordNo;
                                                                            dr["X1"] = cncWord.Left;
                                                                            dr["Y1"] = cncWord.Top;
                                                                            dr["X2"] = cncWord.Right;
                                                                            dr["Y2"] = cncWord.Bottom;
                                                                            dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                                            dtKeywordInfo.Rows.Add(dr);

                                                                            //New added to break Loop  03/06/2021
                                                                            iIndex = loopTo3;
                                                                            jLoop = sKeyarray.Length;
                                                                            iWord = iWord + loopTo3;
                                                                            break;
                                                                            //New added to break Loop
                                                                        }
                                                                    }
                                                                }
                                                                // For MergkeywordData
                                                                else
                                                                { break; }
                                                            }
                                                        }
                                                    }
                                                }
                                            }
                                        }

                                        else if (lengthDiff >= -2 && lengthDiff <= 2)
                                        {
                                            if (CheckValidWord(Word.ToUpper().Trim(), new string[] { skeysplit[0] }) == true)
                                            {
                                                DataRow dr;

                                                dr = dtKeywordInfo.NewRow();
                                                dr["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
                                                dr["Page No"] = cncWord.PageNo;
                                                dr["Line No"] = cncWord.LineNo;
                                                dr["Word No"] = WordNo;
                                                ////dr["LineWordNo"] = WordNo;
                                                dr["X1"] = cncWord.Left;
                                                dr["Y1"] = cncWord.Top;
                                                dr["X2"] = cncWord.Right;
                                                dr["Y2"] = cncWord.Bottom;
                                                dr["NoofWords"] = Convert.ToInt32(sKeyarray[jLoop].Trim().ToUpper().Replace("'", "").Split(new char[] { ' ' }).Length);
                                                dtKeywordInfo.Rows.Add(dr);

                                                //New added to break Loop  03/06/2021
                                                jLoop = sKeyarray.Length;
                                                break;
                                                //New added to break Loop
                                            }
                                        }
                                    }
                                }
                            }
                        }
                    }

                }

            }
            catch (Exception ex)
            {
                // throw;
                //MessageBox.Show(ex.Message, "SearchKeywords");
            }

            return dtKeywordInfo;
        }
        #endregion

        public static DataTable MakeDt()
        {
            var dtKeyword = new DataTable();
            dtKeyword.Columns.Add("Keyword", typeof(string));
            dtKeyword.Columns.Add("Page No", typeof(string));
            dtKeyword.Columns.Add("Line No", typeof(string));
            dtKeyword.Columns.Add("Word No", typeof(string));
            dtKeyword.Columns.Add("X1", typeof(string));
            dtKeyword.Columns.Add("Y1", typeof(string));
            dtKeyword.Columns.Add("X2", typeof(string));
            dtKeyword.Columns.Add("Y2", typeof(string));
            dtKeyword.Columns.Add("NoofWords", typeof(string));
            return dtKeyword;
        }
        //new
        private static bool CheckValidWord(string sWord, string[] arrcheckkeywords)
        {
            bool bFlag = false;
            int lengthofword, KeywordLength;

            var objRestruct056 = new iQPUBLIC.clsCNCSBR();
            iQPUBLIC.RetStructIQSBR058 objRestruct058;

            var ExactMatchKeywords = new List<string>() { "QuoteID","QuoteNumber","SpecialInstructions", "GRANDTOTAL", "NetWeight", "GrossWeight", "TotalAmount", "TotalWeight" };
            ExactMatchKeywords = ExactMatchKeywords.ConvertAll(s=>s.ToUpper());
            if (!string.IsNullOrEmpty(sWord))
            {
                // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
                var arrFinalkeyword = new List<string>();
                sWord = sWord.Replace("'", "");
                sWord = sWord.Replace("`", "");

                //if (SpecialCharacters.Any(sWord.Contains))
                if (SpecialCharacters.Any(sWord.Contains) && sWord.Length > 1)
                {
                    sWord = Regex.Replace(sWord, @"[^0-9a-zA-Z]+", "");
                    // sWord = RemoveSpecialCharacters(sWord);
                }
                if (SpecialCharacters.Any(arrcheckkeywords[0].Contains))
                {
                    arrcheckkeywords[0] = Regex.Replace(arrcheckkeywords[0], @"[^0-9a-zA-Z]+", "");
                    // sWord = RemoveSpecialCharacters(sWord);
                }
                lengthofword = sWord.Length;
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace("'", "");
                arrcheckkeywords[0] = arrcheckkeywords[0].Replace("`", "");
                KeywordLength = arrcheckkeywords[0].Length;

                //new added 14/06/2021
                if (sWord.ToUpper() == arrcheckkeywords[0].ToUpper())
                {
                    return true;
                }
                // new added

                if (KeywordLength <= 3)
                {
                    if (lengthofword <= 3)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }
                else if (KeywordLength > 3 & KeywordLength <= 5)
                {
                    if (lengthofword > 3 && lengthofword <= 7)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }

                }
                else if (KeywordLength >= 6 & KeywordLength <= 8)
                {

                    if (lengthofword >= 4 && lengthofword <= 10)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }

                }
                else if (KeywordLength >= 9 & KeywordLength <= 11)
                {
                    if (lengthofword >= 6 && lengthofword <= 14)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }
                else if (KeywordLength >= 12)
                {
                    if (lengthofword >= 8)
                    {
                        arrFinalkeyword.Add(arrcheckkeywords[0]);
                    }
                }

                for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)
                {
                        objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                        if (SpecialCharacters.Any(arrFinalkeyword[iLoop].Contains))
                        {
                            arrFinalkeyword[iLoop] = Regex.Replace(arrFinalkeyword[iLoop], @"[^0-9a-zA-Z]+", "");
                            // sWord = RemoveSpecialCharacters(sWord);
                        }
                        int lengthofkeyword = arrFinalkeyword[iLoop].Length;
                        //new added for exactmatching of merge word
                        if (ExactMatchKeywords.Contains(arrFinalkeyword[iLoop].ToUpper()))
                        {
                            if (objRestruct058.CharChanged == 0)
                            {
                                bFlag = true;
                            }
                        }
                        //
                       else if (lengthofkeyword <= 3)
                        {
                            if (objRestruct058.CharChanged < 1)
                            {
                                bFlag = true;
                            }
                        }
                       else if ((lengthofkeyword > 3 && lengthofkeyword <= 5) && lengthofword > 3)
                        {
                            if (objRestruct058.CharChanged <= 1)
                            {
                                bFlag = true;
                            }
                        }
                        else if (lengthofkeyword >= 6 & lengthofkeyword <= 8)
                        {
                            if (arrFinalkeyword[iLoop].ToUpper() == "ARRIVE")
                            {
                                if (objRestruct058.CharChanged <= 1)
                                {
                                    bFlag = true;
                                }
                            }
                            else if (objRestruct058.CharChanged <= 2)
                            {
                                bFlag = true;
                            }
                        }
                        else if (lengthofkeyword >= 9 & lengthofkeyword <= 11)
                        {
                            if (objRestruct058.CharChanged <= 3)
                            {
                                bFlag = true;
                            }
                        }
                        else if (lengthofkeyword >= 12)
                        {
                            if (objRestruct058.CharChanged <= 4)
                            {
                                bFlag = true;
                            }
                        }
                   
                    
                }
            }

            return bFlag;
        }
        
        #endregion


        #region Other Required Functions

        public static string RoundOff(string str)
        {
            decimal Value;
            if (decimal.TryParse(str, out Value))
            {
                Value = Math.Round(Value);
                str = Value.ToString();
            }
            return str;
        }

        private static string RemoveSpecialCharacters(string StrWord)
        {
            string strRetWord = string.Empty;
            try
            {
                //char[] SpecialCharacters = new char[] { '!','@','#','$','%','^','&','*','(',')','_','+','-','=',',','.',':',';','/'};
                // var MatchValues = SpecialCharacters.Where(StrWord => StrWord.Contains())

                ; foreach (string ch in SpecialCharacters)
                {
                    if (StrWord.Contains(ch))
                    {
                        StrWord = StrWord.Replace(ch, "");
                    }
                }

            }
            catch (Exception ex)
            {
                // throw;
            }

            return StrWord;

        }
        public static DataTable InstructionDt()
        {
            var dtKeyword = new DataTable();
            //dtKeyword.Columns.Add("Id", typeof(int));
            dtKeyword.Columns.Add("Instruction", typeof(string));
            dtKeyword.Columns.Add("Code", typeof(string));
            dtKeyword.Columns.Add("LineNo", typeof(int));
            dtKeyword.Columns.Add("WordNo", typeof(int));
            dtKeyword.Columns.Add("DbSI", typeof(string));
            dtKeyword.Columns.Add("NoOfWord", typeof(int));
            return dtKeyword;
        }

        public static DataTable AssignInstructionDt()
        {
            Module1.Dt_distinctInstruction = InstructionDt();
            return Module1.Dt_distinctInstruction;
        }
        #endregion
    }
}