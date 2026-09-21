using iQDataProvider;
using iQPUBLIC;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data;
using System.Linq;
using System.Collections;
using System.Globalization;
using System.IO;

namespace BOLSET3IQ
{
    public class clsF3
    {
        //BOL_SET3.clsF3 objBOLF3;
        // Id for functions:
        // Taoal pieces =1;
        // Taoal Weight =2;
        // COD =3;
        // Quote No = 4;
        // RADF =5;
        //Common Instructions = 6;

        public RetStructF3 F3(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strRoutineNo, string strArg)
        {
            string[] strCordinates = strArg.Split('~');
            Module1.MX1 = Convert.ToInt32(strCordinates[0]);
            Module1.MY1 = Convert.ToInt32(strCordinates[1]);
            Module1.MX2 = Convert.ToInt32(strCordinates[2]);
            Module1.MY2 = Convert.ToInt32(strCordinates[3]);

            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;
            //objRetStructF3.Status = "F";
            //objRetStructF3.NoOfieldsSuspects = 0;

            if (Module1.imageName.ToLower() != Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath.ToLower()))
            {
                Module1.bFoundSearchKeywordTable = false;
                Module1.imageName = Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath);
                Module1.RADF_PDT = string.Empty;
                Module1.MABD_PDT = string.Empty;
            }

            //objBOLF3 = new BOL_SET3.clsF3();


            switch (strRoutineNo)
            {
                case "CKM249"://Total Pieces
                    Module1.objRetStructF3 = F3_AutoLocate_TotalPieces(ObjMetaData, intCurrPageNumber, strArg);
                    // objRetStructF3 = F3_AutoLocate_TotalPieces(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM250"://Total Weight F3_AutoLocate_NewTotalWeight
                    Module1.objRetStructF3 = F3_AutoLocate_TotalWeight(ObjMetaData, intCurrPageNumber, strArg);
                    // objRetStructF3 = F3_AutoLocate_TotalWeight(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM216"://COD
                    Module1.objRetStructF3 = F3_AutoLocate_COD(ObjMetaData, intCurrPageNumber, strArg);
                    // objRetStructF3 = F3_AutoLocate_COD(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM251"://Special Instruction
                   Module1.objRetStructF3 = F3_AutoLocate_Special_Instruction(ObjMetaData, intCurrPageNumber, strArg);
                   // objRetStructF3 = F3_AutoLocate_Special_Instruction(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM220"://Delivery Quote No
                    Module1.objRetStructF3 = F3_AutoLocate_QUOTENO(ObjMetaData, intCurrPageNumber, strArg);
                    //objRetStructF3 = F3_AutoLocate_QUOTENO(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM254"://RADF
                    Module1.objRetStructF3 = F3_AutoLocate_RADF(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM255"://MABD
                    Module1.objRetStructF3 = F3_AutoLocate_MABD(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM217"://DLV
                    Module1.objRetStructF3 = F3_AutoLocate_DLV(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM253"://Delivery Requirments
                    Module1.objRetStructF3 = F3_AutoLocate_Delivery_Requirment(ObjMetaData, intCurrPageNumber, strArg);
                    break;
                case "CKM252"://Delivery Requirments
                    Module1.objRetStructF3 = F3_AutoLocate_Accessorial(ObjMetaData, intCurrPageNumber, strArg);
                    break;


            }

            return Module1.objRetStructF3;
            //return  objRetStructF3;

        }

        #region Total pieces
        private RetStructF3 F3_AutoLocate_TotalPieces(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {

            //DataSet ds = new DataSet();
            DataSet ds = Module1.func_GetKeywords(ObjMetaData, intCurrPageNumber);

            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            DataSet dsKeys = new DataSet();
            DataTable dt = new DataTable();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;
            string TotalWeight = "";
            string bWord = "";
            bool UnitFlag = false;
           

            System.Data.DataRow[] foundRows;
            ClsCNC objCNC = new ClsCNC();
            clsCnCWord cWord = new clsCnCWord();
            clsCnCWord CombinedKeyword = new clsCnCWord();

            //string[] palletCount = { "pallet count", "No. of Pallets", "Number of Pallets", "TOTAL PALLETS", "Handling Units Total", "Total Handling Units" };
            string[] palletCount = { "pallet count"};
            DataTable dt_Pallet = new DataTable();
            dt_Pallet.Columns.Add("keyword");
            dt_Pallet.Columns.Add("Line No");
            dt_Pallet.Columns.Add("Word No");
            dt_Pallet.Columns.Add("flag");


        Line1:
            ;

            if (iRecursiveCallCount == 0)
            {
                // -------------------------------------------------PIECES KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                //sKeyword = "'Pallet Count','NO. OF PKGS','No. Packages','Total Skids','Total Pieces','Colis/Package','Package Count','H. Units','HU','QTY','Pallets', 'Total','TOTAL PCS','TOTAL CARTONS','TOTALCARTONS','TOTAL PKGS','Pallets',";
                sKeyword = "'Pallet Count','Total Skids','Colis/Package','Package Count','H. Units','HU','No. of Pallets','Number of Pallets','TOTAL PALLETS','TOTAL PALLE TS','Handling Units Total','Total Handling Units','H/Us','Handling Units','TOTAL PLTS'";

            }
            // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            else
            {
                sKeyword = "''";
            }

            try
            {
                returnZones.Status = "F";
                returnZones.ManualConfirmation = "Y";
                returnZones.Flag = "Y";

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalPieces"))
                {
                    TotalWeight = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_TotalPieces"];
                }


                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_keyword"];
                }
                if (dsKeys != null)
                {
                    dt = dsKeys.Tables[0];
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();
                    string keywordFound = "";
                    int pgno = 0;
                    int lineno = 0;

                    int wordno = 0;
                    int Newwordno = 0;

                    //only for pallet count
                    if (foundRows.Length > 0)
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            for (int iarr = 0; iarr < palletCount.Length; iarr++)
                            {
                                if (keywordFound.ToUpper() == palletCount[iarr].ToUpper())
                                {
                                    DataRow Dr = dt_Pallet.NewRow();
                                    Dr["keyword"] = keywordFound;
                                    Dr["Line No"] = lineno;
                                    Dr["Word No"] = wordno;
                                    Dr["flag"] = 1;
                                    dt_Pallet.Rows.Add(Dr);
                                }
                            }

                        }
                    }
                    if (dt_Pallet.Rows.Count > 0)
                    {
                        for (int i = 0; i < dt_Pallet.Rows.Count; i++)
                        {
                            keywordFound = Convert.ToString(dt_Pallet.Rows[i]["keyword"]).Trim();
                            lineno = Convert.ToInt32(dt_Pallet.Rows[i]["Line No"]);
                            wordno = Convert.ToInt32(dt_Pallet.Rows[i]["Word No"]);

                            string[] splitKeword = keywordFound.Split(' ');
                            if (splitKeword.Length > 1)
                            {
                                Newwordno = wordno + splitKeword.Length - 1;
                            }
                            else
                            {
                                Newwordno = wordno;
                            }

                            CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);
                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);
                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                                if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "&" || Module1.ObjBoundingWord.RightWord.strWord == "$")
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                }
                                if (Module1.ObjBoundingWord.RightWord != null)
                                {
                                    Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                    if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.RightWord.strWord))
                                    {

                                        bWord = Module1.RoundOff(Module1.ObjBoundingWord.RightWord.strWord);
                                        if (bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                        {
                                            // New added
                                            Module1.ObjBoundingWord.RightWord.strWord = bWord;
                                            // New added

                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                            {
                                                PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                            }
                                        }

                                    }
                                }
                            }
                        }
                        if (PossibleWords.Count > 1)
                        {
                            CompairConfidence(PossibleWords);
                        }
                        if (PossibleWords.Count == 1)
                        {
                            Boolean ConfidenceFlag = HighestConfidence(PossibleWords[0]);
                            if (ConfidenceFlag == true)
                            {
                                ////PossibleWords[0].Confidence = 100;
                                if (PossibleWords[0].Flag == null && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                {
                                    PossibleWords[0].Flag = "1";
                                    PossibleWords[0].Remarks = "1F";
                                }
                                else
                                {
                                    PossibleWords[0].Flag = "0";
                                    PossibleWords[0].Remarks = "NA";
                                }
                            }
                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                        }
                    }
                    if (PossibleWords.Count == 0)
                    {

                        if (foundRows.Length == 1)
                        {
                            keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                            pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                            lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                            wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);


                            string[] splitKeword = keywordFound.Split(' ');
                            if (splitKeword.Length > 1)
                            {
                                Newwordno = wordno + splitKeword.Length - 1;
                            }
                            else
                            {
                                Newwordno = wordno;
                            }

                            CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                            if (!string.IsNullOrEmpty(TotalWeight))
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);
                                if (Module1.ObjBoundingWord.RightWord != null)
                                {
                                    if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "&" || Module1.ObjBoundingWord.RightWord.strWord == "$")
                                    {
                                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                    }
                                    if (Module1.ObjBoundingWord.RightWord != null)
                                    {
                                        Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");


                                        if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.RightWord.strWord))
                                        {

                                            bWord = Module1.RoundOff(Module1.ObjBoundingWord.RightWord.strWord);
                                            if (bWord == TotalWeight && bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                            {
                                                // New added
                                                Module1.ObjBoundingWord.RightWord.strWord = bWord;
                                                // New added
                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                                {

                                                    PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                }
                                            }
                                            if (PossibleWords.Count == 1)
                                            {
                                                if (!string.IsNullOrEmpty(TotalWeight))
                                                {
                                                    if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                    {
                                                        PossibleWords[0].Flag = "1";
                                                        PossibleWords[0].Remarks = "2F";
                                                    }
                                                    else
                                                    {
                                                        PossibleWords[0].Flag = "0";
                                                        PossibleWords[0].Remarks = "NA";
                                                    }
                                                }
                                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                            }
                                        }
                                        else
                                        {
                                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                            if (Module1.ObjBoundingWord.BottomWords != null)
                                            {
                                                for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                                {
                                                    Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                                    //if ((Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)) || UnitFlag)
                                                    if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                                    {

                                                        bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                        if (bWord == TotalWeight && bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                                        {
                                                            // New added
                                                            Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                            // New added

                                                            if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) ==true)
                                                            {
                                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                                {
                                                                    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                }
                                                            }
                                                            ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                            ////{
                                                            ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                            ////}
                                                        }

                                                    }
                                                }
                                            }
                                            if (PossibleWords.Count == 1)
                                            {
                                                if (!string.IsNullOrEmpty(TotalWeight))
                                                {
                                                    if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                    {
                                                        PossibleWords[0].Flag = "1";
                                                        PossibleWords[0].Remarks = "2F";
                                                    }
                                                    else
                                                    {
                                                        PossibleWords[0].Flag = "0";
                                                        PossibleWords[0].Remarks = "NA";
                                                    }
                                                }
                                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                    if (Module1.ObjBoundingWord.BottomWords != null)
                                    {
                                        for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                        {
                                            Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                            if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                            {

                                                bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                if (bWord == TotalWeight && bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                                {
                                                    // New added
                                                    Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                    // New added

                                                    if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                    {
                                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                        {
                                                            PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                        }
                                                    }
                                                    ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                    ////{
                                                    ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                    ////}
                                                }

                                            }
                                        }
                                    }
                                    if (PossibleWords.Count == 1)
                                    {
                                        if (!string.IsNullOrEmpty(TotalWeight) && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                        {
                                            PossibleWords[0].Flag = "1";
                                            PossibleWords[0].Remarks = "2F";
                                        }
                                    }

                                    if (PossibleWords.Count > 0)
                                    {
                                        returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                    }

                                }
                                if (PossibleWords.Count == 0)
                                {
                                    //Search keyword to right
                                    DbvalueToKeywordsearch(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalPieces");
                                    if (PossibleWords.Count > 1)
                                    {
                                        CompairConfidence(PossibleWords);
                                    }
                                    if (PossibleWords.Count == 1)
                                    {
                                        if (!string.IsNullOrEmpty(TotalWeight))
                                        {
                                            if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                            {
                                                PossibleWords[0].Flag = "1";
                                                PossibleWords[0].Remarks = "2F";
                                            }
                                            else
                                            {
                                                PossibleWords[0].Flag = "0";
                                                PossibleWords[0].Remarks = "NA";
                                            }
                                        }
                                        returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                    }
                                   
                                }
                                if (PossibleWords.Count == 0)
                                {
                                    SearchDataNextTwoLine(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalPieces");
                                    if (PossibleWords.Count > 1)
                                    {
                                        CompairConfidence(PossibleWords);
                                    }
                                    if (PossibleWords.Count == 1)
                                    {
                                        if (!string.IsNullOrEmpty(TotalWeight))
                                        {
                                            if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                            {
                                                PossibleWords[0].Flag = "1";
                                                PossibleWords[0].Remarks = "2F";
                                            }
                                            else
                                            {
                                                PossibleWords[0].Flag = "0";
                                                PossibleWords[0].Remarks = "NA";
                                            }
                                        }
                                        returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                    }
                                }
                                if (PossibleWords.Count == 0)
                                {
                                    cWord = StringToCNC(intCurrPageNumber, TotalWeight);
                                    PossibleWords.Add(cWord);
                                    PossibleWords[0].Flag = "0";
                                    PossibleWords[0].Remarks = "NA";
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                }
                            }
                            if (PossibleWords.Count == 0)
                            {
                                returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                            }

                        }
                        else
                        {
                            if (!string.IsNullOrEmpty(TotalWeight))
                            {
                                if (foundRows.Length > 1)
                                {
                                    foreach (DataRow drDataRow in foundRows)
                                    {
                                        keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                                        pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                                        lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                                        wordno = Convert.ToInt32(drDataRow.ItemArray[3]);

                                        string[] splitKeword = keywordFound.Split(' ');


                                        CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);
                                        if (splitKeword.Length > 1)
                                        {
                                            Newwordno = wordno + splitKeword.Length - 1;
                                        }
                                        else
                                        {
                                            Newwordno = wordno;
                                        }

                                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                                        if (Module1.ObjBoundingWord.RightWord != null)
                                        {
                                            if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "&" || Module1.ObjBoundingWord.RightWord.strWord == "$")
                                            {
                                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                            }
                                            if (Module1.ObjBoundingWord.RightWord != null)
                                            {
                                                Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");

                                                
                                                if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || UnitFlag)
                                                {
                                                   
                                                    bWord = Module1.RoundOff(Module1.ObjBoundingWord.RightWord.strWord);
                                                    
                                                    if (bWord == TotalWeight && bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                                    {
                                                        //New added
                                                        Module1.ObjBoundingWord.RightWord.strWord = bWord;
                                                        // New added
                                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                                        {
                                                            PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                        }
                                                    }

                                                }
                                                else
                                                {
                                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                                    if (Module1.ObjBoundingWord.BottomWords != null)
                                                    {
                                                        for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                                        {
                                                            Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                                            
                                                            if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                                            {

                                                                bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                                if (bWord == TotalWeight && bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                                                {
                                                                    //New added
                                                                    Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                                    // New added

                                                                    if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                                    {
                                                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                                        {
                                                                            PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                        }
                                                                    }
                                                                    ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                                    ////{
                                                                    ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                    ////}
                                                                }

                                                            }
                                                        }
                                                    }
                                                }

                                            }
                                        }
                                        else
                                        {
                                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                            if (Module1.ObjBoundingWord.BottomWords != null)
                                            {
                                                for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                                {
                                                    Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");

                                                    if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                                    {
                                                        bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                        if (bWord == TotalWeight && bWord.Length <= 5 && Convert.ToInt32(bWord) > 0)
                                                        {
                                                            //New added
                                                            Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                            // New added
                                                           

                                                            if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                            {
                                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                                {
                                                                    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                }
                                                            }
                                                            ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                            ////{
                                                            ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                            ////}
                                                        }

                                                    }
                                                }
                                            }
                                        }

                                    }
                                    if (PossibleWords.Count == 0)
                                    {
                                        DbvalueToKeywordsearch(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalPieces");
                                        if (PossibleWords.Count > 1)
                                        {
                                            CompairConfidence(PossibleWords);
                                        }
                                        if (PossibleWords.Count == 1)
                                        {
                                            if (!string.IsNullOrEmpty(TotalWeight))
                                            {
                                               
                                                if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                {
                                                    PossibleWords[0].Flag = "1";
                                                    PossibleWords[0].Remarks = "2F";
                                                }
                                                else
                                                {
                                                    PossibleWords[0].Flag = "0";
                                                    PossibleWords[0].Remarks = "NA";
                                                }
                                            }
                                        }


                                    }
                                    if (PossibleWords.Count == 0)
                                    {
                                        SearchDataNextTwoLine(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalPieces");
                                        if (PossibleWords.Count > 1)
                                        {
                                            CompairConfidence(PossibleWords);
                                        }
                                        if (PossibleWords.Count == 1)
                                        {
                                            if (!string.IsNullOrEmpty(TotalWeight))
                                            {
                                                if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                {
                                                    PossibleWords[0].Flag = "1";
                                                    PossibleWords[0].Remarks = "2F";
                                                }
                                                else
                                                {
                                                    PossibleWords[0].Flag = "0";
                                                    PossibleWords[0].Remarks = "NA";
                                                }
                                            }
                                        }
                                    }
                                }
                                if (PossibleWords.Count == 0)
                                {
                                    cWord = StringToCNC(intCurrPageNumber, TotalWeight);
                                    PossibleWords.Add(cWord);
                                    PossibleWords[0].Flag = "0";
                                    PossibleWords[0].Remarks = "NA";
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                                }
                            }


                            if (PossibleWords.Count >= 0)
                            {
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 1);
                            }

                        }

                    }

                    // end of the F3
                }
            }
            catch (Exception)
            {
                // throw;
            }
            finally
            {
                //ds.Clear();
                ds.Dispose();
                ds = null;
                //dsKeys.Clear();
                dsKeys.Dispose();
                dsKeys = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                foundRows = null;
                CombinedKeyword = null;
                sKeyword = null;
                objCNC.Dispose();
                cWord = null;
                TotalWeight = null;
                bWord = null;
                conflevel = null;
                PossibleWords = null;
            }
            return returnZones;
        }
        #endregion

        #region Total Weight
        private RetStructF3 F3_AutoLocate_TotalWeight(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {


            //DataSet ds = new DataSet();
            DataSet ds = Module1.func_GetKeywords(ObjMetaData, intCurrPageNumber);

            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            //clsCNCSBR objclsCNCSBR = new clsCNCSBR();
            DataSet dsKeys = new DataSet();
            DataTable dt = new DataTable();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;
            string TotalWeight = "";
            string bWord = "";
            Dictionary<bool, bool> Dic_Unitflag = new Dictionary<bool, bool>();
            bool UnitFlag = false;
            bool KgUnit = false;

            System.Data.DataRow[] foundRows;
            ClsCNC objCNC = new ClsCNC();
            clsCnCWord cWord = new clsCnCWord();
            clsCnCWord CombinedKeyword = new clsCnCWord();

        Line1:
            ;

            if (iRecursiveCallCount == 0)
            {
                // -------------------------------------------------Weight KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                sKeyword = "'Total Amount', 'Gross Weight','Net Weight','Grand Total','Total Weight','Gross WT','GRANDTOTAL','NetWeight','GrossWeight','TotalAmount','Weight Total','TOTAL WT','PAGE SUBTOTAL','Weight','Total','Total Gross','Gross'";

            }
            // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            else
            {
                sKeyword = "''";
            }

            try
            {
                returnZones.Status = "F";
                returnZones.ManualConfirmation = "Y";
                returnZones.Flag = "Y";

                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_TotalWeight"))
                {
                    TotalWeight = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_TotalWeight"];
                }

                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_keyword"];
                }
                if (dsKeys != null)
                {
                    dt = dsKeys.Tables[0];
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();
                    string keywordFound = "";
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int Newwordno = 0;

                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);


                        string[] splitKeword = keywordFound.Split(' ');
                        if (splitKeword.Length > 1)
                        {
                            Newwordno = wordno + splitKeword.Length - 1;
                        }
                        else
                        {
                            Newwordno = wordno;
                        }

                        CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                        if (!string.IsNullOrEmpty(TotalWeight))
                        {
                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);
                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                                Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "");
                            Label1:
                                if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "&" || Module1.ObjBoundingWord.RightWord.strWord == "$" || Module1.ObjBoundingWord.RightWord.strWord == ">>" || Module1.ObjBoundingWord.RightWord.strWord == "" || Module1.ObjBoundingWord.RightWord.strWord == "***" || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "LBS." || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "LBS" || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "LB")
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                }


                                if (Module1.ObjBoundingWord.RightWord != null)
                                {
                                    Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                    if (Module1.ObjBoundingWord.RightWord.strWord == ":")
                                    {
                                        goto Label1;
                                    }
                                    Dic_Unitflag = CheckUnitForWeight(objCNC, ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.RightWord);
                                    foreach (var item in Dic_Unitflag)
                                    {
                                        UnitFlag = item.Key;
                                        KgUnit = item.Value;
                                    }
                                    if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || UnitFlag)
                                    //if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.RightWord.strWord))
                                    {
                                        
                                        bWord = Module1.RoundOff(Module1.ObjBoundingWord.RightWord.strWord);
                                        
                                        try
                                        {
                                            if (KgUnit)
                                            {
                                                bWord = Convert.ToString((Convert.ToInt32(bWord) * 2.205));
                                                bWord = Module1.RoundOff(bWord);
                                                Module1.ObjBoundingWord.RightWord = ModifiedStringToCNC(Module1.ObjBoundingWord.RightWord, bWord); ;
                                            }
                                        }
                                        catch
                                        {

                                            bWord = Module1.ObjBoundingWord.RightWord.strWord;
                                        }
                                        if (bWord == TotalWeight && bWord.Length <= 7 && Convert.ToInt32(bWord) > 0)
                                        {
                                            // new added
                                            Module1.ObjBoundingWord.RightWord.strWord = bWord;
                                            // new added
                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                            {

                                                PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                            }
                                        }
                                        if (PossibleWords.Count == 1)
                                        {
                                            if (!string.IsNullOrEmpty(TotalWeight))
                                            {
                                                ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                ////{
                                                ////    PossibleWords[0].Flag = "1";
                                                ////    PossibleWords[0].Remarks = "2F";
                                                ////}
                                                ////else
                                                ////{
                                                    PossibleWords[0].Flag = "0";
                                                    PossibleWords[0].Remarks = "NA";
                                                ////}
                                            }
                                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                                        }


                                    }
                                    else
                                    {
                                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                        if (Module1.ObjBoundingWord.BottomWords != null)
                                        {
                                            for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                            {
                                                Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                                Dic_Unitflag = CheckUnitForWeight(objCNC, ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j]);
                                                foreach (var item in Dic_Unitflag)
                                                {
                                                    UnitFlag = item.Key;
                                                    KgUnit = item.Value;
                                                }
                                                if ((Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)) || UnitFlag)
                                                {
                                                    
                                                    bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                    
                                                    try
                                                    {
                                                        if (KgUnit)
                                                        {
                                                            bWord = Convert.ToString((Convert.ToInt32(bWord) * 2.205));
                                                            bWord = Module1.RoundOff(bWord);
                                                            Module1.ObjBoundingWord.BottomWords[j] = ModifiedStringToCNC(Module1.ObjBoundingWord.BottomWords[j], bWord); ;
                                                        }
                                                    }
                                                    catch
                                                    {

                                                        bWord = Module1.ObjBoundingWord.BottomWords[j].strWord;
                                                    }

                                                    if (bWord == TotalWeight && bWord.Length <= 7 && Convert.ToInt32(bWord) > 0)
                                                    {
                                                        // new added
                                                        Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                        // new added
                                                        if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                        {
                                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                            {
                                                                PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                            }
                                                        }
                                                        ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                        ////{
                                                        ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                        ////}
                                                    }

                                                }
                                            }
                                        }
                                        if (PossibleWords.Count == 1)
                                        {
                                            if (!string.IsNullOrEmpty(TotalWeight))
                                            {
                                                ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                ////{
                                                ////    PossibleWords[0].Flag = "1";
                                                ////    PossibleWords[0].Remarks = "2F";
                                                ////}
                                                ////else
                                                ////{
                                                    PossibleWords[0].Flag = "0";
                                                    PossibleWords[0].Remarks = "NA";
                                                ////}
                                            }
                                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                                        }

                                    }
                                }
                            }
                            else
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                if (Module1.ObjBoundingWord.BottomWords != null)
                                {
                                    for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                    {
                                        Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                        Dic_Unitflag = CheckUnitForWeight(objCNC, ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j]);
                                        foreach (var item in Dic_Unitflag)
                                        {
                                            UnitFlag = item.Key;
                                            KgUnit = item.Value;
                                        }
                                        if ((Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)) || UnitFlag)
                                        {

                                            bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                            
                                            try
                                            {
                                                if (KgUnit)
                                                {
                                                    bWord = Convert.ToString((Convert.ToInt32(bWord) * 2.205));
                                                    bWord = Module1.RoundOff(bWord);
                                                    Module1.ObjBoundingWord.BottomWords[j] = ModifiedStringToCNC(Module1.ObjBoundingWord.BottomWords[j], bWord); ;
                                                }
                                            }
                                            catch
                                            {

                                                bWord = Module1.ObjBoundingWord.BottomWords[j].strWord;
                                            }
                                            if (bWord == TotalWeight && bWord.Length <= 7 && Convert.ToInt32(bWord) > 0)
                                            {
                                                // new added
                                                Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                // new added

                                                if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                {
                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                    {
                                                        PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                    }
                                                }
                                                ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                ////{
                                                ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                ////}
                                            }

                                        }
                                    }
                                }
                                if (PossibleWords.Count == 1)
                                {
                                    if (!string.IsNullOrEmpty(TotalWeight))
                                    {
                                        ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                        ////{
                                        ////    PossibleWords[0].Flag = "1";
                                        ////    PossibleWords[0].Remarks = "2F";
                                        ////}
                                        ////else
                                        ////{
                                            PossibleWords[0].Flag = "0";
                                            PossibleWords[0].Remarks = "NA";
                                        ////}
                                    }
                                }

                                if (PossibleWords.Count > 0)
                                {
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                                }

                            }
                            if (PossibleWords.Count == 0)
                            {
                                //Search keyword to right
                                DbvalueToKeywordsearch(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalWeight");
                                if (PossibleWords.Count > 1)
                                {
                                    CompairConfidence(PossibleWords);
                                }
                                if (PossibleWords.Count == 1)
                                {
                                    if (!string.IsNullOrEmpty(TotalWeight))
                                    {
                                        ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                        ////{
                                        ////    PossibleWords[0].Flag = "1";
                                        ////    PossibleWords[0].Remarks = "2F";
                                        ////}
                                        ////else
                                        ////{
                                            PossibleWords[0].Flag = "0";
                                            PossibleWords[0].Remarks = "NA";
                                       //// }
                                    }
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                                }
                            }
                            if (PossibleWords.Count == 0)
                            {
                                SearchDataNextTwoLine(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalWeight");
                                if (PossibleWords.Count > 1)
                                {
                                    CompairConfidence(PossibleWords);
                                }
                                if (PossibleWords.Count == 1)
                                {
                                    if (!string.IsNullOrEmpty(TotalWeight))
                                    {
                                        ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                        ////{
                                        ////    PossibleWords[0].Flag = "1";
                                        ////    PossibleWords[0].Remarks = "2F";
                                        ////}
                                        ////else
                                        ////{
                                            PossibleWords[0].Flag = "0";
                                            PossibleWords[0].Remarks = "NA";
                                        ////}
                                    }
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                                }
                            }
                            if (PossibleWords.Count == 0)
                            {
                                //Search keyword to right
                                KeywordTODbvaluesearch(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalWeight");
                                if (PossibleWords.Count > 1)
                                {
                                    CompairConfidence(PossibleWords);
                                }
                                if (PossibleWords.Count == 1)
                                {
                                    if (!string.IsNullOrEmpty(TotalWeight))
                                    {
                                        ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                        ////{
                                        ////    PossibleWords[0].Flag = "1";
                                        ////    PossibleWords[0].Remarks = "2F";
                                        ////}
                                        ////else
                                        ////{
                                            PossibleWords[0].Flag = "0";
                                            PossibleWords[0].Remarks = "NA";
                                       //// }
                                    }
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                                }
                            }
                            if (PossibleWords.Count == 0)
                            {
                                cWord = StringToCNC(intCurrPageNumber, TotalWeight);
                                PossibleWords.Add(cWord);
                                PossibleWords[0].Flag = "0";
                                PossibleWords[0].Remarks = "NA";
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                            }
                        }
                        if (PossibleWords.Count == 0)
                        {
                            returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                        }

                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(TotalWeight))
                        {
                            if (foundRows.Length > 1)
                            {
                                foreach (DataRow drDataRow in foundRows)
                                {
                                    keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                                    pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                                    lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                                    wordno = Convert.ToInt32(drDataRow.ItemArray[3]);

                                    string[] splitKeword = keywordFound.Split(' ');


                                    CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);
                                    if (splitKeword.Length > 1)
                                    {
                                        Newwordno = wordno + splitKeword.Length - 1;
                                    }
                                    else
                                    {
                                        Newwordno = wordno;
                                    }

                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                                    if (Module1.ObjBoundingWord.RightWord != null)
                                    {
                                        Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "");
                                    Label1:
                                        if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "&" || Module1.ObjBoundingWord.RightWord.strWord == "$" || Module1.ObjBoundingWord.RightWord.strWord == ">>" || Module1.ObjBoundingWord.RightWord.strWord == "" || Module1.ObjBoundingWord.RightWord.strWord == "***" || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "LBS." || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "LBS" || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "LB")
                                        {
                                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                        }

                                        if (Module1.ObjBoundingWord.RightWord != null)
                                        {
                                            Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                            if (Module1.ObjBoundingWord.RightWord.strWord == ":")
                                            {
                                                goto Label1;
                                            }
                                            Dic_Unitflag = CheckUnitForWeight(objCNC, ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.RightWord);
                                            foreach (var item in Dic_Unitflag)
                                            {
                                                UnitFlag = item.Key;
                                                KgUnit = item.Value;
                                            }
                                            if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || UnitFlag)
                                            {
                                                
                                                bWord = Module1.RoundOff(Module1.ObjBoundingWord.RightWord.strWord);
                                                
                                                try
                                                {
                                                    if (KgUnit)
                                                    {
                                                        bWord = Convert.ToString((Convert.ToInt32(bWord) * 2.205));
                                                        bWord = Module1.RoundOff(bWord);
                                                        Module1.ObjBoundingWord.RightWord = ModifiedStringToCNC(Module1.ObjBoundingWord.RightWord, bWord); ;
                                                    }
                                                }
                                                catch
                                                {

                                                    bWord = Module1.ObjBoundingWord.RightWord.strWord;
                                                }
                                                if (bWord == TotalWeight && bWord.Length <= 7 && Convert.ToInt32(bWord) > 0)
                                                {
                                                    // new added
                                                    Module1.ObjBoundingWord.RightWord.strWord = bWord;
                                                    // new added
                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                                    {
                                                        PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                    }
                                                }

                                            }
                                            else
                                            {
                                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                                if (Module1.ObjBoundingWord.BottomWords != null)
                                                {
                                                    for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                                    {
                                                        Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                                        Dic_Unitflag = CheckUnitForWeight(objCNC, ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j]);
                                                        foreach (var item in Dic_Unitflag)
                                                        {
                                                            UnitFlag = item.Key;
                                                            KgUnit = item.Value;
                                                        }
                                                        if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || UnitFlag)
                                                        {
                                                            bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                            
                                                            try
                                                            {
                                                                if (KgUnit)
                                                                {
                                                                    bWord = Convert.ToString((Convert.ToInt32(bWord) * 2.205));
                                                                    bWord = Module1.RoundOff(bWord);
                                                                    Module1.ObjBoundingWord.BottomWords[j] = ModifiedStringToCNC(Module1.ObjBoundingWord.BottomWords[j], bWord);
                                                                }
                                                            }
                                                            catch
                                                            {

                                                                bWord = Module1.ObjBoundingWord.BottomWords[j].strWord;
                                                            }
                                                            if (bWord == TotalWeight && bWord.Length <= 7 && Convert.ToInt32(bWord) > 0)
                                                            {
                                                                // new added
                                                                Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                                // new added

                                                                if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                                {
                                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                                    {
                                                                        PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                    }
                                                                }
                                                                ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                                ////{
                                                                ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                ////}
                                                            }

                                                        }
                                                    }
                                                }
                                            }

                                        }
                                    }
                                    else
                                    {
                                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                        if (Module1.ObjBoundingWord.BottomWords != null)
                                        {
                                            for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                            {
                                                Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                                Dic_Unitflag = CheckUnitForWeight(objCNC, ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j]);
                                                foreach (var item in Dic_Unitflag)
                                                {
                                                    UnitFlag = item.Key;
                                                    KgUnit = item.Value;
                                                }
                                                if (Module1.regexTW.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || UnitFlag)
                                                {
                                                    bWord = Module1.RoundOff(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                    
                                                    try
                                                    {
                                                        if (KgUnit)
                                                        {
                                                            bWord = Convert.ToString((Convert.ToInt32(bWord) * 2.205));
                                                            bWord = Module1.RoundOff(bWord);
                                                            Module1.ObjBoundingWord.BottomWords[j] = ModifiedStringToCNC(Module1.ObjBoundingWord.BottomWords[j], bWord);
                                                        }
                                                    }
                                                    catch
                                                    {

                                                        bWord = Module1.ObjBoundingWord.BottomWords[j].strWord;
                                                    }

                                                    if (bWord == TotalWeight && bWord.Length <= 7 && Convert.ToInt32(bWord) > 0)
                                                    {
                                                        // new added
                                                        Module1.ObjBoundingWord.BottomWords[j].strWord = bWord;
                                                        // new added

                                                        if (Keyword_value_Heightdiff(CombinedKeyword, Module1.ObjBoundingWord.BottomWords[j]) == true)
                                                        {
                                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                            {
                                                                PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                            }
                                                        }
                                                        ////if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                        ////{
                                                        ////    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                        ////}
                                                    }

                                                }
                                            }
                                        }
                                    }

                                }
                                if (PossibleWords.Count == 0)
                                {
                                    DbvalueToKeywordsearch(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalWeight");
                                    if (PossibleWords.Count > 1)
                                    {
                                        CompairConfidence(PossibleWords);
                                    }
                                    if (PossibleWords.Count == 1)
                                    {
                                        if (!string.IsNullOrEmpty(TotalWeight))
                                        {
                                            ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                            ////{
                                            ////    PossibleWords[0].Flag = "1";
                                            ////    PossibleWords[0].Remarks = "2F";
                                            ////}
                                            ////else
                                            ////{
                                                PossibleWords[0].Flag = "0";
                                                PossibleWords[0].Remarks = "NA";
                                           //// }
                                        }
                                    }


                                }
                                if (PossibleWords.Count == 0)
                                {
                                    SearchDataNextTwoLine(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalWeight");
                                    if (PossibleWords.Count > 1)
                                    {
                                        CompairConfidence(PossibleWords);
                                    }
                                    if (PossibleWords.Count == 1)
                                    {
                                        if (!string.IsNullOrEmpty(TotalWeight))
                                        {
                                            ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                            ////{
                                            ////    PossibleWords[0].Flag = "1";
                                            ////    PossibleWords[0].Remarks = "2F";
                                            ////}
                                            ////else
                                            ////{
                                                PossibleWords[0].Flag = "0";
                                                PossibleWords[0].Remarks = "NA";
                                            ////}
                                        }
                                    }
                                }
                                if (PossibleWords.Count == 0)
                                {
                                    //Search keyword to right
                                    KeywordTODbvaluesearch(TotalWeight, foundRows, sKeyword, ObjMetaData, intCurrPageNumber, PossibleWords, "TotalWeight");
                                    if (PossibleWords.Count > 1)
                                    {
                                        CompairConfidence(PossibleWords);
                                    }
                                    if (PossibleWords.Count == 1)
                                    {
                                        if (!string.IsNullOrEmpty(TotalWeight))
                                        {
                                            ////if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                            ////{
                                            ////    PossibleWords[0].Flag = "1";
                                            ////    PossibleWords[0].Remarks = "2F";
                                            ////}
                                            ////else
                                            ////{
                                                PossibleWords[0].Flag = "0";
                                                PossibleWords[0].Remarks = "NA";
                                            ////}
                                        }
                                    }
                                    
                                }
                            }
                            if (PossibleWords.Count == 0)
                            {
                                cWord = StringToCNC(intCurrPageNumber, TotalWeight);
                                PossibleWords.Add(cWord);
                                PossibleWords[0].Flag = "0";
                                PossibleWords[0].Remarks = "NA";
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                            }
                        }

                        if (PossibleWords.Count >= 0)
                        {
                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 2);
                        }

                    }

                    // end of the F3
                }
            }
            catch (Exception)
            {
                // throw;
            }
            finally
            {
                //ds.Clear();
                ds.Dispose();
                ds = null;
                //dsKeys.Clear();
                dsKeys.Dispose();
                dsKeys = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                foundRows = null;
                CombinedKeyword = null;
                sKeyword = null;
                objCNC.Dispose();
                cWord = null;
                TotalWeight = null;
                bWord = null;
                Dic_Unitflag = null;
                conflevel = null;
                PossibleWords = null;
            }
            return returnZones;
        }
        #endregion

        #region COD
        private RetStructF3 F3_AutoLocate_COD(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {

            //DataSet ds = new DataSet();
            DataSet ds = Module1.func_GetKeywords(ObjMetaData, intCurrPageNumber);
            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            //clsCNCSBR objclsCNCSBR = new clsCNCSBR();
            DataSet dsKeys = new DataSet();
            DataTable dt = new DataTable();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;

            System.Data.DataRow[] foundRows;
            ClsCNC objCNC = new ClsCNC();
            clsCnCWord cWord = new clsCnCWord();
            clsCnCWord CombinedKeyword = new clsCnCWord();

        Line1:
            ;

            if (iRecursiveCallCount == 0)
            {
                // -------------------------------------------------COD KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                sKeyword = " 'COD AMOUNT', 'COD AMT', 'COD'";
            }
            // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            else
            {
                sKeyword = "''";
            }

            try
            {
                returnZones.Status = "F";
                returnZones.ManualConfirmation = "Y";
                returnZones.Flag = "Y";
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_keyword"];
                }
                if (dsKeys != null)
                {
                    dt = dsKeys.Tables[0];
                    ////System.Data.DataRow[] foundRows;
                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
                    // foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber);
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();
                    string keywordFound = "";
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int Newwordno = 0;


                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);


                        string[] splitKeword = keywordFound.Split(' ');
                        if (splitKeword.Length > 1)
                        {
                            Newwordno = wordno + splitKeword.Length - 1;
                        }
                        else
                        {
                            Newwordno = wordno;
                        }

                        CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                        if (Module1.ObjBoundingWord.RightWord != null)
                        {
                        Label:
                            
                            if (!Module1.ObjBoundingWord.RightWord.strWord.Any(char.IsLetter) && !Module1.regex.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) && Module1.SpecialCharacters.Any(Module1.ObjBoundingWord.RightWord.strWord.Contains))
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                            }
                            
                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                                if (!Module1.ObjBoundingWord.RightWord.strWord.Any(char.IsLetter) && !Module1.regex.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) && Module1.SpecialCharacters.Any(Module1.ObjBoundingWord.RightWord.strWord.Contains))
                                {
                                    goto Label;
                                }
                                
                                if (Module1.regex.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.ObjBoundingWord.RightWord.strWord == "0")
                                {
                                    //new added
                                    Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Replace(",","").Replace(".", "");
                                    // New added

                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                    {
                                        PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                    }

                                    if (PossibleWords.Count > 0)
                                    {
                                        returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 3);
                                    }
                                }
                            }
                            else
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                if (Module1.ObjBoundingWord.BottomWords != null)
                                {
                                    for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                    {
                                        if ((Module1.regex.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) && (!(Module1.ObjBoundingWord.BottomWords[j].strWord.Contains(":")))))
                                        {
                                            //new added
                                            Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace(",", "").Replace(".", "");
                                            // New added

                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                            {
                                                PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                            }
                                        }
                                    }
                                }
                                if (PossibleWords.Count > 0)
                                {
                                    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 3);
                                }

                            }

                        }
                        else
                        {
                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                            if (Module1.ObjBoundingWord.BottomWords != null)
                            {
                                for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                {
                                    if ((Module1.regex.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) && (!(Module1.ObjBoundingWord.BottomWords[j].strWord.Contains(":")))))
                                    {
                                        //new added
                                        Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace(",", "").Replace(".", "");
                                        // New added

                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                        {
                                            PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                        }
                                    }
                                }
                            }

                            if (PossibleWords.Count > 0)
                            {
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 3);
                            }

                        }
                        if (PossibleWords.Count == 0)
                        {
                            returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);

                        }
                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);

                            string[] splitKeword = keywordFound.Split(' ');


                            CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);
                            if (splitKeword.Length > 1)
                            {
                                Newwordno = wordno + splitKeword.Length - 1;
                            }
                            else
                            {
                                Newwordno = wordno;
                            }


                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                            Label:
                                if (!Module1.ObjBoundingWord.RightWord.strWord.Any(char.IsLetter) && !Module1.regex.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) && Module1.SpecialCharacters.Any(Module1.ObjBoundingWord.RightWord.strWord.Contains))
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                }
                                
                                if (Module1.ObjBoundingWord.RightWord != null)
                                {
                                    
                                    if (!Module1.ObjBoundingWord.RightWord.strWord.Any(char.IsLetter) && !Module1.regex.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) && Module1.SpecialCharacters.Any(Module1.ObjBoundingWord.RightWord.strWord.Contains))
                                    {
                                        goto Label;
                                    }
                                    if (Module1.regex.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.ObjBoundingWord.RightWord.strWord == "0")
                                    {
                                        //new added
                                        Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Replace(",", "").Replace(".", "");
                                        // New added
                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                        {
                                            PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                        }
                                    }

                                }
                                else
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                    if (Module1.ObjBoundingWord.BottomWords != null)
                                    {
                                        for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                        {
                                            if (Module1.regex.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                            {
                                                //new added
                                                Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace(",", "").Replace(".", "");
                                                // New added
                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                                {
                                                    PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                }
                                            }
                                        }
                                    }
                                }

                            }
                            else
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                if (Module1.ObjBoundingWord.BottomWords != null)
                                {
                                    for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                    {
                                        if (Module1.regex.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                        {
                                            //new added
                                            Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace(",", "").Replace(".", "");
                                            // New added

                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                            {
                                                PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                            }
                                        }
                                    }
                                }
                            }
                        } // end of for 


                        if (PossibleWords.Count >= 0)
                        {
                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 3);
                        }

                    }
                    // end of the F3
                }
            }
            catch (Exception)
            {
                // throw;
            }
            finally
            {
                //ds.Clear();
                ds.Dispose();
                ds = null;
                //dsKeys.Clear();
                dsKeys.Dispose();
                dsKeys = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                foundRows = null;
                CombinedKeyword = null;
                sKeyword = null;
                objCNC.Dispose();
                cWord = null;
                conflevel = null;
                PossibleWords = null;
            }

            return returnZones;
        }
        #endregion

        #region Quote No
        private RetStructF3 F3_AutoLocate_QUOTENO(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {

            //DataSet ds = new DataSet();
            DataSet ds = Module1.func_GetKeywords(ObjMetaData, intCurrPageNumber);

            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            DataSet dsKeys = new DataSet();
            DataTable dt = new DataTable();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;
            bool AllAlphabets;

            System.Data.DataRow[] foundRows;
            ClsCNC objCNC = new ClsCNC();
            clsCnCWord cWord = new clsCnCWord();
            clsCnCWord CombinedKeyword = new clsCnCWord();
        Line1:
            ;

            if (iRecursiveCallCount == 0)
            {
                // -------------------------------------------------QUOTE NO KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                sKeyword = "'Quote', 'Quote No','Quote Number','QuoteID','Quote ID','Quote ID Number','QuoteNumber','Quote Reference ID','Q#'"; //,'Quot e'
            }
            // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            else
            {
                sKeyword = "''";
            }

            try
            {
                returnZones.Status = "F";
                returnZones.ManualConfirmation = "Y";
                returnZones.Flag = "Y";
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_keyword"];
                }
                if (dsKeys != null)
                {
                    dt = dsKeys.Tables[0];
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[2]);
                    foundRows = orderedRows.ToArray();
                    string keywordFound = "";
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int Newwordno = 0;


                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);


                        string[] splitKeword = keywordFound.Split(' ');
                        if (splitKeword.Length > 1)
                        {
                            Newwordno = wordno + splitKeword.Length - 1;
                        }
                        else
                        {
                            Newwordno = wordno;
                        }

                        CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                        if (Module1.ObjBoundingWord.RightWord != null)
                        {
                        Lable:
                            if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "·" || Module1.ObjBoundingWord.RightWord.strWord == "." || (Module1.ObjBoundingWord.RightWord.strWord.Length <= 2))
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                            }
                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                                if (Module1.ObjBoundingWord.RightWord.strWord.Length <= 2)
                                {
                                    goto Lable;
                                }
                                string oldStrVal = Module1.ObjBoundingWord.RightWord.strWord;
                                if (oldStrVal.Contains("#"))
                                {
                                    int indxofhash = oldStrVal.IndexOf("#");
                                    Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Substring(indxofhash);
                                }
                                Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Replace("#", "").Replace(":", "").Replace(";", "").Replace(",", "").Replace(".", "").Replace("'", "").Replace("`", "").Replace("~", "").Replace("!", "").Replace("#", "").Replace("$", "").Replace("/", "").Replace("*", "");
                                //Module1.ObjBoundingWord.RightWord.strWord = Regex.Replace(Module1.ObjBoundingWord.RightWord.strWord, @"[^0-9a-zA-Z]+", "");
                                string chkAllapha = Regex.Replace(Module1.ObjBoundingWord.RightWord.strWord, @"[^0-9a-zA-Z]+", "");
                                AllAlphabets = chkAllapha.All(char.IsLetter);
                                //AllAlphabets = Module1.ObjBoundingWord.RightWord.strWord.All(char.IsLetter);

                                if (AllAlphabets == false && Module1.ObjBoundingWord.RightWord.strWord.Length >= 7 && Module1.ObjBoundingWord.RightWord.strWord.Length <= 12)
                                {
                                    if ((Module1.reg1.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.reg2.IsMatch(Module1.ObjBoundingWord.RightWord.strWord))
                                        && !Char.IsLetter(Module1.ObjBoundingWord.RightWord.strWord[Module1.ObjBoundingWord.RightWord.strWord.Length - 1]))
                                    {
                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                        {
                                            PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                            ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1" : "0";
                                            ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1F" : "NA";
                                            if (HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) == true)
                                            {
                                                PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                            }
                                            else
                                            {
                                                PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                            }
                                            ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) ? "1" : "0";
                                            ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) ? "1F" : "NA";
                                            // new added
                                            if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                            {
                                                PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                            }
                                            //new added
                                        }
                                    }
                                    else if (Module1.ObjBoundingWord.RightWord.strWord.Contains('-'))
                                    {
                                        Module1.ObjBoundingWord.RightWord = vilidateQuoteNo(Module1.ObjBoundingWord.RightWord);
                                        if (Module1.ObjBoundingWord.RightWord != null)
                                        {
                                            if (Module1.ObjBoundingWord.RightWord.strWord != null)
                                            {
                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                                {
                                                    PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                    ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1" : "0";
                                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1F" : "NA";

                                                    if (HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) == true)
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                    }
                                                    else
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],false) ? "1" : "0";
                                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) ? "1F" : "NA";
                                                    // new added
                                                    if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    //new added
                                                }
                                            }
                                        }
                                    }
                                    // new added 31082021
                                    else if (ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.Length > 7 && ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Substring(0, 5) == "QUOTE")
                                    {
                                        string QNoCheck = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Replace("#", "").Replace(":", "").Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Trim();

                                        if ((Module1.reg1.IsMatch(QNoCheck) || Module1.reg2.IsMatch(QNoCheck)) && !Char.IsLetter(QNoCheck[QNoCheck.Length - 1]))
                                        {
                                            Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);

                                            // new added 16062021
                                            ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                            // new added 16062021

                                            if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                            {
                                                PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                                ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                                ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                                if (confidence == true)
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                }
                                                else
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                }
                                                if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                }
                                                //new added

                                            }

                                        }
                                        else if (!Char.IsLetter(QNoCheck[QNoCheck.Length - 1])&& QNoCheck.Length>7)
                                        {
                                            Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);

                                            // new added 16062021
                                            ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                            // new added 16062021

                                            if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                            {
                                                PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                                ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                                ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                                if (confidence == true)
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                }
                                                else
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                }
                                                if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                }
                                                //new added

                                            }

                                        }
                                        ////if (PossibleWords.Count > 0)
                                        ////{
                                        ////    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 4);
                                        ////}

                                    }
                                    //new added 31082021
                                    else if(Module1.ObjBoundingWord.RightWord.strWord.Length >= 7 && Module1.ObjBoundingWord.RightWord.strWord.Length <= 12)
                                    {
                                        if (Module1.ObjBoundingWord.RightWord != null)
                                        {
                                            if (Module1.ObjBoundingWord.RightWord.strWord != null)
                                            {
                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                                {
                                                    PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                    ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1" : "0";
                                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1F" : "NA";

                                                    if (HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], true) == true)
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                    }
                                                    else
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],true) ? "1" : "0";
                                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],true) ? "1F" : "NA";
                                                    // new added
                                                    if (Module1.SpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    //new added
                                                }
                                            }
                                        }
                                    }
                                    if (PossibleWords.Count > 0)
                                    {
                                        returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 4);
                                    }

                                }
                            }
                           
                        }
                        if (PossibleWords.Count == 0 && ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.Length > 7 && ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Substring(0, 5) == "QUOTE")
                        {
                            string QNoCheck = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Replace("#", "").Replace(":", "").Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Trim();

                            if ((Module1.reg1.IsMatch(QNoCheck) || Module1.reg2.IsMatch(QNoCheck)) && !Char.IsLetter(QNoCheck[QNoCheck.Length - 1]))
                            {
                                Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);
                                
                                // new added 16062021
                                ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                // new added 16062021

                                if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                {
                                    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                    PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                    PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";
                                    if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                    {
                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                    }
                                    //new added

                                }

                            }
                            else if (!Char.IsLetter(QNoCheck[QNoCheck.Length - 1]) && QNoCheck.Length > 7)
                            {
                                Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);

                                // new added 16062021
                                ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                // new added 16062021

                                if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                {
                                    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                    ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                    if (confidence == true)
                                    {
                                        PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                        PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                    }
                                    else
                                    {
                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                    }
                                    if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                    {
                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                    }
                                    //new added

                                }

                            }
                            if (PossibleWords.Count > 0)
                            {
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 4);
                            }

                        }
                        if (PossibleWords.Count == 0)
                        {
                            returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                        }
                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);

                            string[] splitKeword = keywordFound.Split(' ');


                            CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);
                            if (splitKeword.Length > 1)
                            {
                                Newwordno = wordno + splitKeword.Length - 1;
                            }
                            else
                            {
                                Newwordno = wordno;
                            }


                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                            Lable:
                                if (Module1.ObjBoundingWord.RightWord.strWord == ":" || Module1.ObjBoundingWord.RightWord.strWord == "·" || Module1.ObjBoundingWord.RightWord.strWord == "." || Module1.ObjBoundingWord.RightWord.strWord == "#" || Module1.ObjBoundingWord.RightWord.strWord.Length <= 2)
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                }
                                if (Module1.ObjBoundingWord.RightWord != null)
                                {
                                    if (Module1.ObjBoundingWord.RightWord.strWord.Length <= 2)
                                    {
                                        goto Lable;
                                    }
                                    string oldStrVal = Module1.ObjBoundingWord.RightWord.strWord;
                                    if (oldStrVal.Contains("#"))
                                    {
                                        int indxofhash = oldStrVal.IndexOf("#");
                                        Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Substring(indxofhash);
                                    }
                                    Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Replace("#", "").Replace(":", "").Replace(";", "").Replace(",", "").Replace(".", "").Replace("'", "").Replace("`", "").Replace("~", "").Replace("!", "").Replace("#", "").Replace("$", "").Replace("/", "").Replace("*", "");
                                    //Module1.ObjBoundingWord.RightWord.strWord = Regex.Replace(Module1.ObjBoundingWord.RightWord.strWord, @"[^0-9a-zA-Z]+", "");
                                    string chkAllapha = Regex.Replace(Module1.ObjBoundingWord.RightWord.strWord, @"[^0-9a-zA-Z]+", "");
                                    AllAlphabets = chkAllapha.All(char.IsLetter);
                                   // AllAlphabets = Module1.ObjBoundingWord.RightWord.strWord.All(char.IsLetter);

                                    if (AllAlphabets == false && Module1.ObjBoundingWord.RightWord.strWord.Length >= 7 && Module1.ObjBoundingWord.RightWord.strWord.Length <= 12)
                                    {
                                        if ((Module1.reg1.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.reg2.IsMatch(Module1.ObjBoundingWord.RightWord.strWord))
                                            && (!Char.IsLetter(Module1.ObjBoundingWord.RightWord.strWord[Module1.ObjBoundingWord.RightWord.strWord.Length - 1])))
                                        {
                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                            {
                                                PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1" : "0";
                                                ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1F" : "NA";
                                                if (HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) == true)
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                }
                                                else
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                }

                                                ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],false) ? "1" : "0";
                                                ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],false) ? "1F" : "NA";

                                                //new added
                                                if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                {
                                                    PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                    PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                }
                                                //new added
                                            }
                                        }
                                        else if (Module1.ObjBoundingWord.RightWord.strWord.Contains('-'))
                                        {
                                            Module1.ObjBoundingWord.RightWord = vilidateQuoteNo(Module1.ObjBoundingWord.RightWord);
                                            if (Module1.ObjBoundingWord.RightWord != null)
                                            {
                                                if (Module1.ObjBoundingWord.RightWord.strWord != null)
                                                {
                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                                    {
                                                        PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                        ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1" : "0";
                                                        ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1F" : "NA";

                                                        if (HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], false) == true)
                                                        {
                                                            PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                            PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                        }
                                                        else
                                                        {
                                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                        }
                                                        ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],false) ? "1" : "0";
                                                        ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1],false) ? "1F" : "NA";
                                                        //new added
                                                        if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                        {
                                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                        }
                                                        //new added
                                                    }
                                                }

                                            }

                                        }
                                        // new added 31082021
                                        else if (ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.Length > 7 && ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Substring(0, 5) == "QUOTE")
                                        {
                                            string QNoCheck = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Replace("#", "").Replace(":", "").Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Trim();

                                            if ((Module1.reg1.IsMatch(QNoCheck) || Module1.reg2.IsMatch(QNoCheck)) && !Char.IsLetter(QNoCheck[QNoCheck.Length - 1]))
                                            {
                                                Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);

                                                // new added 16062021
                                                ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                                // new added 16062021

                                                if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                                {
                                                    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                                    ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                                    if (confidence == true)
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                    }
                                                    else
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    //new added

                                                }

                                            }
                                            else if (!Char.IsLetter(QNoCheck[QNoCheck.Length - 1]) && QNoCheck.Length > 7)
                                            {
                                                Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);

                                                // new added 16062021
                                                ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                                // new added 16062021

                                                if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                                {
                                                    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                                    ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                                    ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                                    if (confidence == true)
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                    }
                                                    else
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                    {
                                                        PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                        PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                    }
                                                    //new added

                                                }

                                            }
                                            ////if (PossibleWords.Count > 0)
                                            ////{
                                            ////    returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 4);
                                            ////}

                                        }
                                        //new added 31082021
                                        else if (Module1.ObjBoundingWord.RightWord.strWord.Length >= 7 && Module1.ObjBoundingWord.RightWord.strWord.Length <= 12)
                                        {
                                            if (Module1.ObjBoundingWord.RightWord != null)
                                            {
                                                if (Module1.ObjBoundingWord.RightWord.strWord != null)
                                                {
                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord) && !(Module1.ObjBoundingWord.RightWord.strWord == ":"))
                                                    {
                                                        PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                        ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1" : "0";
                                                        ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence(PossibleWords[PossibleWords.Count - 1]) ? "1F" : "NA";

                                                        if (HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], true) == true)
                                                        {
                                                            PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                                            PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                                        }
                                                        else
                                                        {
                                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                        }
                                                        ////PossibleWords[PossibleWords.Count - 1].Flag = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], true) ? "1" : "0";
                                                        ////PossibleWords[PossibleWords.Count - 1].Remarks = HighestConfidence_Quote(PossibleWords[PossibleWords.Count - 1], true) ? "1F" : "NA";
                                                        // new added
                                                        if (Module1.SpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                                        {
                                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                                        }
                                                        //new added
                                                    }
                                                }
                                            }
                                        }

                                    }

                                }
                               
                            }
                            // new added for merge data

                            if (PossibleWords.Count == 0 && ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.Length > 7 && ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Substring(0, 5) == "QUOTE")
                            {
                                string QNoCheck = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord.ToUpper().Replace("#", "").Replace(":", "").Replace("QUOTENUMBER", "").Replace("QUOTEID", "").Replace("QUOTE", "").Trim();
                                if ((Module1.reg1.IsMatch(QNoCheck) || Module1.reg2.IsMatch(QNoCheck)) && (!Char.IsLetter(QNoCheck[QNoCheck.Length - 1]) || !QNoCheck.Any(char.IsLetter)))
                                {
                                    Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);
                                    // new added 16062021
                                    ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                    // new added
                                    if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                    {
                                        PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                        ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                        ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                        if (confidence == true)
                                        {
                                            PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                            PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                        }
                                        else
                                        {
                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                        }
                                        if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                        {
                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                        }
                                        //new added
                                    }

                                }
                                else if (!Char.IsLetter(QNoCheck[QNoCheck.Length - 1]) && QNoCheck.Length > 7)
                                {
                                    Boolean confidence = HighConfidenceforMergedWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], QNoCheck);

                                    // new added 16062021
                                    ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno].strWord = QNoCheck;
                                    // new added 16062021

                                    if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]))
                                    {
                                        PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno]);
                                        ////PossibleWords[PossibleWords.Count - 1].Flag = confidence ? "1" : "0";
                                        ////PossibleWords[PossibleWords.Count - 1].Remarks = confidence ? "1F" : "NA";

                                        if (confidence == true)
                                        {
                                            PossibleWords[PossibleWords.Count - 1].Flag = "1";
                                            PossibleWords[PossibleWords.Count - 1].Remarks = "1F";
                                        }
                                        else
                                        {
                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                        }
                                        if (Module1.MarkRedSpecialCharacters.Any(PossibleWords[PossibleWords.Count - 1].strWord.Contains))
                                        {
                                            PossibleWords[PossibleWords.Count - 1].Flag = "0";
                                            PossibleWords[PossibleWords.Count - 1].Remarks = "NA";
                                        }
                                        //new added

                                    }

                                }
                            }
                            ///// end here

                        } // end of for 

                        if (PossibleWords.Count >= 0)
                        {
                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 4);
                        }

                    }
                    // end of the F3
                }
            }
            catch (Exception)
            {
                // throw;
            }
            finally
            {
                ds.Dispose();
                ds = null;
                //dsKeys.Clear();
                dsKeys.Dispose();
                dsKeys = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                foundRows = null;
                CombinedKeyword = null;
                sKeyword = null;
                objCNC.Dispose();
                cWord = null;
                conflevel = null;
                PossibleWords = null;
            }

            return returnZones;
        }
        #endregion

        #region RADF
        private RetStructF3 F3_AutoLocate_RADF(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            Module1.MABDDate.strWord = null;
            Module1.Dlv_O_flag = false;
            // DataSet ds = new DataSet();
            DataSet ds = Module1.func_GetKeywords(ObjMetaData, intCurrPageNumber);


            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            DataSet dsKeys = new DataSet();
            DataTable dt = new DataTable();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;
            Hashtable ht = new Hashtable();

            string[] MABDArray = { "MABD", "MABD Date", "Must Arrive by date", "Must Deliver by date", "Must deliver by", "deliver by", "Deliver No Later Than", "ARRIVE BY", "MUST DELIVER ON OR BEFORE", "MUST DELIVER BETWEEN", "Must Arrive By Window", "DELIVERY MUST BE MADE BY", "Do Not Del After", "Do Not Deliver After" };
            bool isMABDKeywordFound = false;
            bool isTwodatePresent = false;

            string[] Dlv_O_Keywords = { "Deliver On", "Delivery Date", "Schedule Delivery date", "Sch Del date", "Delivery Date/Time", "Requested Date/Time", "REQ ARV DTE", "Requested Delivery Date", "REQ DATE", "ARRIVE DATE", "Delv Date", "Del Date", "REQ DELIVERY DATE", "Request Date", "Must deliver only on", "Dropoff Date/Time" };
            bool isDLV_O_KeywordFound = false;

            bool isValidDate = false;

            // For Date Rang 
            DataTable dt_Twodate = new DataTable();
            dt_Twodate.Columns.Add("keyword");
            dt_Twodate.Columns.Add("Line No");
            Dictionary<int, clsCnCWord> RangeDateMatch = new Dictionary<int, clsCnCWord>();
            // For Date Rang 

            // For Dlv_O_Keyword
            DataTable dt_O_Dlv = new DataTable();
            dt_O_Dlv.Columns.Add("keyword");
            dt_O_Dlv.Columns.Add("Line No");
            dt_O_Dlv.Columns.Add("Word No");
            dt_O_Dlv.Columns.Add("flag");
            // For Dlv_O_Keyword

            System.Data.DataRow[] foundRows;
            //System.Data.DataRow[] foundRowsTest;
            ClsCNC objCNC = new ClsCNC();
            clsCnCWord cWord = new clsCnCWord();
            clsCnCWord CombinedKeyword = new clsCnCWord();
            clsCNCSBR objSBR = new clsCNCSBR();

            //bool isFuturedate = false;
            Module1.isFuturedate = false;
            Module1.FuturePDTRADF = false;
            Module1.FuturePDTMABD = false;

        Line1:
            ;

            if (iRecursiveCallCount == 0)
            {
                // -------------------------------------------------RADF KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                sKeyword = "'Due Date', 'Delivery Date','Estimated Delivery Date', 'Guaranteed before','Deliver On','Delivery Requested Date','Requested Delivery Date','Deliv From','Delivery Date/Time','Delivery Appointment','Deliver Before Date','To Arrive','Deliver No Earlier Than','Deliver No Later Than','Req Delivery date','Schedule Delivery date','Sch Del date','Estimated Delivery','Requested Date/Time','REQ ARV DTE','DLV Date','REQ DATE','Delv Date','Del Date','Delivery Dt','Delivery','Est. Arrival Date','Appt Date/Time','DE L DATE','Delivery Window','Appt. Date / Time','Request Date','Dropoff Date/Time','Do Not Del Before','Do Not Del After','Do Not Deliver Before','Do Not Deliver After','Est. Arrival','Must deliver only on','MABD','MABD Date','Must Arrive by date','Must Deliver by date','Must Deliver By','Deliver by','ARRIVE BY','MUST DELIVER ON OR BEFORE','ARRIVE DATE','MUST DELIVER BETWEEN','Must Arrive By Window','DELIVERY MUST BE MADE BY'";
            }
            // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            else
            {
                sKeyword = "''";
            }

            try
            {
                returnZones.Status = "F";
                returnZones.ManualConfirmation = "Y";
                returnZones.Flag = "Y";
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_keyword"];
                }
                if (dsKeys != null)
                {
                    dt = dsKeys.Tables[0];                    
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                   
                    var orderedRows = foundRows.OrderBy(item => item.ItemArray[2]).ThenByDescending(item => item.ItemArray[8]);
                    
                    foundRows = orderedRows.ToArray();
                    //remove same line keyword
                    if (foundRows.Length > 1)
                    {
                        foundRows = RemoveDuplicate(foundRows, dt, sKeyword);
                        
                    }
                    string keywordFound = "";
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int Newwordno = 0;
                    int y2 = 0;
                    int w = 0;
                    clsCnCWord[] TopKeywords= new clsCnCWord[10];  // for merge keyword
                    clsCnCWord MergWord = new clsCnCWord();

                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        int KeywordsStartCharY2Bottom = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno].Bottom;
                        int KeywordsStartCharTop = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno].Top;
                        int KeywordHeight = KeywordsStartCharY2Bottom - KeywordsStartCharTop;

                        string[] splitKeword = keywordFound.Split(' ');
                        if (splitKeword.Length > 1)
                        {
                            Newwordno = wordno + splitKeword.Length - 1;
                        }
                        else
                        {
                            Newwordno = wordno;
                        }
                        CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                        ////// for Merge Data and Keyword
                        ////if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                        ////{
                        ////    ContainsKeyword(ObjMetaData, pgno, TopKeywords, TopWord, 2, iRecursiveCallCount, keyword, KeywordLength);
                        ////}
                        ////else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1 && TopKeywords[KeywordLength - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                        ////{
                        ////    ContainsKeyword(ObjMetaData, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 2, iRecursiveCallCount, keyword.Split()[KeywordLength - 1], KeywordLength);
                        ////}
                        //////end

                       

                        // if two date present in line then assign second date to global variable
                        //isTwodatePresent = CheckMultiple_dateMatch(pgno, ObjMetaData.Page[intCurrPageNumber].Line[lineno].strLine);
                        isTwodatePresent = CheckMultiple_dateMatch2(ObjMetaData, intCurrPageNumber, lineno, RangeDateMatch, wordno);
                        if (isTwodatePresent)
                        {
                            DataRow Dr = dt_Twodate.NewRow();
                            Dr["keyword"] = keywordFound;
                            Dr["Line No"] = lineno;
                            dt_Twodate.Rows.Add(Dr);
                        }
                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);

                        if (Module1.ObjBoundingWord.RightWord != null)
                        {
                            // if (Module1.ObjBoundingWord.RightWord.strWord == ":"|| Module1.ObjBoundingWord.RightWord.strWord == "-" || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "IS" || Module1.ObjBoundingWord.RightWord.strWord.ToUpper() == "="|| Module1.ObjBoundingWord.RightWord.strWord.Length<=2)
                            if (Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace("(", "").Replace(")", "").Replace(":", "") == "MABD" || Module1.ObjBoundingWord.RightWord.strWord.Length <= 2)
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                            }
                            if (Module1.ObjBoundingWord.RightWord != null)
                            {
                                Module1.ObjBoundingWord.RightWord.strWord = ReplaceStartEndSpecialchar(Module1.ObjBoundingWord.RightWord.strWord);

                                Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(":", "");
                                RetStructIQSBR008 iQSBR008 = objSBR.IQSBR008(Module1.ObjBoundingWord.RightWord.strWord, "E");
                                if ((iQSBR008.Status == "S" 
                                    || Module1.DMYwithoutSpace.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.YMDwithoutSpace.IsMatch(Module1.ObjBoundingWord.RightWord.strWord)
                                    || Module1.MDYwithChar.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.YMD.IsMatch(Module1.ObjBoundingWord.RightWord.strWord)
                                    )
                                    && Module1.ObjBoundingWord.RightWord.strWord.Length >= 6)
                                {
                                     Module1.isFuturedate = CheckForFutureDate(Module1.ObjBoundingWord.RightWord.strWord);

                                    isValidDate = CheckForValidDate(Module1.ObjBoundingWord.RightWord.strWord);
                                    if (isValidDate == true)
                                    {
                                        isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                        isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);
                                        if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                        {
                                            if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                                            {
                                                PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                            }
                                            if (isMABDKeywordFound && !isTwodatePresent)
                                            {
                                                Module1.MABDDate = Module1.ObjBoundingWord.RightWord;
                                            }
                                        }

                                        if (dt_Twodate.Rows.Count >= 1){
                                            bool DiffdateFound = false;
                                            int linenofound = Convert.ToInt32(dt_Twodate.Rows[dt_Twodate.Rows.Count - 1]["Line No"]);
                                            for (int i = 0; i <= PossibleWords.Count - 1; i++)
                                            {
                                                if (PossibleWords[i].LineNo == linenofound)
                                                {
                                                    DiffdateFound = true;
                                                    cWord = PossibleWords[i];
                                                    PossibleWords.Clear();
                                                    PossibleWords.Add(cWord);
                                                    Module1.MABDDate = RangeDateMatch[linenofound];
                                                }
                                            }
                                            if (!DiffdateFound)
                                            {
                                                Module1.MABDDate = RangeDateMatch[PossibleWords[0].LineNo];
                                            }
                                        }
                                        if (PossibleWords.Count > 0)
                                        {
                                            if (PossibleWords.Count == 1 && isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                            {
                                                Module1.Dlv_O_flag = true;
                                            }
                                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 5);
                                        }
                                    }

                                }
                                else
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                    if (Module1.ObjBoundingWord.BottomWords != null)
                                    {
                                        for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                        {
                                            Module1.ObjBoundingWord.BottomWords[j].strWord = ReplaceStartEndSpecialchar(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                            Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(":", "");

                                            //Bottom TWO date 15-02-2021
                                            isTwodatePresent = CheckMultiple_dateMatch2(ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j].LineNo, RangeDateMatch, (int)Module1.ObjBoundingWord.BottomWords[j].WordNumber);

                                            if (isTwodatePresent)
                                            {
                                                DataRow Dr = dt_Twodate.NewRow();
                                                Dr["keyword"] = keywordFound;
                                                Dr["Line No"] = lineno;
                                                dt_Twodate.Rows.Add(Dr);
                                            }
                                            iQSBR008 = objSBR.IQSBR008(Module1.ObjBoundingWord.BottomWords[j].strWord, "E");
                                            if ((iQSBR008.Status == "S" //|| Module1.DMYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                                || Module1.DMYwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMDwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                                || Module1.MDYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMD.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                               && Module1.ObjBoundingWord.BottomWords[j].strWord.Length >= 6)
                                            {
                                               Module1.isFuturedate = CheckForFutureDate(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                isValidDate = CheckForValidDate(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                if (isValidDate == true)
                                                {
                                                    isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                                    isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);

                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]))
                                                    {
                                                        int TopValue = Module1.ObjBoundingWord.BottomWords[j].Top;
                                                        int DistanceDiff = TopValue - KeywordsStartCharY2Bottom;
                                                        if (DistanceDiff <= KeywordHeight * 3)
                                                        {
                                                            if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                                                            {
                                                                PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                            }
                                                            if (isMABDKeywordFound && !isTwodatePresent)
                                                            {
                                                                Module1.MABDDate = Module1.ObjBoundingWord.BottomWords[j];
                                                            }
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }

                                    if (dt_Twodate.Rows.Count >= 1)
                                    {
                                        bool DiffdateFound = false;
                                        int linenofound = Convert.ToInt32(dt_Twodate.Rows[dt_Twodate.Rows.Count - 1]["Line No"]);
                                        for (int i = 0; i <= PossibleWords.Count - 1; i++)
                                        {
                                            if (PossibleWords[i].LineNo == linenofound)
                                            {
                                                DiffdateFound = true;
                                                cWord = PossibleWords[i];
                                                PossibleWords.Clear();
                                                PossibleWords.Add(cWord);
                                                Module1.MABDDate = RangeDateMatch[linenofound];
                                            }
                                        }
                                        if (!DiffdateFound)
                                        {
                                            if (PossibleWords.Count != 0)
                                            {
                                                Module1.MABDDate = RangeDateMatch[PossibleWords[0].LineNo];
                                            }

                                        }
                                    }
                                    if (PossibleWords.Count > 0)
                                    {
                                        if (PossibleWords.Count == 1 && isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                        {
                                            Module1.Dlv_O_flag = true;
                                        }
                                        returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 5);
                                    }

                                }
                            }

                        }
                        else
                        {
                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                            if (Module1.ObjBoundingWord.BottomWords != null)
                            {
                                for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                {
                                    Module1.ObjBoundingWord.BottomWords[j].strWord = ReplaceStartEndSpecialchar(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                    Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(":", "");

                                    //Botom TWO date 22-02-2021
                                    isTwodatePresent = CheckMultiple_dateMatch2(ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j].LineNo, RangeDateMatch, (int)Module1.ObjBoundingWord.BottomWords[j].WordNumber);

                                    if (isTwodatePresent)
                                    {
                                        DataRow Dr = dt_Twodate.NewRow();
                                        Dr["keyword"] = keywordFound;
                                        Dr["Line No"] = lineno;
                                        dt_Twodate.Rows.Add(Dr);
                                    }
                                    RetStructIQSBR008 iQSBR008 = objSBR.IQSBR008(Module1.ObjBoundingWord.BottomWords[j].strWord, "E");
                                    if ((iQSBR008.Status == "S" //|| Module1.DMYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                        || Module1.DMYwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMDwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                        || Module1.MDYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMD.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                        )
                                        && Module1.ObjBoundingWord.BottomWords[j].strWord.Length >= 6)
                                    {
                                         Module1.isFuturedate = CheckForFutureDate(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                        isValidDate = CheckForValidDate(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                        if (isValidDate == true)
                                        {
                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                            {
                                                isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                                isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);

                                                int TopValue = Module1.ObjBoundingWord.BottomWords[j].Top;
                                                int DistanceDiff = TopValue - KeywordsStartCharY2Bottom;
                                                if (DistanceDiff <= KeywordHeight * 3)
                                                {
                                                    if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                                                    {
                                                        PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                    }
                                                    if (isMABDKeywordFound && !isTwodatePresent)
                                                    {
                                                        Module1.MABDDate = Module1.ObjBoundingWord.BottomWords[j];
                                                    }
                                                    break;
                                                }

                                            }
                                        }

                                    }
                                }
                            }

                            if (dt_Twodate.Rows.Count >= 1)
                            {
                                bool DiffdateFound = false;
                                int linenofound = Convert.ToInt32(dt_Twodate.Rows[dt_Twodate.Rows.Count - 1]["Line No"]);
                                for (int i = 0; i <= PossibleWords.Count - 1; i++)
                                {
                                    if (PossibleWords[i].LineNo == linenofound)
                                    {
                                        DiffdateFound = true;
                                        cWord = PossibleWords[i];
                                        PossibleWords.Clear();
                                        PossibleWords.Add(cWord);
                                        Module1.MABDDate = RangeDateMatch[linenofound];
                                    }
                                }
                                if (!DiffdateFound)
                                {
                                    Module1.MABDDate = RangeDateMatch[PossibleWords[0].LineNo];
                                }
                            }

                            if (PossibleWords.Count > 0)
                            {
                                if (PossibleWords.Count == 1 && isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                {
                                    Module1.Dlv_O_flag = true;
                                }
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 5);
                            }
                        }

                        if (PossibleWords.Count == 0)
                        {
                            isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);
                            Matched_Date(ObjMetaData, intCurrPageNumber, ht, PossibleWords, keywordFound, lineno, wordno, CombinedKeyword);
                            if (PossibleWords.Count > 0)
                            {
                                if (PossibleWords.Count == 1 && isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                {
                                    Module1.Dlv_O_flag = true;
                                }
                                returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, conflevel, iRecursiveCallCount, intCurrPageNumber, cWord, 5);
                            }
                            if (PossibleWords.Count == 0)
                            {
                                returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                            }
                        }

                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);

                            w = iQPUBLIC.PublicComponents.ImageWidth;
                            y2 = iQPUBLIC.PublicComponents.ImageHeight;
                            string[] splitKeword = keywordFound.Split(' ');

                            int KeywordsStartCharY2Bottom = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno].Bottom;
                            int KeywordsStartCharTop = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno].Top;
                            int KeywordHeight = KeywordsStartCharY2Bottom - KeywordsStartCharTop;

                            CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                            //merge keyword and data
                            for (int i = 0; i < splitKeword.Length; i++)
                            {
                                TopKeywords[i] = ObjMetaData.Page[pgno].Line[lineno].Word[wordno + i];
                            }
                            string lastWordOfKeyword = splitKeword[splitKeword.Length - 1];
                            if (CombinedKeyword.strWord.Any(char.IsDigit) && splitKeword.Length == 1 && CombinedKeyword.strWord.Count(char.IsNumber) > 1)
                            {
                                MergWord= ContainsKeyword(ObjMetaData, pgno, TopKeywords, TopKeywords[splitKeword.Length - 1], 5, iRecursiveCallCount, keywordFound, splitKeword.Length);
                            }
                            else if (TopKeywords[splitKeword.Length - 1].strWord.Any(char.IsDigit) && TopKeywords[splitKeword.Length - 1].strWord.Count(char.IsNumber) > 1 && TopKeywords[splitKeword.Length - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                            {
                                MergWord= ContainsKeyword(ObjMetaData, pgno, TopKeywords, TopKeywords[splitKeword.Length - 1], 5, iRecursiveCallCount, keywordFound.Split()[splitKeword.Length - 1], splitKeword.Length);
                            }
                            //end merge keyword and data

                            if (splitKeword.Length > 1)
                                {
                                    Newwordno = wordno + splitKeword.Length - 1;
                                }
                                else
                                {
                                    Newwordno = wordno;
                                }
                                isTwodatePresent = CheckMultiple_dateMatch2(ObjMetaData, intCurrPageNumber, lineno, RangeDateMatch, wordno);
                               
                                if (isTwodatePresent)
                                {
                                    DataRow Dr = dt_Twodate.NewRow();
                                    Dr["keyword"] = keywordFound;
                                    Dr["Line No"] = lineno;
                                    dt_Twodate.Rows.Add(Dr);
                                }
                            if (string.IsNullOrEmpty(MergWord.strWord))
                            {
                                Module1.ObjBoundingWord = objCNC.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[Newwordno], ObjMetaData, pgno);
                            }
                            else
                            {
                                Module1.ObjBoundingWord.RightWord = MergWord;
                            }
                                if (Module1.ObjBoundingWord.RightWord != null)
                                {
                                   
                                    if (Module1.ObjBoundingWord.RightWord.strWord.ToUpper().Replace("(", "").Replace(")", "").Replace(":", "") == "MABD" || Module1.ObjBoundingWord.RightWord.strWord.Length <= 2)
                                    {
                                        Module1.ObjBoundingWord = objCNC.GetBoundingWords(Module1.ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                    }
                                    if (Module1.ObjBoundingWord.RightWord != null)
                                    {
                                        Module1.ObjBoundingWord.RightWord.strWord = ReplaceStartEndSpecialchar(Module1.ObjBoundingWord.RightWord.strWord);

                                        Module1.ObjBoundingWord.RightWord.strWord = Module1.ObjBoundingWord.RightWord.strWord.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(":", "");
                                        RetStructIQSBR008 iQSBR008 = objSBR.IQSBR008(Module1.ObjBoundingWord.RightWord.strWord.Trim(), "E");
                                        if ((iQSBR008.Status == "S" //|| Module1.DMYwithChar.IsMatch(Module1.ObjBoundingWord.RightWord.strWord)
                                            || Module1.DMYwithoutSpace.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.YMDwithoutSpace.IsMatch(Module1.ObjBoundingWord.RightWord.strWord)
                                            || Module1.MDYwithChar.IsMatch(Module1.ObjBoundingWord.RightWord.strWord) || Module1.YMD.IsMatch(Module1.ObjBoundingWord.RightWord.strWord)
                                            )
                                            && Module1.ObjBoundingWord.RightWord.strWord.Length >= 6)
                                        {
                                            Module1.isFuturedate = CheckForFutureDate(Module1.ObjBoundingWord.RightWord.strWord);

                                            isValidDate = CheckForValidDate(Module1.ObjBoundingWord.RightWord.strWord);
                                            if (isValidDate)
                                            {
                                                if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                                                {
                                                  
                                                    isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                                    isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);

                                                    if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                                                    {
                                                        PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                                                    }
                                                    if (isMABDKeywordFound && !isTwodatePresent)
                                                    {
                                                        Module1.MABDDate = Module1.ObjBoundingWord.RightWord;
                                                    }

                                                    if (isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                                    {
                                                        DataRow Dr = dt_O_Dlv.NewRow();
                                                        Dr["keyword"] = keywordFound;
                                                        Dr["Line No"] = lineno;
                                                        Dr["Word No"] = Module1.ObjBoundingWord.RightWord.WordNumber;
                                                        Dr["flag"] = 1;
                                                        dt_O_Dlv.Rows.Add(Dr);
                                                    }
                                                }
                                            }

                                        }
                                        else
                                        {
                                            Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                            if (Module1.ObjBoundingWord.BottomWords != null)
                                            {
                                                for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                                {
                                                    Module1.ObjBoundingWord.BottomWords[j].strWord = ReplaceStartEndSpecialchar(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                    Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(":", "");
                                                    //Botom TWO date 15-02-2021
                                                    isTwodatePresent = CheckMultiple_dateMatch2(ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j].LineNo, RangeDateMatch, (int)Module1.ObjBoundingWord.BottomWords[j].WordNumber);

                                                    if (isTwodatePresent)
                                                    {
                                                        DataRow Dr = dt_Twodate.NewRow();
                                                        Dr["keyword"] = keywordFound;
                                                        Dr["Line No"] = lineno;
                                                        dt_Twodate.Rows.Add(Dr);
                                                    }
                                                    //

                                                    iQSBR008 = objSBR.IQSBR008(Module1.ObjBoundingWord.BottomWords[j].strWord, "E");
                                                    if ((iQSBR008.Status == "S" //|| Module1.DMYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                                        || Module1.DMYwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMDwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                                        || Module1.MDYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMD.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                                        && Module1.ObjBoundingWord.BottomWords[j].strWord.Length >= 6)
                                                    {
                                                        Module1.isFuturedate = CheckForFutureDate(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                        isValidDate = CheckForValidDate(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                        if (isValidDate)
                                                        {
                                                            isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                                            isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);
                                                            if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                                            {
                                                                int TopValue = Module1.ObjBoundingWord.BottomWords[j].Top;
                                                                int DistanceDiff = TopValue - KeywordsStartCharY2Bottom;
                                                                if (DistanceDiff <= KeywordHeight * 3)
                                                                {
                                                                    if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                                                                    {
                                                                        PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                                    }
                                                                    if (isMABDKeywordFound && !isTwodatePresent)
                                                                    {
                                                                        Module1.MABDDate = Module1.ObjBoundingWord.BottomWords[j];
                                                                    }
                                                                    if (isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                                                    {
                                                                        DataRow Dr = dt_O_Dlv.NewRow();
                                                                        Dr["keyword"] = keywordFound;
                                                                        Dr["Line No"] = lineno;
                                                                        Dr["Word No"] = Module1.ObjBoundingWord.BottomWords[j].WordNumber;
                                                                        Dr["flag"] = 1;
                                                                        dt_O_Dlv.Rows.Add(Dr);
                                                                    }
                                                                    break;
                                                                }
                                                                
                                                            }
                                                        }

                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    Module1.ObjBoundingWord = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, pgno);
                                    if (Module1.ObjBoundingWord.BottomWords != null)
                                    {
                                        for (int j = 0; j < Module1.ObjBoundingWord.BottomWords.Length; j++)
                                        {
                                            Module1.ObjBoundingWord.BottomWords[j].strWord = ReplaceStartEndSpecialchar(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                            Module1.ObjBoundingWord.BottomWords[j].strWord = Module1.ObjBoundingWord.BottomWords[j].strWord.Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(":", "");
                                            //Botom TWO date 15-02-2021
                                            isTwodatePresent = CheckMultiple_dateMatch2(ObjMetaData, intCurrPageNumber, Module1.ObjBoundingWord.BottomWords[j].LineNo, RangeDateMatch, (int)Module1.ObjBoundingWord.BottomWords[j].WordNumber);

                                            if (isTwodatePresent)
                                            {
                                                DataRow Dr = dt_Twodate.NewRow();
                                                Dr["keyword"] = keywordFound;
                                                Dr["Line No"] = lineno;
                                                dt_Twodate.Rows.Add(Dr);
                                            }
                                            //
                                            RetStructIQSBR008 iQSBR008 = objSBR.IQSBR008(Module1.ObjBoundingWord.BottomWords[j].strWord, "E");
                                            if ((iQSBR008.Status == "S"
                                                || Module1.DMYwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMDwithoutSpace.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord)
                                                || Module1.MDYwithChar.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord) || Module1.YMD.IsMatch(Module1.ObjBoundingWord.BottomWords[j].strWord))
                                                && Module1.ObjBoundingWord.BottomWords[j].strWord.Length >= 6)
                                            {
                                                Module1.isFuturedate = CheckForFutureDate(Module1.ObjBoundingWord.BottomWords[j].strWord);

                                                isValidDate = CheckForValidDate(Module1.ObjBoundingWord.BottomWords[j].strWord);
                                                if (isValidDate)
                                                {
                                                    isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                                    isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);

                                                    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.BottomWords[j]) && !(Module1.ObjBoundingWord.BottomWords[j].strWord == ":"))
                                                    {
                                                        int TopValue = Module1.ObjBoundingWord.BottomWords[j].Top;
                                                        int DistanceDiff = TopValue - KeywordsStartCharY2Bottom;
                                                        if (DistanceDiff <= KeywordHeight * 3)
                                                        {
                                                            if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                                                            {
                                                                PossibleWords.Add(Module1.ObjBoundingWord.BottomWords[j]);
                                                            }
                                                            if (isMABDKeywordFound && !isTwodatePresent)
                                                            {
                                                                Module1.MABDDate = Module1.ObjBoundingWord.BottomWords[j];
                                                            }
                                                            if (isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                                            {
                                                                DataRow Dr = dt_O_Dlv.NewRow();
                                                                Dr["keyword"] = keywordFound;
                                                                Dr["Line No"] = lineno;
                                                                Dr["Word No"] = Module1.ObjBoundingWord.BottomWords[j].WordNumber;
                                                                Dr["flag"] = 1;
                                                                dt_O_Dlv.Rows.Add(Dr);
                                                            }
                                                            break;
                                                        }
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                                if (PossibleWords.Count == 0)
                                {
                                    isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                                    isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);

                                    Matched_Date(ObjMetaData, intCurrPageNumber, ht, PossibleWords, keywordFound, lineno, wordno, CombinedKeyword);
                                    if (PossibleWords.Count > 0 && isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                    {
                                        Module1.Dlv_O_flag = true;

                                        if (isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                                        {
                                            DataRow Dr = dt_O_Dlv.NewRow();
                                            Dr["keyword"] = keywordFound;
                                            Dr["Line No"] = PossibleWords[PossibleWords.Count - 1].LineNo;
                                            Dr["Word No"] = PossibleWords[PossibleWords.Count - 1].WordNumber;
                                            Dr["flag"] = 1;
                                            dt_O_Dlv.Rows.Add(Dr);
                                        }
                                    }
                                }
                           
                        } // end of for 
                        if (PossibleWords.Count > 0)
                        {
                            if (dt_Twodate.Rows.Count >= 1)
                            {
                                bool DiffdateFound = false;
                                int linenofound = Convert.ToInt32(dt_Twodate.Rows[dt_Twodate.Rows.Count - 1]["Line No"]);
                                for (int i = 0; i <= PossibleWords.Count - 1; i++)
                                {
                                    if (PossibleWords[i].LineNo == linenofound)
                                    {
                                        DiffdateFound = true;
                                        cWord = PossibleWords[i];
                                        PossibleWords.Clear();
                                        PossibleWords.Add(cWord);
                                        if (RangeDateMatch.Count > 0)
                                        {
                                            Module1.MABDDate = RangeDateMatch[linenofound];
                                        }
                                    }
                                }
                                if (!DiffdateFound)
                                {
                                    if (RangeDateMatch.Count > 0)
                                    {
                                        Module1.MABDDate = RangeDateMatch[PossibleWords[0].LineNo];
                                    }

                                }
                            }
                        }
                        if (PossibleWords.Count > 3)
                        {
                            PossibleWords = GetTopThreeValue(PossibleWords);
                            CompairConfidence(PossibleWords);
                        }
                        if (PossibleWords.Count == 1)
                        {
                            Boolean ConfidenceFlag = HighestConfidence(PossibleWords[0]);
                            //if (ConfidenceFlag == true)// without future date chk 
                            if (ConfidenceFlag == true && Module1.isFuturedate) 
                            {
                                PossibleWords[0].Confidence = 100;
                                if (PossibleWords[0].Flag == null)
                                {
                                    // new added
                                    if (Module1.MarkRedDateSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                    {
                                        PossibleWords[0].Flag = "0";
                                        PossibleWords[0].Remarks = "NA";
                                        Module1.FuturePDTRADF = false;
                                    }
                                    else
                                    {
                                        PossibleWords[0].Flag = "1";
                                        PossibleWords[0].Remarks = "1F";
                                        Module1.FuturePDTRADF = true;
                                    }
                                    // new added

                                    ////PossibleWords[0].Flag = "1";
                                    ////PossibleWords[0].Remarks = "1F";
                                    ////Module1.FuturePDTRADF = true;
                                }
                            }
                            else
                            {
                                if (PossibleWords[0].Flag == null)
                                {
                                    PossibleWords[0].Flag = "0";
                                    PossibleWords[0].Remarks = "NA";
                                    Module1.FuturePDTRADF = false;
                                }
                            }

                            if (dt_O_Dlv.Rows.Count > 0)
                            {
                                foreach (DataRow row in dt_O_Dlv.Rows)
                                {
                                    if ((Convert.ToInt32(row["Line No"]) == PossibleWords[0].LineNo && Convert.ToInt32(row["Word No"]) == (int)PossibleWords[0].WordNumber))
                                    {
                                        Module1.Dlv_O_flag = true;
                                    }
                                }
                            }
                            returnZones.Words = PossibleWords;
                            returnZones.NoOfieldsSuspects = PossibleWords.Count;
                            conflevel.Add(90);
                            returnZones.ConfidenceLevelofSuspect = conflevel;
                            returnZones.Flag = "Y";
                            if (iRecursiveCallCount == 1)
                            {
                                if (PossibleWords[0].Flag == null)
                                {
                                    PossibleWords[0].Flag = "0";
                                    PossibleWords[0].Remarks = "NA";
                                }
                                returnZones.ManualConfirmation = "Y";
                            }
                            else
                            {
                                returnZones.ManualConfirmation = "N";
                            }
                            returnZones.Status = "S";
                            return returnZones;
                        }
                        else
                        {
                            returnZones.Words = PossibleWords;
                            returnZones.NoOfieldsSuspects = PossibleWords.Count;
                            switch (returnZones.NoOfieldsSuspects)
                            {
                                case 0:
                                    {
                                        returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                                    }
                                    break;
                                case 2:
                                    {
                                        Boolean ConfidenceFlag = HighestConfidence(PossibleWords[0]);
                                        //Date diff find
                                        clsCnCWord date1 = PossibleWords[0];
                                        clsCnCWord date2 = PossibleWords[1];
                                        //CultureInfo provider = new CultureInfo("en-US");

                                        try
                                        {
                                            DateTime dt1 = DateTime.Parse(PossibleWords[0].strWord.ToString(), Module1.provider);
                                            DateTime dt2 = DateTime.Parse(PossibleWords[1].strWord.ToString(), Module1.provider);

                                            if (Module1.ToadysDate.Date <= dt1.Date)
                                            {
                                                Module1.isFuturedate = true;
                                            }
                                            if (Module1.ToadysDate.Date <= dt2.Date)
                                            {
                                                Module1.isFuturedate = true;
                                            }
                                            else
                                            {
                                                Module1.isFuturedate = false;
                                            }
                                            if (Module1.ToadysDate.Date <= dt1.Date) // to check future date
                                            {
                                                if (dt1 < dt2)
                                                {
                                                    PossibleWords.Clear();
                                                    PossibleWords.Add(date1);
                                                    Module1.MABDDate = date2;
                                                }
                                                else
                                                {
                                                    PossibleWords.Clear();
                                                    PossibleWords.Add(date2);
                                                    Module1.MABDDate = date1;
                                                }
                                            }
                                            else { PossibleWords.Clear(); }


                                            if (PossibleWords.Count > 0)
                                            {
                                                ////if (ConfidenceFlag == true) //without future date
                                                if (ConfidenceFlag = true && Module1.isFuturedate)
                                                {
                                                    PossibleWords[0].Confidence = 100;
                                                    if (PossibleWords[0].Flag == null)
                                                    {
                                                        // new added
                                                        if (Module1.MarkRedDateSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                                                        {
                                                            PossibleWords[0].Flag = "0";
                                                            PossibleWords[0].Remarks = "NA";
                                                            Module1.FuturePDTRADF = false;
                                                        }
                                                        // new added
                                                        else
                                                        {
                                                            PossibleWords[0].Flag = "1";
                                                            PossibleWords[0].Remarks = "1F";
                                                            Module1.FuturePDTRADF = true;
                                                        }
                                                        ////PossibleWords[0].Flag = "1";
                                                        ////PossibleWords[0].Remarks = "1F";
                                                        ////Module1.FuturePDTRADF = true;
                                                    }
                                                }
                                                else
                                                {
                                                    if (PossibleWords[0].Flag == null)
                                                    {
                                                        PossibleWords[0].Flag = "0";
                                                        PossibleWords[0].Remarks = "NA";
                                                        Module1.FuturePDTRADF = false;
                                                    }
                                                }
                                                if (dt_O_Dlv.Rows.Count > 0)
                                                {
                                                    foreach (DataRow row in dt_O_Dlv.Rows)
                                                    {
                                                        if ((Convert.ToInt32(row["Line No"]) == PossibleWords[0].LineNo && Convert.ToInt32(row["Word No"]) == (int)PossibleWords[0].WordNumber))
                                                        {
                                                            Module1.Dlv_O_flag = true;
                                                        }
                                                    }
                                                }
                                                returnZones.Words = PossibleWords;
                                                returnZones.NoOfieldsSuspects = PossibleWords.Count;
                                                conflevel.Add(90);
                                                returnZones.ConfidenceLevelofSuspect = conflevel;
                                                returnZones.Flag = "Y";
                                                returnZones.ManualConfirmation = "N";
                                                returnZones.Status = "S";
                                                ////return returnZones; //without Future date
                                            }
                                            else
                                            {
                                                // PossibleWords.Clear();
                                                returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                                            }
                                            return returnZones;
                                        }
                                        catch (Exception ex)
                                        {
                                            PossibleWords.Clear();

                                            returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                                        }

                                        return returnZones;
                                    }

                                case 3:
                                    {
                                        PossibleWords.Clear();
                                        returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                                        return returnZones;
                                    }
                            }
                        }

                    }
                    // end of the F3
                }
            }
            catch (Exception ex)
            {
                //throw;
            }
            finally
            {
                ds.Dispose();
                ds = null;
                //dsKeys.Clear();
                dsKeys.Dispose();
                dsKeys = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                foundRows = null;
                CombinedKeyword = null;
                sKeyword = null;
                objCNC = null;
                cWord = null;
                objSBR = null;
                ht = null;
                MABDArray = null;
                Dlv_O_Keywords = null;
                dt_Twodate.Dispose();
                dt_O_Dlv.Dispose();
                conflevel = null;
                PossibleWords = null;
            }
            return returnZones;
        }
        #endregion

        #region MABD
        private RetStructF3 F3_AutoLocate_MABD(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            clsCnCWord cWord = new clsCnCWord();
            Module1.FuturePDTMABD = false;
            
            try
            {
                if (Module1.MABDDate.strWord != null)
                {
                    PossibleWords.Add(Module1.MABDDate);
                    Module1.isFuturedate = CheckForFutureDate(Module1.MABDDate.strWord);
                    if (PossibleWords.Count > 0)
                    {
                        Boolean ConfidenceFlag = HighestConfidence(PossibleWords[0]);
                        //if (ConfidenceFlag == true) // without Future date
                        if (ConfidenceFlag == true && Module1.isFuturedate) 
                        {
                            // new added
                            if (Module1.MarkRedDateSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                            {
                                PossibleWords[0].Flag = "0";
                                PossibleWords[0].Remarks = "NA";
                                PossibleWords[0].Confidence = 100;
                                Module1.FuturePDTMABD = false;
                            }
                            else
                            {
                                PossibleWords[0].Flag = "1";
                                PossibleWords[0].Remarks = "1F";
                                Module1.FuturePDTMABD = true;
                            }
                            // new added

                            ////PossibleWords[0].Flag = "1";
                            ////PossibleWords[0].Remarks = "1F";
                            ////PossibleWords[0].Confidence = 100;
                            ////Module1.FuturePDTMABD = true;
                        }
                        else
                        {
                            PossibleWords[0].Flag = "0";
                            PossibleWords[0].Remarks = "NA";
                            Module1.FuturePDTMABD = false;
                        }
                        returnZones.Words = PossibleWords;
                        conflevel.Add(90);
                        returnZones.ManualConfirmation = "N";
                        returnZones.Flag = "Y";
                        returnZones.Status = "S";
                        returnZones.ConfidenceLevelofSuspect = conflevel;
                        returnZones.NoOfieldsSuspects = PossibleWords.Count;
                        return returnZones;
                    }
                }
                else
                {
                    returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
                }

            }
            catch (Exception ex)
            { }
            finally
            {
                cWord = null;
                conflevel = null;
                PossibleWords = null;
            }

            return returnZones;
        }
        #endregion

        #region DLV
        private RetStructF3 F3_AutoLocate_DLV(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            List<int> conflevel = new List<int>();
            conflevel.Clear();
            clsCnCWord cWord = new clsCnCWord();
            try
            {
                string RADFDate = Module1.RADF_PDT;
                string MABDDate = Module1.MABD_PDT;
                if (!string.IsNullOrEmpty(RADFDate) && !string.IsNullOrEmpty(MABDDate) && RADFDate != "*****" && MABDDate != "*****")
                {
                    cWord = StringToCNC(intCurrPageNumber, "W");
                }
                else if ((string.IsNullOrEmpty(RADFDate) && string.IsNullOrEmpty(MABDDate)) || (RADFDate == "*****" && MABDDate == "*****") || (string.IsNullOrEmpty(RADFDate) && MABDDate == "*****"))
                {
                    cWord = StringToCNC(intCurrPageNumber, "*****");
                }
                else if ((!string.IsNullOrEmpty(RADFDate) && Module1.Dlv_O_flag == false && (string.IsNullOrEmpty(MABDDate) || MABDDate == "*****")) || (!string.IsNullOrEmpty(MABDDate) && (string.IsNullOrEmpty(RADFDate) || RADFDate == "*****")))
                {
                    cWord = StringToCNC(intCurrPageNumber, "B");
                }
                else if (!string.IsNullOrEmpty(RADFDate) && Module1.Dlv_O_flag == true && (string.IsNullOrEmpty(MABDDate) || MABDDate == "*****"))
                {
                    cWord = StringToCNC(intCurrPageNumber, "O");
                }
                PossibleWords.Add(cWord);
                if (PossibleWords.Count > 0)
                {
                    if (PossibleWords[0].strWord != "*****")
                    {
                        Boolean ConfidenceFlag = HighestConfidence(PossibleWords[0]);
                        if (ConfidenceFlag == true && (Module1.FuturePDTRADF == true || Module1.FuturePDTMABD == true)) 
                        {
                            PossibleWords[0].Flag = "1";
                            PossibleWords[0].Remarks = "1F";
                            PossibleWords[0].Confidence = 100;
                        }
                        else
                        {
                            PossibleWords[0].Flag = "0";
                            PossibleWords[0].Remarks = "NA";
                        }
                    }

                    returnZones.Words = PossibleWords;
                    conflevel.Add(90);
                    returnZones.ManualConfirmation = "N";
                    returnZones.Flag = "Y";
                    returnZones.Status = "S";
                    returnZones.ConfidenceLevelofSuspect = conflevel;
                    returnZones.NoOfieldsSuspects = PossibleWords.Count;
                    return returnZones;
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                cWord = null;
                conflevel = null;
                PossibleWords = null;
            }

            return returnZones;
        }
        #endregion

        #region Special Instruction
        private RetStructF3 F3_AutoLocate_Special_Instruction(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            RetStructF3 returnZones = new RetStructF3();
            string _sKeyword = " 'SPECIAL INSTRUCTION', 'Comments/Special Instruction', 'Shipping Instructions','Consignee Special Instructions', 'Shipper Special Instructions','Special Services and Instructions', 'Carrier Special delivery Instructions','Origin Instructions','SpecialInstructions','Special delivery Instructions','Shipper Instructions','Shipper Instruction','Consignee Instruction' ";
            try
            {
                returnZones = F3_AutoLocate_Common_Instructions(ObjMetaData, intCurrPageNumber, strArg, _sKeyword, Module1.dtAccessorials, "SI");
            }
            catch (Exception ex)
            { }
            finally
            {
                _sKeyword = null;
            }

            return returnZones;
        }
        #endregion

        #region Delivery Requirment
        private RetStructF3 F3_AutoLocate_Delivery_Requirment(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            RetStructF3 returnZones = new RetStructF3();
            string _sKeyword = " 'DESTINATION INSTRUCTIONS','Delivery Instructions','Delivery Notes','Delivery Remarks'  ";
            try
            {
                returnZones = F3_AutoLocate_Common_Instructions(ObjMetaData, intCurrPageNumber, strArg, _sKeyword, Module1.dtAccessorials, "DR");
            }
            catch (Exception ex)
            { }
            finally
            {
                _sKeyword = null;
            }
            return returnZones;
        }
        #endregion

        #region
        private RetStructF3 F3_AutoLocate_Accessorial(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            RetStructF3 returnZones = new RetStructF3();
            string _sKeyword = " 'Destination Accessorials','Delivery Accessorials','Accessorial Required','Accessorials'  ";
            try
            {
                returnZones = F3_AutoLocate_Common_Instructions(ObjMetaData, intCurrPageNumber, strArg, _sKeyword, Module1.dtAccessorials, "Acc");
            }
            catch (Exception ex)
            { }
            finally
            {
                _sKeyword = null;
            }

            return returnZones;
        }
        #endregion

        #region Common Instructions
        private RetStructF3 F3_AutoLocate_Common_Instructions(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg, string i_Keywords, DataTable db_datatable, string Function_Id)
        {
            string ShipperName = string.Empty;
            //bool IgnoreshipperExist = false;
            if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("BOLSET3_ShipperName"))
            {
                ShipperName = (string)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_ShipperName"];

            }

            if (Function_Id == "SI")
            {
                Module1.ComnInstrConatainer = Module1.AssignInstructionDt();
                Module1.pgSearchInstrut = false;
                Module1.IgnoreshipperExist = false;
                Module1.IgnoreshipperExist = IgnoreShipperMatching(ShipperName);
            }
            //for ignoreshippers special instruction
            // bool IgnoreshipperExist = Module1.Dt_IgnoreShipperDetails.AsEnumerable().Any(row => ShipperName == row.Field<String>("ShipperName"));
            //bool IgnoreshipperExist= IgnoreShipperMatching("Integrated Supply Network,LLC");
            ////end

            // DataSet ds = new DataSet();
            DataSet ds = Module1.func_GetKeywords(ObjMetaData, intCurrPageNumber);

            RetStructF3 returnZones = new RetStructF3();
            List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
            DataSet dsKeys = new DataSet();
            DataTable dt = new DataTable();
            List<int> ConfLevel = new List<int>();
            ConfLevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;

            List<string> Inst_Indicator = new List<string>();

            ClsCNC objCNC = new ClsCNC();
            clsCnCWord cWord = new clsCnCWord();

            System.Data.DataRow[] foundRows;
            clsCnCBoundingWords ObjBoundingWord = new clsCnCBoundingWords();
            clsCnCBoundingWords ObjBoundingWord1 = new clsCnCBoundingWords();
            clsCnCBoundingWords ObjBoundingWordNext = new clsCnCBoundingWords();
            clsCnCWord[] SIContainer = new clsCnCWord[500];
            clsCnCWord SInstrction = new clsCnCWord();
            clsCnCWord CombinedKeyword = new clsCnCWord();

        Line1:
            ;

            if (iRecursiveCallCount == 0 && Module1.IgnoreshipperExist == false)
            {
                // -------------------------------------------------Delivery Requirement KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                sKeyword = i_Keywords;
            }

            // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
            else
            {
                sKeyword = "''";
            }
            try
            {
                returnZones.Status = "F";
                returnZones.ManualConfirmation = "Y";
                returnZones.Flag = "Y";
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSET3_keyword"))
                {
                    dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSET3_keyword"];
                }
                if (dsKeys != null)
                {

                    dt = dsKeys.Tables[0];
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();
                    string keywordFound = "";
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;

                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        
                        string[] splitKeword = keywordFound.Split(' ');
                        int KeywordLineNo = 0;
                        int KeywordStratWordNo = 0;
                        int Firstcharx1 = 0;
                        int LastcharX2 = 0;
                        int TempLastWordX2Right = 0;
                        int TempstratX1Left = 0;
                        int TempstratY2Bottom = 0;
                        int BottomwordNo = 0;
                        int BottomwordNoNext = 0;
                        bool validLine = false;
                        bool flagNextword = false;
                        int MatchkeywordLength = 0;

                        int k = 0;
                        int Icounter = 0;
                        int NexrRightwordX1Left = 0;
                        int validdiff = 0;
                        bool flagBotoomwordFound = false;
                        bool stopLoop = false;
                        KeywordLineNo = lineno;
                        KeywordStratWordNo = wordno;
                        MatchkeywordLength = splitKeword.Length;
                        CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                        int KeywordsStartCharX1Left = ObjMetaData.Page[intCurrPageNumber].Line[KeywordLineNo].Word[wordno].Left;
                        int KeywordsLastWordX2Right = ObjMetaData.Page[intCurrPageNumber].Line[KeywordLineNo].Word[KeywordStratWordNo + MatchkeywordLength - 1].Right;
                        int KeywordsStartCharY2Bottom = ObjMetaData.Page[intCurrPageNumber].Line[KeywordLineNo].Word[KeywordStratWordNo].Bottom;
                        int DiffInword = 0;
                        TempLastWordX2Right = KeywordsLastWordX2Right;
                        TempstratX1Left = KeywordsStartCharX1Left;
                        TempstratY2Bottom = KeywordsStartCharY2Bottom;
                        
                        int firstcount = 1;
                        for (int iLine = KeywordLineNo; iLine < KeywordLineNo + 10; iLine++)
                        {
                            if (stopLoop)
                            { break; }
                            if (iLine == KeywordLineNo && iLine < ObjMetaData.Page[intCurrPageNumber].LineCount)
                            {
                                for (int iWord = KeywordStratWordNo + MatchkeywordLength - 1; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; iWord++)
                                {
                                    clsCnCWord Word = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                    ObjBoundingWord = objCNC.GetBoundingWords(Word, ObjMetaData, intCurrPageNumber);
                                    // check if Right word present or not
                                    if (ObjBoundingWord.RightWord != null)
                                    {
                                        if (ObjBoundingWord.RightWord.strWord == ":" || ObjBoundingWord.RightWord.strWord == "|")
                                        {
                                            ObjBoundingWord = objCNC.GetBoundingWords(ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                        }
                                        if (firstcount == 1)
                                        {
                                            if (ObjBoundingWord.RightWord != null)
                                            {
                                                //If bottom word not present exact below the keyword
                                                ObjBoundingWordNext = objCNC.GetBoundingWords(ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                                flagNextword = true;
                                                NexrRightwordX1Left = ObjBoundingWord.RightWord.Left;
                                                if (ObjBoundingWordNext.BottomWords != null)
                                                {
                                                    if (ObjBoundingWordNext.BottomWords.Length > 0)
                                                    {
                                                        BottomwordNoNext = (int)ObjBoundingWordNext.BottomWords[0].WordNumber;
                                                    }
                                                }
                                            }
                                            firstcount++;
                                        }

                                    }

                                    if (iWord == KeywordStratWordNo + MatchkeywordLength - 1)
                                    {
                                        ObjBoundingWord1 = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, intCurrPageNumber);
                                        if (ObjBoundingWord1.BottomWords != null)
                                        {
                                            if (ObjBoundingWord1.BottomWords.Length > 0)
                                            {
                                                BottomwordNo = (int)ObjBoundingWord1.BottomWords[0].WordNumber;
                                                flagBotoomwordFound = true;
                                            }
                                        }
                                    }
                                    if (iWord == KeywordStratWordNo + MatchkeywordLength - 1 && flagNextword && !flagBotoomwordFound)
                                    {
                                        BottomwordNo = BottomwordNoNext;
                                    }


                                    if (ObjBoundingWord.RightWord != null)
                                    {
                                        Firstcharx1 = ObjBoundingWord.RightWord.Left;
                                        LastcharX2 = ObjBoundingWord.RightWord.Right;
                                        DiffInword = Firstcharx1 - TempLastWordX2Right;

                                        if (iWord == KeywordStratWordNo + MatchkeywordLength - 1)
                                        {
                                            validdiff = DiffInword;
                                        }

                                        if (DiffInword < 60)//(DiffInword < 69)
                                        {
                                            SIContainer[k] = ObjBoundingWord.RightWord;
                                            k++;
                                            TempLastWordX2Right = LastcharX2;
                                        }
                                        else
                                        {
                                            iWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount;
                                        }
                                    }
                                }
                            }
                            else if (iLine > KeywordLineNo && iLine < KeywordLineNo + 10 && iLine < ObjMetaData.Page[intCurrPageNumber].LineCount)
                            {
                                Icounter = 0;
                                if (BottomwordNo != 0)
                                {
                                    for (int iWord = BottomwordNo; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; iWord++)
                                    {
                                        clsCnCWord Word = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                        ObjBoundingWord = objCNC.GetBoundingWords(Word, ObjMetaData, intCurrPageNumber);

                                        int x1left = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Left;
                                        int Y1Top = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Top;
                                        if (Y1Top - TempstratY2Bottom > 50)
                                        {
                                            stopLoop = true;
                                            break;
                                        }
                                        if (iWord == BottomwordNo && Icounter < 1)
                                        {
                                            if (ObjBoundingWord.BottomWords != null)
                                            {
                                                if (ObjBoundingWord.BottomWords.Length > 0)
                                                {
                                                    BottomwordNo = (int)ObjBoundingWord.BottomWords[0].WordNumber;
                                                    TempstratY2Bottom = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Bottom;
                                                }
                                            }

                                            if (!flagBotoomwordFound && iLine == KeywordLineNo + 1 && validdiff < 60)
                                            //if (!flagBotoomwordFound && iLine == KeywordLineNo + 1 && validdiff < 69)
                                            {
                                                if (((x1left - NexrRightwordX1Left) >= -15 && (x1left - NexrRightwordX1Left) < 60))
                                                {
                                                    validLine = true;
                                                    SIContainer[k] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                                    k++;
                                                    Icounter++;
                                                    TempLastWordX2Right = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Right;
                                                    KeywordsStartCharX1Left = NexrRightwordX1Left;
                                                }
                                                else
                                                {
                                                    validLine = false;
                                                }
                                            }
                                            else if (((x1left - KeywordsStartCharX1Left) >= -15 && (x1left - KeywordsStartCharX1Left) < 60))
                                            {
                                                validLine = true;
                                                SIContainer[k] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                                k++;
                                                Icounter++;
                                                TempLastWordX2Right = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Right;
                                            }
                                            else
                                            {
                                                validLine = false;
                                            }
                                        }

                                        if (validLine)
                                        {
                                            ObjBoundingWord = objCNC.GetBoundingWords(Word, ObjMetaData, intCurrPageNumber);
                                            if (iWord == BottomwordNo)
                                            {
                                                if (ObjBoundingWord.BottomWords != null)
                                                {
                                                    if (ObjBoundingWord.BottomWords.Length > 0)
                                                    {
                                                        BottomwordNo = (int)ObjBoundingWord.BottomWords[0].WordNumber;
                                                    }
                                                }

                                            }
                                            if (ObjBoundingWord.RightWord != null)
                                            {
                                                Firstcharx1 = ObjBoundingWord.RightWord.Left;
                                                LastcharX2 = ObjBoundingWord.RightWord.Right;
                                                DiffInword = Firstcharx1 - TempLastWordX2Right;

                                                if (DiffInword < 60)
                                                {
                                                    SIContainer[k] = ObjBoundingWord.RightWord;
                                                    k++;
                                                    TempLastWordX2Right = LastcharX2;
                                                }
                                                else
                                                {
                                                    iWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount;
                                                }
                                            }
                                        }
                                        else
                                        { break; }
                                    }
                                }
                            }
                        }
                        SInstrction = objCNC.MergeWords(SIContainer);

                        bool Dbflag = MatchingSIwithDatabasevalues(SInstrction, intCurrPageNumber, PossibleWords, db_datatable);
                        if (!Dbflag && Function_Id != "Acc")
                        {
                            if (!string.IsNullOrEmpty(SInstrction.strWord))
                            {
                                cWord = ModifiedStringToCNC(SInstrction, "*****");
                                cWord.Confidence = 20;
                                if (!FoundWord(PossibleWords, cWord))
                                {
                                    PossibleWords.Add(cWord);
                                }
                            }
                        }
                        if (PossibleWords.Count == 1 && string.IsNullOrEmpty(PossibleWords[0].strWord))
                        {
                            PossibleWords.Clear();
                        }

                        if (Function_Id == "Acc" && PossibleWords.Count == 0)
                        {
                            Module1.pgSearchInstrut = true;
                            Search_SI_WholePg(ObjMetaData, intCurrPageNumber, PossibleWords);
                        }

                        if (PossibleWords.Count > 0)
                        {
                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, ConfLevel, iRecursiveCallCount, intCurrPageNumber, cWord, 6);
                        }
                        if (PossibleWords.Count == 0)
                        {
                            returnZones = StarReturn(returnZones, PossibleWords, ConfLevel, intCurrPageNumber, SInstrction);
                        }
                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);

                            string[] splitKeword = keywordFound.Split(' ');
                            int KeywordLineNo = 0;
                            int KeywordStratWordNo = 0;
                            int Firstcharx1 = 0;
                            int LastcharX2 = 0;
                            int TempLastWordX2Right = 0;
                            int TempstratX1Left = 0;
                            int TempstratY2Bottom = 0;
                            int BottomwordNo = 0;
                            int BottomwordNoNext = 0;
                            int NexrRightwordX1Left = 0;
                            bool validLine = false;
                            bool flagNextword = false;
                            int MatchkeywordLength = 0;
                            bool stopLoop = false;
                            bool flagBotoomwordFound = false;
                            int validdiff = 0;
                            int k = 0;
                            int Icounter = 0;

                            KeywordLineNo = lineno;
                            KeywordStratWordNo = wordno;
                            MatchkeywordLength = splitKeword.Length;
                            CombinedKeyword = GetJoinKeywordWord(objCNC, ObjMetaData, intCurrPageNumber, lineno, wordno, splitKeword.Length);

                            int KeywordsStartCharX1Left = ObjMetaData.Page[intCurrPageNumber].Line[KeywordLineNo].Word[KeywordStratWordNo].Left;
                            int KeywordsLastWordX2Right = ObjMetaData.Page[intCurrPageNumber].Line[KeywordLineNo].Word[KeywordStratWordNo + MatchkeywordLength - 1].Right;
                            int KeywordsStartCharY2Bottom = ObjMetaData.Page[intCurrPageNumber].Line[KeywordLineNo].Word[KeywordStratWordNo].Bottom;
                            int DiffInword = 0;
                            TempLastWordX2Right = KeywordsLastWordX2Right;
                            TempstratX1Left = KeywordsStartCharX1Left;
                            TempstratY2Bottom = KeywordsStartCharY2Bottom;
                            int firstcount = 1;

                            for (int iLine = KeywordLineNo; iLine < KeywordLineNo + 10; iLine++)
                            {
                                if (stopLoop)
                                { break; }
                                if (iLine == KeywordLineNo && iLine < ObjMetaData.Page[intCurrPageNumber].LineCount)
                                {
                                    for (int iWord = KeywordStratWordNo + MatchkeywordLength - 1; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; iWord++)
                                    {
                                        clsCnCWord Word = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                        ObjBoundingWord = objCNC.GetBoundingWords(Word, ObjMetaData, intCurrPageNumber);

                                        // If bottomword present exact bottom of first keyword

                                        if (iWord == KeywordStratWordNo + MatchkeywordLength - 1)
                                        {
                                            ObjBoundingWord1 = objCNC.GetBoundingWords(CombinedKeyword, ObjMetaData, intCurrPageNumber);
                                            if (ObjBoundingWord1.BottomWords != null)
                                            {
                                                if (ObjBoundingWord1.BottomWords.Length > 0)
                                                {
                                                    BottomwordNo = (int)ObjBoundingWord1.BottomWords[0].WordNumber;
                                                    flagBotoomwordFound = true;
                                                }
                                            }
                                        }

                                        // check if Right word present or not
                                        if (ObjBoundingWord.RightWord != null)
                                        {
                                            if (ObjBoundingWord.RightWord.strWord == ":" || ObjBoundingWord.RightWord.strWord == "|")
                                            {
                                                ObjBoundingWord = objCNC.GetBoundingWords(ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                            }
                                            if (firstcount == 1)
                                            {
                                                if (ObjBoundingWord.RightWord != null)
                                                {
                                                    ObjBoundingWordNext = objCNC.GetBoundingWords(ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                                                    flagNextword = true;
                                                    NexrRightwordX1Left = ObjBoundingWord.RightWord.Left;
                                                    if (ObjBoundingWordNext.BottomWords != null)
                                                    {
                                                        if (ObjBoundingWordNext.BottomWords.Length > 0)
                                                        {
                                                            BottomwordNoNext = (int)ObjBoundingWordNext.BottomWords[0].WordNumber;

                                                        }
                                                    }
                                                }
                                                firstcount++;
                                            }
                                        }

                                        if (iWord == KeywordStratWordNo + MatchkeywordLength - 1 && flagNextword && !flagBotoomwordFound)
                                        {
                                            BottomwordNo = BottomwordNoNext;
                                        }

                                        if (ObjBoundingWord.RightWord != null)
                                        {
                                            Firstcharx1 = ObjBoundingWord.RightWord.Left;
                                            LastcharX2 = ObjBoundingWord.RightWord.Right;
                                            DiffInword = Firstcharx1 - TempLastWordX2Right;
                                            if (iWord == KeywordStratWordNo + MatchkeywordLength - 1)
                                            {
                                                validdiff = DiffInword;
                                            }
                                            if (DiffInword < 60)
                                            {
                                                SIContainer[k] = ObjBoundingWord.RightWord;
                                                k++;
                                                TempLastWordX2Right = LastcharX2;
                                            }
                                            else
                                            {
                                                iWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount;
                                            }
                                        }
                                    }
                                }
                                else if (iLine > KeywordLineNo && iLine < KeywordLineNo + 10 && iLine < ObjMetaData.Page[intCurrPageNumber].LineCount)
                                {
                                    Icounter = 0;
                                    if (BottomwordNo != 0)
                                    {
                                        for (int iWord = BottomwordNo; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; iWord++)
                                        {
                                            clsCnCWord Word = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                            ObjBoundingWord = objCNC.GetBoundingWords(Word, ObjMetaData, intCurrPageNumber);
                                            int x1left = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Left;
                                            int Y1Top = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Top;
                                            if (Y1Top - TempstratY2Bottom > 50)
                                            {
                                                stopLoop = true;
                                                break;
                                            }
                                            if (iWord == BottomwordNo && Icounter < 1)
                                            {
                                                if (ObjBoundingWord.BottomWords != null)
                                                {
                                                    if (ObjBoundingWord.BottomWords.Length > 0)
                                                    {
                                                        BottomwordNo = (int)ObjBoundingWord.BottomWords[0].WordNumber;
                                                        TempstratY2Bottom = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Bottom;
                                                    }
                                                }

                                                if (!flagBotoomwordFound && iLine == KeywordLineNo + 1 && validdiff < 60)
                                                {
                                                    if (((x1left - NexrRightwordX1Left) >= -15 && (x1left - NexrRightwordX1Left) < 60))
                                                    {
                                                        validLine = true;
                                                        SIContainer[k] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                                        k++;
                                                        Icounter++;
                                                        TempLastWordX2Right = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Right;
                                                        KeywordsStartCharX1Left = NexrRightwordX1Left;
                                                    }
                                                    else
                                                    {
                                                        validLine = false;
                                                    }
                                                }
                                                else if (((x1left - KeywordsStartCharX1Left) >= -15 && (x1left - KeywordsStartCharX1Left) < 60))
                                                {
                                                    validLine = true;
                                                    SIContainer[k] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord];
                                                    k++;
                                                    Icounter++;
                                                    TempLastWordX2Right = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Right;
                                                }
                                                else
                                                {
                                                    validLine = false;
                                                }
                                            }

                                            if (validLine)
                                            {

                                                ObjBoundingWord = objCNC.GetBoundingWords(Word, ObjMetaData, intCurrPageNumber);
                                                if (iWord == BottomwordNo)
                                                {
                                                    if (ObjBoundingWord.BottomWords != null)
                                                    {
                                                        if (ObjBoundingWord.BottomWords.Length > 0)
                                                        {
                                                            BottomwordNo = (int)ObjBoundingWord.BottomWords[0].WordNumber;
                                                        }
                                                    }
                                                }

                                                if (ObjBoundingWord.RightWord != null)
                                                {
                                                    Firstcharx1 = ObjBoundingWord.RightWord.Left;
                                                    LastcharX2 = ObjBoundingWord.RightWord.Right;
                                                    DiffInword = Firstcharx1 - TempLastWordX2Right;

                                                    if (DiffInword < 60)
                                                    {
                                                        SIContainer[k] = ObjBoundingWord.RightWord;
                                                        k++;
                                                        TempLastWordX2Right = LastcharX2;
                                                    }
                                                    else
                                                    {
                                                        iWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount;
                                                    }
                                                }
                                            }
                                            else
                                            { break; }
                                        }
                                    }
                                }
                            }
                            SInstrction = objCNC.MergeWords(SIContainer);
                            bool Dbflag = MatchingSIwithDatabasevalues(SInstrction, intCurrPageNumber, PossibleWords, db_datatable);

                            if (!Dbflag && Function_Id != "Acc")
                            {
                                if (!string.IsNullOrEmpty(SInstrction.strWord))
                                {
                                    if (!FoundWord(PossibleWords, SInstrction))
                                    {
                                        cWord = ModifiedStringToCNC(SInstrction, "*****");
                                        cWord.Confidence = 20;
                                        if (!FoundWord(PossibleWords, cWord))
                                        {
                                            PossibleWords.Add(cWord);
                                        }
                                    }
                                }
                            }


                            if (PossibleWords.Count == 1 && string.IsNullOrEmpty(PossibleWords[0].strWord))
                            {
                                PossibleWords.Clear();
                            }

                        }
                        
                        if (Module1.IgnoreshipperExist == false && Function_Id == "Acc" && PossibleWords.Count == 0)
                        {
                            Module1.pgSearchInstrut = true;
                            Search_SI_WholePg(ObjMetaData, intCurrPageNumber, PossibleWords);
                        }
                        if (PossibleWords.Count >= 0)
                        {
                            returnZones = reurnPossibleWordData(ObjMetaData, returnZones, PossibleWords, ConfLevel, iRecursiveCallCount, intCurrPageNumber, cWord, 6);
                        }

                    }

                }
            }
            catch (Exception)
            {
                //throw;
            }
            finally
            {
                //ds.Clear();
                ds.Dispose();
                ds = null;
                //dsKeys.Clear();
                dsKeys.Dispose();
                dsKeys = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                foundRows = null;
                SIContainer = null;
                SInstrction = null;
                CombinedKeyword = null;
                ObjBoundingWord = null;
                ObjBoundingWord1 = null;
                ObjBoundingWordNext = null;
                sKeyword = null;
                objCNC.Dispose();
                cWord = null;
                ConfLevel = null;
                PossibleWords = null;
            }
            return returnZones;

        }
        #endregion

        #region Supporting Functions For Keyword Finding
        private static bool IsDate(Object obj)
        {
            string strDate = obj.ToString();
            DateTime dtout;
            try
            {
                if (DateTime.TryParse(strDate, out dtout))

                {
                    DateTime dt = DateTime.Parse(strDate);
                    if ((dt.Month != System.DateTime.Now.Month) || (dt.Day < 1 && dt.Day > 31) || dt.Year != System.DateTime.Now.Year)
                        return false;
                    else
                        return true;
                }
                else
                {
                    return false;
                }
            }
            catch
            {
                return false;
            }
        }

        private clsCnCWord GetNextJoinWord(clsCncMetaData oMeta, clsCnCWord Word, int PageNo, int LineNo)
        {
            clsCnCWord oWord = null;
            clsCnCWord sJoinWord = Word;
            string[] sStringArray = new string[oMeta.Page[PageNo].Line[LineNo].WordCount + 1];
            Regex regex = new Regex(@"^[\s]*(rs|inr|rs|US\.|£|€|US\$|\$)?[\s]*((([0-9]+)([\.|\,]([0-9]{1,2})))|((([0-9]{2}[\,|\.|'])+[0-9]{3})([\,|\.]([0-9]{1,2}))?)|(((([0-9]{1,2})[\,|\.|'])+([0-9]{3}))([\.|\,]([0-9]{1,2}))?)|((([0-9]{1,2})[\,|\.])+([0-9]{3}))|([0-9]{3}))[\s]*(\£|\€|euro|EUR|Eur|eur)?[\s]*$");
            RetStructIQSBR008 objRetStructIQSBR008 = new RetStructIQSBR008();
            clsCNCSBR objclsCNCSBR = new clsCNCSBR();
            string[] sSpecialCharacter = new string[] { "-" };
            int iCharacterIndex = 0;
            int iWordIndex;
            ClsCNC oClsCNC = new ClsCNC();
            clsCnCWord[] oWordsarr = new clsCnCWord[6];
            try
            {
                foreach (string Character in sSpecialCharacter)
                {
                    if (oMeta.Page[PageNo].Line[LineNo].strLine.ToUpper().Contains(Character))
                    {
                        clsCnCWord[] words = oMeta.Page[PageNo].Line[LineNo].Word;
                        for (int iIndex = 0, loopTo = words.Length - 1; iIndex <= loopTo; iIndex++)
                        {
                            if ((words[iIndex]) != null)
                            {
                                sStringArray[iIndex] = words[iIndex].strWord;
                            }
                        }

                        iCharacterIndex = Array.IndexOf(sStringArray, Character);
                        iWordIndex = Array.IndexOf(sStringArray, Word.strWord);
                        if (iCharacterIndex == iWordIndex + 1)
                        {
                            if (regex.IsMatch(sStringArray[iCharacterIndex + 1]))
                            {
                                if (IsDate(sStringArray[iCharacterIndex + 1]))
                                    return sJoinWord;
                                objRetStructIQSBR008 = objclsCNCSBR.IQSBR008(sStringArray[iCharacterIndex + 1], "E");
                                if (objRetStructIQSBR008.Status == "S")
                                    return sJoinWord;
                                oWordsarr[0] = sJoinWord;
                                oWordsarr[1] = words[iCharacterIndex];
                                oWordsarr[2] = words[iCharacterIndex + 1];
                                oWord = oClsCNC.MergeWords(oWordsarr);
                            }
                            //else
                            //{
                            //    oWord = sJoinWord;
                            //}
                        }
                        else
                        {
                            oWord = sJoinWord;
                        }
                    }
                    else
                    {
                        oWord = sJoinWord;
                    }
                }
            }
            catch (Exception ex)
            {
                //MsgBox(ex.Message, MsgBoxStyle.Critical, "F3:GetNextJoinWord");

            }

            return oWord;
        }

        private bool FoundWord(List<clsCnCWord> CNCWord, clsCnCWord NewWord)
        {
            try
            {
                foreach (clsCnCWord Word in CNCWord)
                {
                    if ((Word != null) && Word.strWord.Replace("#", "").Equals(NewWord.strWord.Replace("#", "")))
                    {
                        return true;
                    }
                }
            }
            catch (Exception ex)
            {
                // MessageBox.Show(ex.Message);
            }

            return false;
        }

        private List<clsCnCWord> GetTopThreeValue(List<clsCnCWord> ArrayList)
        {
            var lstTopThreeValue = new List<clsCnCWord>();
            int iCount = 0;
            try
            {
                foreach (clsCnCWord Word in ArrayList)
                {
                    if (iCount > 2)
                    {
                        break;
                    }

                    lstTopThreeValue.Add(Word);
                    iCount = iCount + 1;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }

            return lstTopThreeValue;
        }

        private RetStructF3 GetTotalInvoices_InvoiceNumber(clsCncMetaData ObjMetaData, int intCurrPageNo, List<clsCnCWord> FiltredInvoices, string strArgs)
        {
            var ObjRetStructF3 = new RetStructF3();
            var oFoundWord = new List<clsCnCWord>();
            var oClsCNC = new ClsCNC();
            string KeyWords = strArgs;
            string strFirst = string.Empty;
            string strSecond = string.Empty;
            clsCnCWord oTempword;
            for (int iLoop = 0, loopTo = FiltredInvoices.Count - 1; iLoop <= loopTo; iLoop++)
            {
                try
                {
                    string sLeftWords = string.Empty;
                    string sTopWords = string.Empty;
                    int iCount = 0;

                    // If Amount are Found More Than 3
                    if (oFoundWord.Count >= 3)
                    {
                        break;
                    }

                    // Check left words
                    string sLeft = string.Empty;
                    sLeftWords = GetLeftWords(ObjMetaData, FiltredInvoices[iLoop], FiltredInvoices[iLoop].PageNo, iCount, ref sLeft);
                    if (FindKeywordsInvoice_InvoiceNumber(KeyWords, sLeftWords))
                    {
                        iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(FiltredInvoices[iLoop], ObjMetaData, FiltredInvoices[iLoop].PageNo);
                        if (oBoundingWords.LeftWord is object)
                        {
                            //if (Information.IsNumeric(oBoundingWords.LeftWord.strWord) && oBoundingWords.LeftWord.strWord.Length <= 3)
                            if ((oBoundingWords.LeftWord.strWord.All(char.IsNumber)) && oBoundingWords.LeftWord.strWord.Length <= 3)
                            {
                                var oWord = new clsCnCWord();
                                var oWords = new clsCnCWord[3];
                                oWords[0] = oBoundingWords.LeftWord;
                                oWords[1] = FiltredInvoices[iLoop];
                                // To Check Space Separated words if they are at minimum space merge that words else tke right side word
                                if (oWords[1].Left - oWords[0].Right <= 40)
                                {
                                    oWord = oClsCNC.MergeWords(oWords);
                                    oFoundWord.Add(oWord);
                                    continue;
                                }
                                else
                                {
                                    oFoundWord.Add(oWords[1]);
                                    continue;
                                }
                            }
                            else
                            {
                                oFoundWord.Add(FiltredInvoices[iLoop]);
                                continue;
                            }
                        }
                        else
                        {
                            oFoundWord.Add(FiltredInvoices[iLoop]);
                            continue;
                        }
                    }

                    // Check Top words
                    sTopWords = GetTopWords(ObjMetaData, FiltredInvoices[iLoop], FiltredInvoices[iLoop].PageNo);
                    if (FindKeywordsInvoice_InvoiceNumber(KeyWords, sTopWords))
                    {
                        iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(FiltredInvoices[iLoop], ObjMetaData, FiltredInvoices[iLoop].PageNo);
                        if (oBoundingWords.LeftWord is object)
                        {
                            //if (Information.IsNumeric(oBoundingWords.LeftWord.strWord) && oBoundingWords.LeftWord.strWord.Length <= 3)
                            if ((oBoundingWords.LeftWord.strWord.All(char.IsNumber) && oBoundingWords.LeftWord.strWord.Length <= 3))
                            {
                                var oWord = new clsCnCWord();
                                var oWords = new clsCnCWord[3];
                                oWords[0] = oBoundingWords.LeftWord;
                                oWords[1] = FiltredInvoices[iLoop];
                                // To Check Space Separated words if they are at minimum space merge that words else tke right side word
                                if (oWords[1].Left - oWords[0].Right <= 40)
                                {
                                    oWord = oClsCNC.MergeWords(oWords);
                                    oFoundWord.Add(oWord);
                                }
                                else
                                {
                                    oFoundWord.Add(oWords[1]);
                                }
                            }
                            else
                            {
                                oFoundWord.Add(FiltredInvoices[iLoop]);
                            }
                        }
                        else
                        {
                            oFoundWord.Add(FiltredInvoices[iLoop]);
                        }
                    }
                }
                catch (Exception ex)
                {
                    // MsgBox(ex.Message, MsgBoxStyle.Critical, "iQGeneric-F3A17_IsValidDate")
                    iQPUBLIC.Common.ErrorLog(ex, "IQGeneric-F3C90_IsValidAmount");
                }
            }

            var ConfLevel = new List<int>();

            // *********************************************************************************************


            // To Decide Confidence of Amounts on Occurance
            ObjRetStructF3.Words = oFoundWord;
            ObjRetStructF3.NoOfieldsSuspects = oFoundWord.Count;
            switch (oFoundWord.Count)
            {
                case 1:
                    {
                        ConfLevel.Add(90);
                        ObjRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                        ObjRetStructF3.Flag = "Y";
                        ObjRetStructF3.ManualConfirmation = "Y";
                        ObjRetStructF3.Status = "S";
                        break;
                    }

                case 2:
                    {
                        int intConf = 80;
                        for (int intI = 0, loopTo1 = ObjRetStructF3.Words.Count - 1; intI <= loopTo1; intI++)
                        {
                            ObjRetStructF3.Words[intI].Confidence = intConf;
                            ConfLevel.Add(intConf);
                            intConf = intConf - 5;
                        }

                        ObjRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                        ObjRetStructF3.Flag = "Y";
                        ObjRetStructF3.ManualConfirmation = "Y";
                        ObjRetStructF3.Status = "S";
                        break;
                    }

                case 3:
                    {
                        int intConf = 85;
                        for (int intI = 0, loopTo2 = ObjRetStructF3.Words.Count - 1; intI <= loopTo2; intI++)
                        {
                            ObjRetStructF3.Words[intI].Confidence = intConf;
                            ConfLevel.Add(intConf);
                            intConf = intConf - 5;
                        }

                        ObjRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
                        ObjRetStructF3.Flag = "Y";
                        ObjRetStructF3.ManualConfirmation = "Y";
                        ObjRetStructF3.Status = "S";
                        break;
                    }

                default:
                    {
                        ObjRetStructF3.Status = "F";
                        break;
                    }
            }

            return ObjRetStructF3;
        }

        private string GetLeftWords(clsCncMetaData metaData, clsCnCWord targetWord, int pageNo, int iCount, ref string sLeftWords)
        {
            try
            {
                var oClsCNC = new ClsCNC();
                // Static sLeftWords As String = String.Empty

                if (targetWord == null)
                {
                    return sLeftWords;
                }

                iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(targetWord, metaData, targetWord.PageNo);
                if (oBoundingWords == null)
                {
                    return sLeftWords;
                }

                if (oBoundingWords.LeftWord == null)
                {
                    return sLeftWords;
                }

                // changes by manoj kad on 6 march 2015

                sLeftWords = oBoundingWords.LeftWord.strWord + " " + sLeftWords;

                // sLeftWords = sLeftWords & " " & oBoundingWords.LeftWord.strWord

                while (iCount < 8)
                {
                    iCount += 1;
                    GetLeftWords(metaData, oBoundingWords.LeftWord, oBoundingWords.LeftWord.PageNo, iCount, ref sLeftWords);
                }
            }
            catch (Exception ex)
            {
                // MsgBox(ex.Message, MsgBoxStyle.Critical, "iQGeneric-F3A17_GetLeftWords")
                iQPUBLIC.Common.ErrorLog(ex, "INVOICE : GetLeftWords");
            }

            return sLeftWords;
        }

        private bool FindKeywordsInvoice_InvoiceNumber(string Keywords, string ROIWords)
        {
            if (string.IsNullOrEmpty(ROIWords.Trim()))
                return false;
            var sKeyarray = new string[] { "Invoice no.", "Invoice #", "invoice Nurnber", "Document No", "Invoice no", "INVOIC E NO:", "No.", "NUMBER", "HP Sales order:", "Sales order :", "Lieferung :", "Shipper No:", "Delivery Note No.", "Delvery Note No.", "Lieferung Nummer:", "Lieferschein", "Your Reference:", "Ihre Bestellnr.", "bestellung", "Kundenauftragsnummer", "Ref. :", "Your order no. :", "Ihr Ansprechpartner :", "Rechnungsempfänger:", "Your Contact", "Lieferschein:", "IBAN :", "IBAN No:", "IBAN NO TL :", "IBAN code Austria", "MOMSNR.", "IBAN-Code =", "IBAN :", "IBAN", "IBAN:", "N° FACT.", "N° commande d`achat", "Ref Commande :", "Werkplaatsordernr", "Uw-bestelnummer", "Uw bestelling:", "Bestelhon", "Bestelbon", "Ihre Bestellnummer :", "Auftragsnr. Kunde", "Orderref.", "Ihr Auftrag", "Your orderno", "Deres reference", "Uw bestelbon :", "Commande de client", "Contract number:", "Référence", "ihre Referenznummer:", "Uw Ref. :", "Uw Ref.:", "Your Order", "Ihre Referenz:", "Ordre nummer :", "Work Order No :", "Order No.", "Bestellnummer:", "Order Numb", "ORDER No", "Bilagsnr. / Side :", "Nr.documento", "N° FACTURE", "FACTURE :", "Nr fv:", "Doc. No./Date", "Fattura #", "FACTURE N° :", "Numero della fattura:", "Beleg Nr;", "Document No.", "Number/Date", "Fatture N.", "Account Number", "INVolCE", "lnvoice #", "InvoiceNumber:", "Invoice#", "N. DOCUMENTO/DOCUMENTNO", "INVOICE", "Documento nr.:", "NUMERO DOCUMENTO", "Fattura n:", "Customer invoice", "Beleg:", "N° Fattura", "Nr. Documento", "Fattura nr", "N°Documento", "invoice #", "Invoice #", "invoice No:", "INVOICE NR.", "PF. INV. NO.:", "Fattura nr .", "Dokument Nr.", "Rechnungsnummer", "Beleg Nr.", "Invoice :", "invoice No", "invoice Number", "Invoice number :", "N° de facture :", "N° FA.", "Belegnr.:", "Numéro de facture", "Invoice No :", "InvoiceNO.", "Order No:", "Belegnummer", "Invoice Number:", "N. doc.", "Num doc / Date", "Invoice no.", "Fakturanummer", "Faktura", "Factura N°:", "INVOICE NR :", "Rechnung - Nr", "Fakt nr / Kundnr", "Faktura Nr.", "invoice No./Date", "Rechnungsnummer :", "Nummer / Datum", "Rechnung", "Rechnungs-Nr.:", "BELEG-NR.", "Rechnung Nr.", "Rechnungsnr.", "Rechnung:", "Rechnungsnummer:", "Numéro de facture:", "N° de facture", "Beleg-Nr. :", "FACTURE GLOBALE N°", "N° de la facture:", "FACTURE N°", "FACTURE (C) NO", "FACTURE  N':", "No. de la fact.", "Facture:", "Faktura VAT Nr", "Faktura VAT", "Nr Faktury", "Faktura nr", "Faktura", "Numer faktury :", "Numer faktury:", "FAKTURA / INVOICE:", "Rechn.Nr", "INVOICE N°", "Factuur :", "FAKTUUR", "Factuur", "Factuur nr :", "RECHNUNGS-NR.", "Factuurnr.", "FACTUURNUMMER", "factuurnr.", "Factuur:", "RECHNUNG :", "FACTUURNUMMER:", "Factuur nr.", "Fattura nr.", "Fattura -", "FATTURA N°", "Fatt. n°", "Fattura N.", "N. FATTURA", "Fakturanummer :", "Fakturanr.", "Faktura nr.:", "Faktura nummer", "Faktura nr. :", "Fakturanr./Invoice no.", "Faktura nr.", "Your purchase order :", "Your order no.", "Ihre Bestell-Nr. :", "Your order no.:", "Inköpsordernummer", "Ert ordernr", "Er referens", "Best:", "Ert ordernr", "Ert bestar", "Er referens/bestnr", "Ihre  Bestellnummer :", "ihre Bestell-Nr.", "IHRE AUFTRAGSNR :", "ihre Bestellung", "Ihre Bestell-Nr.:", "Kdnbestellnr. /cust. order no/n°de comm. du cl ent:", "Ihre Auftraos-Nr", "Bestelldaten: Nr.", "Auftragsdaten:", "Bestellung Nr.", "Externe Belegnummer", "Bestellung Nr.:", "Bestell Nr.", "BESTELNR.", "Uw ordernummer", "Uw Internet bestelling", "Referentienr.*", "IHRE BESTELLUNG :", "Bestelbon", "Uw referentie :", "Uw Bestelnummer", "Uw Order en ref.:", "Uw order nummer", "Vostro ordine nr.", "Vostro ordine nr.", "Votre no. de cmde:", "Deres reference :", "Deres reference rekv.nr.", "lhr Zeichen", "Bestell-Ref.", "Ihre  Bestellung:", "lhre Bestelinr.", "Ihre Referenz :", "Telefon", "fon", "Phone :", "Contact Telephone", "Tel.:", "Telefon:", "Tel.", "Telefon :", "Tel:", "Tel :", "Telefoon", "Phone:", "Sales Phone nr:", "Phone n° :", "Direct telefoonnr.  :", "doorkiesnummer", "Telephone;", "Telefon", "Telephone:", "Telefoonnummer :", "Telefoonnummer", "tlf", "Telephone No", "Téléphone", "Num. de téléphone", "Tfn:", "Tél. :", "Tél.", "Téléphone :", "Tél :", "Phone", "Téléphone:", "Pbone n° :", "Tél", "NIP", "PART.IVA", "P.iva", "Partita IVA", "IVA", "FISCALE:", "Ust.-IdNr.:", "USt.-ID-Nr.", "T.V.A", "TVA", "CVR nr.", "CVR-nr.", "CVR/SE nr.:", "CVR:", "TVA:", "B.T.W. NUMMER", "BTW nr.:", "unsere", "NIP:", "MOMSNR.", "Ust-Id ATU:", "UST-ID Nr./St.Nr", "USt-IdNr.", "Ihre-UID:", "IVA", "CVR nr.:", "ST-Id-Nr.:", "MOMSNR.", "Vostra P.IVA", "TVA", "Votre N° de TVA", "CVR-nr.:", "Customer VAT N°:", "VAT nr.:", "Momsreg.nr/VATnr:", "Momsreg.nr.", "VAT Nr/VAT No", "Momsreg.nr/VAT-nr:", "VAT no.", "Momsreg nr", "Vertr.Nr.", "Btw-nr.", "BTW:", "UST-ID NrJSt.Nr", "UST-ID Nr./St.Nr", "BTW N°", "BTW-nummer", "BTW", "VAT REG. NO.:", "Credit Invoice", "credit nota", "credit note freight", "credit note", "kreditnota nr.:", "DeblKred.", "entgeltminderung", "NOTA DEBITO", "Credit Number", "Credit memo Number", "Gutschrift", "Avoir", "Kredit nota/faktura", "nota de credito/abono", "nota de credit", "Buchunggutschift", "Stornorechnung", "Retourengutschrift", "Gutschein", "Korrekturrechnung", "Rechnungsstorno", "Abono", "Recticativa", "Warengutchrift", "Storno", "GUTSCHRIFT", "No:", "FATTURA", "Numero", "Jmoice", "Inv.No.  :", "Inv.No.:", "Our Order No" };
            var arrKeyword = Keywords.Split(',');
            for (int iLoop = 0, loopTo = sKeyarray.Length - 1; iLoop <= loopTo; iLoop++)
            {
                if (ROIWords.Trim().ToUpper().Contains(sKeyarray[iLoop].Trim().ToUpper()))
                {
                    return true;
                }
            }

            return false;
        }

        private string GetTopWords(clsCncMetaData metaData, clsCnCWord targetWord, int pageNo)
        {
            string sTopWords = string.Empty;
            try
            {
                var oClsCNC = new ClsCNC();
                if (targetWord == null)
                {
                    return sTopWords;
                }

                iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(targetWord, metaData, targetWord.PageNo);
                if (oBoundingWords == null)
                {
                    return sTopWords;
                }

                if (oBoundingWords.TopWords == null)
                {
                    return sTopWords;
                }

                for (int iLoop = 0, loopTo = oBoundingWords.TopWords.Length - 1; iLoop <= loopTo; iLoop++)
                    sTopWords = sTopWords + " " + oBoundingWords.TopWords[iLoop].strWord;
            }
            catch (Exception ex)
            {
                // MsgBox(ex.Message, MsgBoxStyle.Critical, "iQGeneric-F3A17_GetBottomWords")
                iQPUBLIC.Common.ErrorLog(ex, "INVOICE : GetBottomWords");
            }

            return sTopWords;
        }

        #endregion
        #region Supporting Total Weight and Pieces
        private void DbvalueToKeywordsearch(string dbValue, DataRow[] foundRows, string sKeyword, clsCncMetaData ObjMetaData, int intCurrPageNumber, List<clsCnCWord> PossibleWords, string IdFlag)
        {
            bool Falg_KeywordMatch = false;
            string keywordFound = "";
            int lineno = 0;
            int wordno = 0;
            string wordstring = "";
            int MaxLength = 0;
            try
            {
                if (!string.IsNullOrEmpty(dbValue))
                {
                    if (IdFlag == "TotalWeight")
                    {
                        MaxLength = 7;
                    }
                    else //if (IdFlag == "Pieces")
                    {
                        MaxLength = 5;
                    }
                    //if (foundRows.Length == 1)
                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        //pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);

                        for (int iWord = 1; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount; iWord++)
                        {
                            wordstring = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                            wordstring = Module1.RoundOff(wordstring);
                            if (wordstring == dbValue && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                            {
                                Falg_KeywordMatch = GetRightWord(dbValue, keywordFound, iWord, lineno, ObjMetaData, intCurrPageNumber, Falg_KeywordMatch);
                                if (Falg_KeywordMatch)
                                {
                                    //New added
                                    ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord = wordstring;
                                    // New added

                                    //PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]);
                                    if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]))
                                    {
                                        PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]);
                                    }
                                }
                            }
                        }
                    }
                    else if (foundRows.Length > 1)
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            //pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            for (int iWord = 1; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount; iWord++)
                            {
                                wordstring = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS", "").Replace("LB", "");
                                wordstring = Module1.RoundOff(wordstring);

                                if (wordstring == dbValue && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                                {
                                    Falg_KeywordMatch = GetRightWord(dbValue, keywordFound, iWord, lineno, ObjMetaData, intCurrPageNumber, Falg_KeywordMatch);
                                    if (Falg_KeywordMatch)
                                    {
                                        //New added
                                        ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord = wordstring;
                                        // New added

                                        if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]))
                                        {
                                            PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]);
                                        }

                                    }
                                }
                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            { }

        }
        private void KeywordTODbvaluesearch(string dbValue, DataRow[] foundRows, string sKeyword, clsCncMetaData ObjMetaData, int intCurrPageNumber, List<clsCnCWord> PossibleWords, string IdFlag)
        {
            // bool Falg_KeywordMatch = false;
            string keywordFound = "";
            int lineno = 0;
            int wordno = 0;
            string wordstring = "";
            int MaxLength = 0;
            int Newwordno = 0;
            try
            {
                if (!string.IsNullOrEmpty(dbValue))
                {
                    if (IdFlag == "TotalWeight")
                    {
                        MaxLength = 7;
                    }
                    else //if (IdFlag == "Pieces")
                    {
                        MaxLength = 5;
                    }
                    //if (foundRows.Length == 1)
                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        //pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        string[] splitKeword = keywordFound.Split(' ');

                        if (splitKeword.Length > 1)
                        {
                            Newwordno = wordno + splitKeword.Length - 1;
                        }
                        else
                        {
                            Newwordno = wordno;
                        }

                        for (int iWord = Newwordno; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount; iWord++)
                        {
                            wordstring = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                            wordstring = Module1.RoundOff(wordstring);
                            if (wordstring == dbValue && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                            {
                                
                                // new added
                                ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord = wordstring;
                                // new added

                                if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]))
                                {
                                    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]);
                                }
                                //}
                            }
                        }
                    }
                    else if (foundRows.Length > 1)
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            //pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            string[] splitKeword = keywordFound.Split(' ');
                            if (splitKeword.Length > 1)
                            {
                                Newwordno = wordno + splitKeword.Length - 1;
                            }
                            else
                            {
                                Newwordno = wordno;
                            }
                            for (int iWord = Newwordno; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount; iWord++)
                            {
                                wordstring = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS", "").Replace("LB", "");
                                wordstring = Module1.RoundOff(wordstring);

                                if (wordstring == dbValue && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                                {
                                    // new added
                                    ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord = wordstring;
                                    // new added

                                    if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]))
                                    {
                                        PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord]);
                                    }

                                    //}
                                }
                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            { }


        }
        private Dictionary<bool, bool> CheckUnitForWeight(ClsCNC objCNC, clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCWord cnCWord)
        {
            bool flagUnit = false;
            bool KgUnit = false;
            Dictionary<bool, bool> UnitWeight = new Dictionary<bool, bool>();
            string[] UnitArray = { "LBS", "LB", "KG" };
            clsCnCBoundingWords boundingWords = new clsCnCBoundingWords();
            UnitWeight.Add(flagUnit, KgUnit);
            //bool Allalphabate= cnCWord.strWord.All(char.IsLetter);
            //if (UnitArray.Any(cnCWord.strWord.ToUpper().Contains) && Allalphabate==false)
            if (UnitArray.Any(cnCWord.strWord.ToUpper().Contains))
            {
                UnitWeight.Clear();
                if ((cnCWord.strWord.Replace(".", "").ToUpper().Contains("KG")))
                {
                    KgUnit = true;
                }
                flagUnit = true;
                UnitWeight.Add(flagUnit, KgUnit);
            }
            else
            {

                boundingWords = objCNC.GetBoundingWords(cnCWord, ObjMetaData, intCurrPageNumber);
                if (boundingWords.RightWord != null)
                {
                    string cwrd = boundingWords.RightWord.strWord;

                    //if (UnitArray.Any(boundingWords.RightWord.strWord.Replace(".", "").ToUpper().Contains))
                    if (UnitArray.Any(cwrd.Replace(".", "").ToUpper().Contains))
                    {
                        UnitWeight.Clear();
                        if (cwrd.Replace(".", "").ToUpper() == "KG")
                        {
                            KgUnit = true;
                        }
                        flagUnit = true;
                        UnitWeight.Add(flagUnit, KgUnit);
                    }
                }
            }

            return UnitWeight;
        }
        private void SearchDataNextTwoLine(string dbValue, DataRow[] foundRows, string sKeyword, clsCncMetaData ObjMetaData, int intCurrPageNumber, List<clsCnCWord> PossibleWords, string IdFlag)
        {
            string keywordFound = "";
            int lineno = 0;
            int wordno = 0;
            string wordstring = "";
            int X1Left = 0;
            int WordX1Left = 0;
            int Diff = 0;
            int MaxLength = 0;
            try
            {
                if (!string.IsNullOrEmpty(dbValue))
                {
                    if (IdFlag == "TotalWeight")
                    {
                        MaxLength = 7;
                    }
                    else //if (IdFlag == "Pieces")
                    {
                        MaxLength = 5;
                    }
                    if (foundRows.Length == 1)
                    {
                        keywordFound = Convert.ToString(foundRows[0].ItemArray[0]);
                        //pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        X1Left = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno].Left;
                        for (int iLine = lineno + 1; iLine <= lineno + 2; iLine++)
                        {
                            if (iLine <= ObjMetaData.Page[intCurrPageNumber].LineCount)
                            {
                                for (int iWord = 1; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; iWord++)
                                {

                                    wordstring = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                    wordstring = Module1.RoundOff(wordstring);
                                    if (Module1.regexTW.IsMatch(wordstring) && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                                    {
                                        WordX1Left = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Left;

                                        Diff = X1Left - WordX1Left;
                                        if (Diff >= -160 && Diff < 160) // 60 changes to 150
                                        {
                                            //wordstring = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                            if (Module1.regexTW.IsMatch(wordstring))
                                            {
                                                wordstring = Module1.RoundOff(wordstring);
                                                if (wordstring == dbValue && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                                                {
                                                    //New added
                                                    ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].strWord = wordstring;
                                                    // New added

                                                    if (Keyword_value_Heightdiff(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno], ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]) == true)
                                                    {
                                                        if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]))
                                                        {
                                                            PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]);
                                                        }
                                                    }
                                                    ////if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]))
                                                    ////{
                                                    ////    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]);
                                                    ////}
                                                }

                                            }
                                        }
                                    }
                                }
                            }

                        }


                    }
                    else if (foundRows.Length > 1)
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                            //pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            X1Left = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno].Left;
                            for (int iLine = lineno + 1; iLine <= lineno + 2; iLine++)
                            {
                                if (iLine <= ObjMetaData.Page[intCurrPageNumber].LineCount)
                                {
                                    for (int iWord = 1; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; iWord++)
                                    {
                                        wordstring = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                        wordstring = Module1.RoundOff(wordstring);
                                        if (Module1.regexTW.IsMatch(wordstring) && wordstring.Length <= MaxLength)
                                        {
                                            WordX1Left = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].Left;

                                            Diff = X1Left - WordX1Left;
                                            if (Diff >= -160 && Diff < 160) // 60 changes to 160
                                            {
                                                //wordstring = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].strWord.ToUpper().Replace(",", "").Replace("[", "").Replace("]", "").Replace("(", "").Replace(")", "").Replace(">>", "").Replace("LBS.", "").Replace("LBS", "").Replace("LB", "");
                                                if (Module1.regexTW.IsMatch(wordstring))
                                                {
                                                    wordstring = Module1.RoundOff(wordstring);
                                                    if (wordstring == dbValue && wordstring.Length <= MaxLength && Convert.ToInt32(wordstring) > 0)
                                                    {
                                                        //New added
                                                        ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord].strWord = wordstring;
                                                        // New added
                                                        if (Keyword_value_Heightdiff(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno], ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]) == true)
                                                        {
                                                            if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]))
                                                            {
                                                                PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]);
                                                            }
                                                        }
                                                        ////if (!FoundWord(PossibleWords, ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]))
                                                        ////{
                                                        ////    PossibleWords.Add(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord]);
                                                        ////}
                                                    }

                                                }
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
            { }

        }
        private bool GetRightWord(string dbValue, string keywordFound, int wordno, int lineno, clsCncMetaData ObjMetaData, int intCurrPageNumber, bool Falg_KeywordMatch)
        {
            ClsCNC ObjCnc = new ClsCNC();
            clsCNCSBR ObjSbr = new clsCNCSBR();
            clsCnCBoundingWords ObjBoundingWords;
            string word = "";

            //int counter = 0;
            string[] splitKeyword = keywordFound.Split(' ');

            try
            {
                for (int iWord = wordno; iWord <= ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount; iWord++)
                {
                    word = ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord].strWord;

                    ObjBoundingWords = ObjCnc.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[iWord], ObjMetaData, intCurrPageNumber);
                    if (ObjBoundingWords == null)
                    {
                        return Falg_KeywordMatch;
                    }

                    if (ObjBoundingWords.RightWord == null)
                    {
                        return Falg_KeywordMatch;
                    }
                    if (ObjBoundingWords.RightWord != null)
                    {
                        int counter = 0;

                        for (int j = 0; j <= splitKeyword.Length - 1; j++)
                        {
                            if (ObjBoundingWords.RightWord != null)
                            {
                                if (ObjBoundingWords.RightWord.strWord == ":")
                                {
                                    ObjBoundingWords = ObjCnc.GetBoundingWords(ObjBoundingWords.RightWord, ObjMetaData, intCurrPageNumber);
                                }
                                if (ObjBoundingWords.RightWord == null)
                                {
                                    return Falg_KeywordMatch;
                                }
                                RetStructIQSBR050 Obj50 = ObjSbr.IQSBR050(ObjBoundingWords.RightWord.strWord, splitKeyword[j]);
                                if ((Obj50.Field1Length <= 5 && Obj50.CharChanged <= 1) || ((Obj50.Field1Length > 5 && Obj50.Field1Length <= 8) && Obj50.CharChanged <= 2) || ((Obj50.Field1Length > 8 && Obj50.Field1Length <= 11) && Obj50.CharChanged <= 3) || (Obj50.Field1Length >= 12 && Obj50.CharChanged <= 4))
                                {
                                    ObjBoundingWords = ObjCnc.GetBoundingWords(ObjBoundingWords.RightWord, ObjMetaData, intCurrPageNumber);
                                    counter++;
                                    if (counter == splitKeyword.Length)
                                    {

                                        Falg_KeywordMatch = true;

                                        iWord = ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount;
                                        break;
                                    }

                                }
                                //else
                                //{
                                //    iWord = ObjMetaData.Page[intCurrPageNumber].Line[lineno].WordCount;
                                //}
                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                ObjCnc.Dispose();
                ObjSbr = null;
                word = null;
            }



            return Falg_KeywordMatch;
        }
        private void GetLeftWord(string dbValue, string keywordFound, int wordno, int lineno, clsCncMetaData ObjMetaData, int intCurrPageNumber, List<clsCnCWord> PossibleWords)
        {
            ClsCNC ObjCnc = new ClsCNC();
            clsCNCSBR ObjSbr = new clsCNCSBR();
            clsCnCBoundingWords ObjBoundingWords;
            string wordstring = string.Empty;

            string[] splitKeyword = keywordFound.Split(' ');
            try
            {
                ObjBoundingWords = ObjCnc.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word[wordno], ObjMetaData, intCurrPageNumber);
                if (ObjBoundingWords.LeftWord != null)
                {
                    if (ObjBoundingWords.LeftWord != null)
                    {
                        if (ObjBoundingWords.LeftWord.strWord == ":" || ObjBoundingWords.LeftWord.strWord.ToUpper() == "LBS" || ObjBoundingWords.LeftWord.strWord.ToUpper() == "LB")
                        {
                            ObjBoundingWords = ObjCnc.GetBoundingWords(ObjBoundingWords.LeftWord, ObjMetaData, intCurrPageNumber);
                        }

                        if (ObjBoundingWords.LeftWord != null)
                        {
                            if (ObjBoundingWords.LeftWord.strWord.ToUpper().Contains("LB"))
                            {
                                wordstring = ObjBoundingWords.LeftWord.strWord.ToUpper().Replace("LBS", "").Replace("LB", "");
                            }
                            wordstring = ObjBoundingWords.LeftWord.strWord;
                            wordstring = Module1.RoundOff(wordstring);
                            if (Module1.regexTW.IsMatch(wordstring) && wordstring.Length <= 7 && Convert.ToInt32(wordstring) > 0)
                            {
                                PossibleWords.Add(ObjBoundingWords.LeftWord);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                ObjCnc.Dispose();
                ObjSbr = null;
                wordstring = null;
            }



        }
        #endregion

        #region Supporting RADF-MABD Date
        private void Matched_Date(clsCncMetaData ObjMetaData, int intCurrPageNumber, Hashtable ht, List<clsCnCWord> PossibleWords, string Pri_Date_Keywords, int LineNo, int WordNo, clsCnCWord CombinedKeyword)
        {
            ClsCNC ObjCnc = new ClsCNC();
            clsCNCSBR ObjSbr = new clsCNCSBR();
            //Regex myRegexMDY;
            //Regex myRegexDMY;
            //Regex myRegexYDM;
            //Regex myRegexYMD;
            //Regex myRegexJANDY;
            //Regex myRegexDJANY;
            //Regex myRegexYJAND;
            //Regex myRegexDJANYsuffix;
            //Regex myRegexJANDYsuffix;
            //Regex myRegexMDYws;
            //Regex myRegexDMYws;
            //Regex myRegexYMDws;
            //Regex myRegexYDMws;
            //myRegexMDY = new Regex(@"^[\s]*(0[1-9]|1[012]|[1-9])[\s]*[-/. ][\s]*(0[1-9]|[12][0-9]|3[01]|[1-9])[\s]*[- /. ][\s]*(20[0-4][0-9]|19[0-9][0-9]|[0-9]{2})[\s]*$");
            //myRegexJANDY = new Regex(@"[\s]*((jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)([,]?)|(january|february|march|april|may|june|july|august|september|october|november|december)([,]?)|(janvier|février|mars|avril|pouvoir|juin|juillet|août|septembre|octobre|novembre|décembre)([,]?)|(januar|februar|märz|kann|juni|juli|oktober|dezember)([,]?)|(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)([,]?)|(enero|febrero|marzo|abril|puede|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)([,]?)|(januari|februari|mars|kan|juni|juli|augusti)([,]?)|(tammikuu|helmikuu|maaliskuu|huhtikuu|saattaa|kesäkuu|heinäkuu|elokuu|syyskuu|lokakuu|marraskuu|joulukuu)([,]?))[\s]*[-/. ][\s]*((0[1-9]|[12][0-9]|3[01]|[1-9])|(0[1-9][,]|[12][0-9][,]|3[01][,]|[1-9][,]|))[\s]*[-/. ][\s]*((20[0-4][0-9]|19[2-9][0-9]|[0-9]{2}))[\s]*");
            //myRegexJANDYsuffix = new Regex(@"[\s]*((jan|feb|mar|apr|may|jun|jul|aug|sep|sept|oct|nov|dec)([,]?)|(january|february|march|april|may|june|july|august|september|october|november|december)([,]?)|(janvier|février|mars|avril|pouvoir|juin|juillet|août|septembre|octobre|novembre|décembre)([,]?)|(januar|februar|märz|kann|juni|juli|oktober|dezember)([,]?)|(gennaio|febbraio|marzo|aprile|maggio|giugno|luglio|agosto|settembre|ottobre|novembre|dicembre)([,]?)|(enero|febrero|marzo|abril|puede|junio|julio|agosto|septiembre|octubre|noviembre|diciembre)([,]?)|(januari|februari|mars|kan|juni|juli|augusti)([,]?)|(tammikuu|helmikuu|maaliskuu|huhtikuu|saattaa|kesäkuu|heinäkuu|elokuu|syyskuu|lokakuu|marraskuu|joulukuu)([,]?))[\s]*[-/. ][\s]*(0[4-9]th|1[0-9]th|[023][1]st|[02][2]nd|[02][3]rd|2[4-9]th|30th|[23][0]th|31st|1st|2nd|3rd|[4-9]th)([,]?)[\s]*[-/. ][\s]*((20[0-4][0-9]|19[5-9][0-9]|[0-9]{2} ))[\s]*");
            //myRegexMDYws = new Regex(@"^[\s]*(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])([0-9]{2}|19[2-9][0-9]|20[01][0-9])[\s]*$");
            //myRegexDMYws = new Regex(@"^[\s]*(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[012])([0-9]{2}|19[2-9][0-9]|20[01][0-9])[\s]*$");
            //myRegexYMDws = new Regex(@"^[\s]*([0-9]{2}|19[2-9][0-9]|20[01][0-9])(0[1-9]|1[012])(0[1-9]|[12][0-9]|3[01])[\s]*$");
            //myRegexYDMws = new Regex(@"^[\s]*([0-9]{2}|19[2-9][0-9]|20[01][0-9])(0[1-9]|[12][0-9]|3[01])(0[1-9]|1[012])[\s]*$");


            //Regex MDYwithComma = new Regex(@"^((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s(3[01]|[12][0-9]|0?[1-9]|[1-9]),\s((?:[0-9]{2})?[0-9]{2}))");
            Regex MDYwithComma = new Regex(@"((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s(3[01]|[12][0-9]|0?[1-9]|[1-9]),\s((?:[0-9]{2})?[0-9]{2}))");
            Regex DthMY = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])th\s(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec)\s((?:[0-9]{2})?[0-9]{2})");
            Regex DMY = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])(\s*)(January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)\s((?:[0-9]{2})?[0-9]{2})");
            Regex MDY = new Regex(@"((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)[ ](3[01]|[12][0-9]|0?[1-9]|[1-9])[ ](\d{4}))");
            // Regex MDY = new Regex(@"^((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|Jun|Jul|Aug|Sep|Oct|Nov|Dec|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)(\/|,||-|\s*)(\s*)?(3[01]|[12][0-9]|0?[1-9]|[1-9])(\/|,||-|\s*)(\s*)?((?:[0-9]{2})?[0-9]{2}))");
            //Regex DMYwithChar = new Regex(@"^(3[01]|[12][0-9]|0?[1-9]|[1-9])(\/|-|\s*)?((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec|jan|feb|mar|apr|jun|jul|aug|sep|oct|nov|dec)|(1[012]|0?[1-9]))(\/|,||-|\s*)(\s*)?(\d{4})(?:/[A-Za-z])?");
            clsCnCBoundingWords ObjBoundingWords;

            int KeywordX1Left = CombinedKeyword.Left;
            int keywordX2Right = CombinedKeyword.Right;

            int keywordTop = CombinedKeyword.Top;
            int keywordBottom = CombinedKeyword.Bottom;
            int KeywordHeight = keywordBottom - keywordTop;

            string date = string.Empty;
            try
            {
                for (int iline = LineNo; iline <= LineNo + 1; iline++)
                {
                    if (iline <= ObjMetaData.Page[intCurrPageNumber].LineCount)
                    {
                        if (MDYwithComma.IsMatch(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine))
                        {
                            Match Dates = MDYwithComma.Match(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine);
                            date = Dates.ToString();
                            ht.Add(iline, date);
                        }
                        if (DthMY.IsMatch(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine))
                        {
                            Match Dates = DthMY.Match(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine);
                            date = Dates.ToString();
                            ht.Add(iline, date);
                        }
                        if (DMY.IsMatch(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine))
                        {
                            Match Dates = DMY.Match(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine);
                            date = Dates.ToString();
                            ht.Add(iline, date);
                        }
                        if (MDY.IsMatch(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine))
                        {
                            Match Dates = MDY.Match(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine);
                            date = Dates.ToString();
                            ht.Add(iline, date);
                        }
                        //if (DMYwithChar.IsMatch(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine))
                        //{
                        //    Match Dates = DMYwithChar.Match(ObjMetaData.Page[intCurrPageNumber].Line[iline].strLine);
                        //    string date = Dates.ToString();
                        //    ht.Add(iline, date);
                        //}
                    }

                }

                foreach (int line in ht.Keys)
                {
                    string[] dates = ht[line].ToString().Split();

                    for (int iword = 1; iword <= ObjMetaData.Page[intCurrPageNumber].Line[line].WordCount; iword++)
                    {
                        if (ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword].strWord == dates[0])
                        {
                            ObjBoundingWords = ObjCnc.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword], ObjMetaData, intCurrPageNumber);
                            if (ObjBoundingWords.LeftWord != null)
                            {
                                // for (int i = 0; i < Pri_Date_Keywords.Length; i++)
                                //{
                                int counter = 0;
                                string[] Split_Pri = Pri_Date_Keywords.Split(' ');

                                for (int j = Split_Pri.Length - 1; j >= 0; j--)
                                {
                                    if (ObjBoundingWords.LeftWord != null)
                                    {
                                        if (ObjBoundingWords.LeftWord.strWord == ":")
                                        {
                                            ObjBoundingWords = ObjCnc.GetBoundingWords(ObjBoundingWords.LeftWord, ObjMetaData, intCurrPageNumber);
                                        }
                                        RetStructIQSBR050 Obj50 = ObjSbr.IQSBR050(ObjBoundingWords.LeftWord.strWord, Split_Pri[j]);
                                        if ((Obj50.Field1Length <= 5 && Obj50.CharChanged <= 1) || ((Obj50.Field1Length > 5 && Obj50.Field1Length <= 8) && Obj50.CharChanged <= 2) || ((Obj50.Field1Length > 8 && Obj50.Field1Length <= 11) && Obj50.CharChanged <= 3) || (Obj50.Field1Length >= 12 && Obj50.CharChanged <= 4))
                                        {

                                            ObjBoundingWords = ObjCnc.GetBoundingWords(ObjBoundingWords.LeftWord, ObjMetaData, intCurrPageNumber);
                                            counter++;
                                            if (counter == Split_Pri.Length)
                                            {
                                                clsCnCWord[] mergeWords = { ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword], ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword + 1], ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword + 2] };
                                                clsCnCWord Date = ObjCnc.MergeWords(mergeWords);
                                                Date = ModifiedStringToCNC(Date, Date.strWord.Replace(".", ""));
                                                Module1.ImageDate = DateTime.Parse(Date.strWord.ToString().Trim(), Module1.provider);
                                                if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                                                {
                                                    Module1.isFuturedate = true;
                                                }
                                                else
                                                {
                                                    Module1.isFuturedate = false;
                                                }
                                                if (!FoundWord(PossibleWords, Date))
                                                {
                                                    int TopValue = Date.Top;
                                                    int DistanceDiff = TopValue - keywordBottom;
                                                    if (DistanceDiff <= KeywordHeight * 3)
                                                    {
                                                        PossibleWords.Add(Date);
                                                    }

                                                    //PossibleWords.Add(Date);
                                                }

                                                // i = Pri_Date_Keywords.Length;
                                                j = 0; iword = ObjMetaData.Page[intCurrPageNumber].Line[line].WordCount;
                                                break;
                                                ////}

                                            }

                                        }
                                    }

                                }

                                //}

                            }
                            if (iword < ObjMetaData.Page[intCurrPageNumber].Line[line].WordCount - 1)
                            {
                                clsCnCWord[] mergeWord = { ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword], ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword + 1], ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword + 2] };

                                clsCnCWord DateWord = ObjCnc.MergeWords(mergeWord);
                                DateWord = ModifiedStringToCNC(DateWord, DateWord.strWord.Replace(".", ""));

                                int X1Left = DateWord.Left;
                                int X2Right = DateWord.Right;

                                if (X1Left > KeywordX1Left)//&& X2Right < keywordX2Right)
                                {
                                    Module1.ImageDate = DateTime.Parse(DateWord.strWord.ToString().Trim(), Module1.provider);
                                    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                                    {
                                        Module1.isFuturedate = true;
                                    }
                                    else
                                    {
                                        Module1.isFuturedate = false;
                                    }
                                    if (!FoundWord(PossibleWords, DateWord))
                                    {
                                        int TopValue = DateWord.Top;
                                        int DistanceDiff = TopValue - keywordBottom;
                                        if (DistanceDiff <= KeywordHeight * 3)
                                        {
                                            PossibleWords.Add(DateWord);
                                        }
                                        //PossibleWords.Add(DateWord);
                                    }
                                    //// }
                                }
                                ObjBoundingWords = ObjCnc.GetBoundingWords(DateWord, ObjMetaData, intCurrPageNumber);

                                if (ObjBoundingWords.TopWords != null)
                                {
                                    //for (int i = 0; i < Pri_Date_Keywords.Length; i++)
                                    // {
                                    string[] Split_Pri = Pri_Date_Keywords.Split();
                                    int t = 0, counter = 0;
                                    for (int j = 0; j < Split_Pri.Length; j++)
                                    {
                                        if (t < ObjBoundingWords.TopWords.Length)
                                        {
                                            if (ObjBoundingWords.TopWords[t].strWord == Split_Pri[j])
                                            {
                                                t++; counter++;
                                                if (counter == Split_Pri.Length)
                                                {
                                                    Module1.ImageDate = DateTime.Parse(DateWord.strWord.ToString().Trim(), Module1.provider);
                                                    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                                                    {
                                                        Module1.isFuturedate = true;
                                                    }
                                                    else
                                                    {
                                                        Module1.isFuturedate = false;
                                                    }
                                                    if (PossibleWords.Count > 0)
                                                    {
                                                        PossibleWords.Clear();
                                                    }
                                                    PossibleWords.Add(DateWord);

                                                    //i = Pri_Date_Keywords.Length;
                                                    j = Split_Pri.Length; iword = ObjMetaData.Page[intCurrPageNumber].Line[line].WordCount;

                                                    //// }
                                                }

                                            }
                                        }

                                    }
                                    // }
                                }
                            }
                        }
                    }
                    if (PossibleWords.Count > 0)
                    {
                        break;
                    }


                }
            }
            catch (Exception ex)
            { }
            finally
            {
                ht = null;
                ObjCnc = null;
                ObjSbr = null;
                date = null;
            }

        }
        private bool CheckMultiple_dateMatch2(clsCncMetaData ObjMetaData, int intCurrPageNo, int LineNo, Dictionary<int, clsCnCWord> RangeDateMatch, int KeywordNo)
        {
            string str = ObjMetaData.Page[intCurrPageNo].Line[LineNo].strLine;
            int m = 0;
            //str = str.Replace("~"," ");
            //Regex reg = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])?(\/|-|\s*)((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(1[012]|0?[1-9]))(\/|,||-|\s*)(\s*)?(\d{4})");//(?:/[A-Za-z])?");

            //Regex reg = new Regex(@"(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](((19|20)[0-9]{2})|(\d{2}))");//
            //// Regex reg = new Regex(@"(1[012]|0?[1-9])[- /.](3[01]|[12][0-9]|0?[1-9]|[1-9])[- /.](((19|20)[0-9]{2})|(\d{2}))");  27/08/2021

            Regex reg = new Regex(@"(1[012]|0?[1-9])[- /.](3[01]|[12][0-9]|0?[1-9]|[1-9])[- /.](((19|20)[0-9]{2}))");

            ////Regex regYMD = new Regex(@"(((19|20)[0-9]{2})|(\d{2}))[,-/.](1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])");
            Regex regYMD = new Regex(@"(((19|20)[0-9]{2}))[,-/.](1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])");

           
            if (reg.Matches(str).Count == 2)
            {
                m = reg.Matches(str).Count;
            }
            else if (regYMD.Matches(str).Count == 2)
            {
                m = regYMD.Matches(str).Count;
                // reg = new Regex(@"(((19|20)[0-9]{2})|(\d{2}))[,-/.](1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])");
                reg = new Regex(@"(((19|20)[0-9]{2}))[,-/.](1[012]|0?[1-9])[,-/.](3[01]|[12][0-9]|0?[1-9]|[1-9])");
            }
            //int m = reg.Matches(str).Count;
            int counter = 0;
            bool isdaterange = false;
            bool isdaterangevalidLine = false;
            clsCnCWord compairvalue = new clsCnCWord();
            clsCnCWord ModifiedCword = new clsCnCWord();
            //clsCnCWord date= new clsCnCWord();
            try
            {
                if (m == 2)
                {

                    foreach (Match match in reg.Matches(str))
                    {
                        counter++;
                        if (counter == 1)
                        {
                            if (match.Success)
                            {
                                for (int iword = 1; iword <= ObjMetaData.Page[intCurrPageNo].Line[LineNo].WordCount; iword++)
                                {
                                    if (ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword].strWord.Contains(match.Value.ToString()))
                                    {
                                        if ((int)ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword].WordNumber < KeywordNo)
                                        {
                                            isdaterange = false;
                                            isdaterangevalidLine = false;
                                            break;
                                        }
                                        else
                                        {
                                            isdaterangevalidLine = true;
                                            break;
                                        }
                                    }
                                }
                            }
                        }
                        if (counter == 2)
                        {
                            if (isdaterangevalidLine)
                            {
                                if (match.Success)
                                {
                                    for (int iword = 1; iword <= ObjMetaData.Page[intCurrPageNo].Line[LineNo].WordCount; iword++)
                                    {
                                        if (ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword].strWord.Contains(match.Value.ToString()))
                                        {
                                            if (RangeDateMatch.Count > 0)
                                            {
                                                if (RangeDateMatch.ContainsKey(LineNo))
                                                {
                                                    compairvalue = RangeDateMatch[LineNo];
                                                    if (compairvalue.strWord == ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword].strWord.Replace("*", ""))
                                                    {
                                                        RangeDateMatch.Clear();
                                                    }

                                                }
                                            }
                                            ModifiedCword = ModifiedStringToCNC(ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword], match.Value.ToString());
                                            try
                                            {
                                                if (CheckForValidDate(ModifiedCword.strWord))
                                                {
                                                    RangeDateMatch.Add(LineNo, ModifiedCword);
                                                }
                                                ////RangeDateMatch.Add(LineNo, ModifiedCword);   27/08/2021
                                            }
                                            catch (Exception ex)
                                            { }

                                            ////RangeDateMatch.Add(LineNo, ModifiedCword);
                                            //RangeDateMatch.Add(LineNo, ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword]);
                                            //Module1.MABDDate = ModifiedStringToCNC(ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[iword], match.Value.ToString());
                                        }
                                    }
                                    //Module1.MABDDate = StringToCNC(pageNo, match.Value.ToString());

                                    ///isdaterange = true; 27/08/2021

                                    if (RangeDateMatch.Count > 0)
                                    {
                                        isdaterange = true;
                                    }
                                    else
                                    {
                                        isdaterange = false;
                                    }
                                }
                            }

                        }

                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                compairvalue = null;
                ModifiedCword = null;
            }

            return isdaterange;
        }
        // for MABD and Dlv "O" type
        public bool Common_KeywordMatch(string Keyword, string[] Arry)
        {
            bool isKeywordFound = false;
            for (int iarr = 0; iarr < Arry.Length; iarr++)
            {
                if (Keyword.ToUpper() == Arry[iarr].ToUpper())
                {
                    isKeywordFound = true;
                }
            }
            return isKeywordFound;
        }
        private bool CheckForFutureDate(string strData)
        {
            bool isFutureDate = false;
            string DD = string.Empty;
            string MM = string.Empty;
            string YY = string.Empty;
            try
            {
                RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(strData, "E");
                if (iQSBR008.Status == "S")
                {
                    Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                    {
                        isFutureDate = true;
                    }

                }
                else
                {
                    if (Module1.YMDwithoutSpace.IsMatch(strData))
                    {
                        Match Dates = Module1.YMDwithoutSpace.Match(strData);
                        string date = Dates.ToString();

                        YY = date.Substring(0, 4);
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
                        strData = MM + "/" + DD + "/" + YY;
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }

                    }
                    ////else if (Module1.DMYwithChar.IsMatch(strData))
                    ////{
                    ////    Match Dates = Module1.DMYwithChar.Match(strData);
                    ////    string date = Dates.ToString();
                    ////    Module1.ImageDate = DateTime.Parse(date.ToString(), Module1.provider);
                    ////    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                    ////    {
                    ////        isFutureDate = true;
                    ////    }

                    ////}
                    else if (Module1.MDYwithChar.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithChar.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        Module1.ImageDate = DateTime.Parse(date.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else if (Module1.MDYwithComma.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithComma.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace(",", "-").Replace(" ", "-");
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else if (Module1.DMMMYwithChar.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithComma.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace("/", "-").Replace(",", "-").Replace(".", "-").Replace(" ", "-");
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else if (Module1.YMD.IsMatch(strData))
                    {
                        Match Dates = Module1.YMD.Match(strData);
                        string date = Dates.ToString();

                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace("/", "").Replace(",", "").Replace(".", "").Replace(" ", "").Replace("-", "");
                        YY = strData.Substring(0, 4);
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
                        strData = MM + "/" + DD + "/" + YY;
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else
                    {
                        isFutureDate = false;
                    }

                }
            }
            catch (Exception ex)
            {
                isFutureDate = false;
                strData = "";
            }


            return isFutureDate;
        }
        private bool CheckForValidDate(string strData)
        {
            bool isValidDate = false;
            string DD = string.Empty;
            string MM = string.Empty;
            string YY = string.Empty;
            try
            {
                RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(strData, "E");
                if (iQSBR008.Status == "S")
                {
                    Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                    if (Module1.ImageDate.Year == Module1.ToadysDate.AddYears(-1).Year || Module1.ImageDate.Year == Module1.ToadysDate.Year || Module1.ImageDate.Year == Module1.ToadysDate.AddYears(1).Year)
                    {
                        isValidDate = true;
                    }

                }
                else
                {
                    if (Module1.YMDwithoutSpace.IsMatch(strData))
                    {
                        Match Dates = Module1.YMDwithoutSpace.Match(strData);
                        string date = Dates.ToString();

                        YY = date.Substring(0, 4);
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
                        strData = MM + "/" + DD + "/" + YY;
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ImageDate.Year == Module1.ToadysDate.AddYears(-1).Year || Module1.ImageDate.Year == Module1.ToadysDate.Year || Module1.ImageDate.Year == Module1.ToadysDate.AddYears(1).Year)
                        {
                            isValidDate = true;
                        }

                    }
                    ////else if (Module1.DMYwithChar.IsMatch(strData))
                    ////{
                    ////    Match Dates = Module1.DMYwithChar.Match(strData);
                    ////    string date = Dates.ToString();
                    ////    Module1.ImageDate = DateTime.Parse(date.ToString(), Module1.provider);
                    ////    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                    ////    {
                    ////        isFutureDate = true;
                    ////    }

                    ////}
                    else if (Module1.MDYwithChar.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithChar.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        Module1.ImageDate = DateTime.Parse(date.ToString(), Module1.provider);
                        if (Module1.ImageDate.Year == Module1.ToadysDate.AddYears(-1).Year || Module1.ImageDate.Year == Module1.ToadysDate.Year || Module1.ImageDate.Year == Module1.ToadysDate.AddYears(1).Year)
                        {
                            isValidDate = true;
                        }
                    }
                    else if (Module1.MDYwithComma.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithComma.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace(",", "-").Replace(" ", "-");
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ImageDate.Year == Module1.ToadysDate.AddYears(-1).Year || Module1.ImageDate.Year == Module1.ToadysDate.Year || Module1.ImageDate.Year == Module1.ToadysDate.AddYears(1).Year)
                        {
                            isValidDate = true;
                        }
                    }
                    else if (Module1.DMMMYwithChar.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithComma.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace("/", "-").Replace(",", "-").Replace(".", "-").Replace(" ", "-");
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ImageDate.Year == Module1.ToadysDate.AddYears(-1).Year || Module1.ImageDate.Year == Module1.ToadysDate.Year || Module1.ImageDate.Year == Module1.ToadysDate.AddYears(1).Year)
                        {
                            isValidDate = true;
                        }
                    }
                    else if (Module1.YMD.IsMatch(strData))
                    {
                        Match Dates = Module1.YMD.Match(strData);
                        string date = Dates.ToString();

                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace("/", "").Replace(",", "").Replace(".", "").Replace(" ", "").Replace("-", "");
                        YY = strData.Substring(0, 4);
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
                        strData = MM + "/" + DD + "/" + YY;
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ImageDate.Year == Module1.ToadysDate.AddYears(-1).Year || Module1.ImageDate.Year == Module1.ToadysDate.Year || Module1.ImageDate.Year == Module1.ToadysDate.AddYears(1).Year)
                        {
                            isValidDate = true;
                        }
                    }
                    else
                    {
                        isValidDate = false;
                    }

                }
            }
            catch (Exception ex)
            {
                isValidDate = false;
                strData = "";
            }


            return isValidDate;
        }
        private bool CheckForMatchDateReg(string strData)
        {
            bool isFutureDate = false;
            string DD = string.Empty;
            string MM = string.Empty;
            string YY = string.Empty;
            try
            {
                RetStructIQSBR008 iQSBR008 = Module1.ObjSbr.IQSBR008(strData, "E");
                if (iQSBR008.Status == "S")
                {
                    Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                    {
                        isFutureDate = true;
                    }

                }
                else
                {
                    if (Module1.YMDwithoutSpace.IsMatch(strData))
                    {
                        Match Dates = Module1.YMDwithoutSpace.Match(strData);
                        string date = Dates.ToString();

                        YY = date.Substring(0, 4);
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
                        strData = MM + "/" + DD + "/" + YY;
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }

                    }
                    ////else if (Module1.DMYwithChar.IsMatch(strData))
                    ////{
                    ////    Match Dates = Module1.DMYwithChar.Match(strData);
                    ////    string date = Dates.ToString();
                    ////    Module1.ImageDate = DateTime.Parse(date.ToString(), Module1.provider);
                    ////    if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                    ////    {
                    ////        isFutureDate = true;
                    ////    }

                    ////}
                    else if (Module1.MDYwithChar.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithChar.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        Module1.ImageDate = DateTime.Parse(date.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else if (Module1.MDYwithComma.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithComma.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace(",", "-").Replace(" ", "-");
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else if (Module1.DMMMYwithChar.IsMatch(strData))
                    {
                        Match Dates = Module1.MDYwithComma.Match(strData);
                        string date = Dates.ToString();
                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace("/", "-").Replace(",", "-").Replace(".", "-").Replace(" ", "-");
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else if (Module1.YMD.IsMatch(strData))
                    {
                        Match Dates = Module1.YMD.Match(strData);
                        string date = Dates.ToString();

                        strData = strData.Substring(0, date.Length);
                        strData = strData.Replace("/", "").Replace(",", "").Replace(".", "").Replace(" ", "").Replace("-", "");
                        YY = strData.Substring(0, 4);
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
                        strData = MM + "/" + DD + "/" + YY;
                        Module1.ImageDate = DateTime.Parse(strData.ToString(), Module1.provider);
                        if (Module1.ToadysDate.Date <= Module1.ImageDate.Date)
                        {
                            isFutureDate = true;
                        }
                    }
                    else
                    {
                        isFutureDate = false;
                    }

                }
            }
            catch (Exception ex)
            {
                isFutureDate = false;
                strData = "";
            }


            return isFutureDate;
        }
        #endregion

        #region Supporting SI
        //private bool Sir_MatchingSIwithDatabasevalues(clsCnCLine clsLine, int intCurrPageNumber, List<clsCnCWord> PossibleWords, DataTable dt)
        //{
        //    string dbSI = string.Empty;
        //    string dbSICode = string.Empty;
        //    string Strword = clsLine.strLine.Replace(" ", "");
        //    Strword = clsLine.strLine.Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "");
        //    Strword = Strword.ToUpper();
        //    string[] DbSIArray = new string[100];
        //    string retundatavalue = "";
        //    string MergdataCode = "";
        //    string tildachk = "";
        //    string datacopy = "";
        //    bool matchFound = false;
        //    clsCnCWord clsDbdata = new clsCnCWord();
        //    List<string> Remarks = new List<string>();
        //    List<string> Flags = new List<string>();
        //    string delimiter = "~";
        //    string RemarksList = string.Empty;
        //    string FlagsList = string.Empty;
        //    DataTable Tempdt = dt.Copy();

        //    //For same line instruction remove
        //    //int LineNo = clsWord.LineNo;
        //    //int WordNo = (Int32)clsWord.WordNumber;
        //    // end

        //    try
        //    {
        //        if (!string.IsNullOrEmpty(Strword))
        //        {
        //            foreach (DataRow dr in Tempdt.Rows)
        //            {
        //                if (Tempdt.Columns.Contains("Description"))
        //                {
        //                    dbSI = dr["Description"].ToString().Replace("  ", " ").Trim();
        //                }
        //                else
        //                {
        //                    dbSI = dr["Instruction"].ToString().Replace("  ", " ").Trim();
        //                }

        //                // string[] splitDbSI = dbSI.Split(' ');

        //                datacopy = dbSI.ToUpper().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "").Trim();
        //                dbSICode = dr["Code"].ToString().Trim();
        //                if (Strword.Contains(datacopy))
        //                {
        //                    //For adding distinct values for common instructions
        //                    bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => datacopy == row.Field<String>("Instruction"));
        //                    bool containsCode = Module1.ComnInstrConatainer.AsEnumerable().Any(row => dbSICode == row.Field<String>("Code"));
        //                    if (containsCode == false)
        //                    {
        //                        if (!contains)
        //                        {

        //                            Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
        //                            //Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, splitDbSI.Length);

        //                            MergdataCode = dbSI.Trim() + "#" + dbSICode;
        //                            retundatavalue = retundatavalue + "~" + MergdataCode;
        //                            tildachk = retundatavalue.Substring(0, 1);
        //                            if (tildachk == "~")
        //                            {
        //                                retundatavalue = retundatavalue.Remove(0, 1);
        //                            }
        //                            Remarks.Add("2F");
        //                            Flags.Add("1");
        //                        }
        //                    }

        //                }
        //                else
        //                {
        //                    RetStructIQSBR050 Obj50 = Module1.ObjSbr.IQSBR050(datacopy, Strword);
        //                    if (Obj50.PercentageMatch > 80)
        //                    {
        //                        //For adding distinct values for common instructions
        //                        bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => datacopy == row.Field<String>("Instruction"));
        //                        bool containsCode = Module1.ComnInstrConatainer.AsEnumerable().Any(row => dbSICode == row.Field<String>("Code"));
        //                        if (containsCode == false)
        //                        {
        //                            if (!contains)
        //                            {
        //                                Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
        //                                //Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, splitDbSI.Length);

        //                                MergdataCode = dbSI.Trim() + "#" + dbSICode;
        //                                retundatavalue = retundatavalue + "~" + MergdataCode;
        //                                tildachk = retundatavalue.Substring(0, 1);
        //                                if (tildachk == "~")
        //                                {
        //                                    retundatavalue = retundatavalue.Remove(0, 1);
        //                                }
        //                                Remarks.Add("2F");
        //                                Flags.Add("1");
        //                            }
        //                        }

        //                    }
        //                }
        //            }
        //        }

        //        if (!string.IsNullOrEmpty(retundatavalue))
        //        {
        //            matchFound = true;
        //            clsDbdata = ModifiedStringToCNC(clsLine, retundatavalue.ToUpper());

        //            PossibleWords.Add(clsDbdata);
        //            if (Remarks.Count > 1)
        //            {
        //                RemarksList = string.Join(delimiter, Remarks.ToArray());
        //                FlagsList = string.Join(delimiter, Flags.ToArray());
        //            }
        //            else
        //            {
        //                RemarksList = Remarks[0];
        //                FlagsList = Flags[0];
        //            }

        //            // PossibleWords[0].Flag = "1";
        //            //if (PossibleWords.Count==1)
        //            //{
        //            //    PossibleWords[0].Flag = FlagsList;
        //            //    PossibleWords[0].Remarks = RemarksList;
        //            //}
        //            if (PossibleWords.Count >= 1)
        //            {
        //                PossibleWords[PossibleWords.Count - 1].Flag = FlagsList;
        //                PossibleWords[PossibleWords.Count - 1].Remarks = RemarksList;
        //            }

        //        }
        //    }
        //    catch (Exception ex)
        //    { }
        //    finally
        //    {
        //        Tempdt.Dispose();
        //        Remarks = null;
        //        Flags = null;
        //    }
        //    return matchFound;
        //}
        //private bool New_MatchingSIwithDatabasevalues(clsCnCWord clsWord, int intCurrPageNumber, List<clsCnCWord> PossibleWords, DataTable dt)
        //{
        //    string dbSI = string.Empty;
        //    string dbSICode = string.Empty;
        //    string Strword = clsWord.strWord.Replace(" ", "");
        //    Strword = clsWord.strWord.Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "");
        //    Strword = Strword.ToUpper();
        //    string[] DbSIArray = new string[100];
        //    string retundatavalue = "";
        //    string MergdataCode = "";
        //    string tildachk = "";
        //    string datacopy = "";
        //    bool matchFound = false;
        //    clsCnCWord clsDbdata = new clsCnCWord();
        //    List<string> Remarks = new List<string>();
        //    List<string> Flags = new List<string>();
        //    string delimiter = "~";
        //    string RemarksList = string.Empty;
        //    string FlagsList = string.Empty;
        //    DataTable Tempdt = dt.Copy();

        //    //For same line instruction remove
        //    int LineNo = clsWord.LineNo;
        //    int WordNo = (Int32)clsWord.WordNumber;
        //    // end

        //    try
        //    {
        //        if (!string.IsNullOrEmpty(Strword))
        //        {
        //            foreach (DataRow dr in Tempdt.Rows)
        //            {
        //                if (Tempdt.Columns.Contains("Description"))
        //                {
        //                    dbSI = dr["Description"].ToString().Replace("  ", " ").Trim();
        //                }
        //                else
        //                {
        //                    dbSI = dr["Instruction"].ToString().Replace("  ", " ").Trim();
        //                }

        //                string[] splitDbSI = dbSI.Split(' ');

        //                datacopy = dbSI.ToUpper().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "").Trim();
        //                dbSICode = dr["Code"].ToString().Trim();
        //                if (Strword.Contains(datacopy))
        //                {
        //                    //For adding distinct values for common instructions
        //                    bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => datacopy == row.Field<String>("Instruction"));
        //                    bool containsCode = Module1.ComnInstrConatainer.AsEnumerable().Any(row => dbSICode == row.Field<String>("Code"));
        //                    if (containsCode == false)
        //                    {
        //                        if (!contains)
        //                        {

        //                            //Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
        //                            Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, dbSI, splitDbSI.Length);

        //                            ////MergdataCode = dbSI.Trim() + "#" + dbSICode;
        //                            ////retundatavalue = retundatavalue + "~" + MergdataCode;
        //                            ////tildachk = retundatavalue.Substring(0, 1);
        //                            ////if (tildachk == "~")
        //                            ////{
        //                            ////    retundatavalue = retundatavalue.Remove(0, 1);
        //                            ////}
        //                            ////Remarks.Add("2F");
        //                            ////Flags.Add("1");
        //                        }
        //                    }

        //                }
        //                else
        //                {
        //                    RetStructIQSBR050 Obj50 = Module1.ObjSbr.IQSBR050(datacopy, Strword);
        //                    if (Obj50.PercentageMatch > 80)
        //                    {
        //                        //For adding distinct values for common instructions
        //                        bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => datacopy == row.Field<String>("Instruction"));
        //                        bool containsCode = Module1.ComnInstrConatainer.AsEnumerable().Any(row => dbSICode == row.Field<String>("Code"));
        //                        if (containsCode == false)
        //                        {
        //                            if (!contains)
        //                            {
        //                                //Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
        //                                Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, dbSI, splitDbSI.Length);

        //                                ////MergdataCode = dbSI.Trim() + "#" + dbSICode;
        //                                ////retundatavalue = retundatavalue + "~" + MergdataCode;
        //                                ////tildachk = retundatavalue.Substring(0, 1);
        //                                ////if (tildachk == "~")
        //                                ////{
        //                                ////    retundatavalue = retundatavalue.Remove(0, 1);
        //                                ////}
        //                                ////Remarks.Add("2F");
        //                                ////Flags.Add("1");
        //                            }
        //                        }

        //                    }
        //                }
        //            }
        //        }


        //        ////if (!string.IsNullOrEmpty(retundatavalue))
        //        ////{
        //        ////    matchFound = true;
        //        ////    clsDbdata = ModifiedStringToCNC(clsWord, retundatavalue.ToUpper());
        //        ////    PossibleWords.Add(clsDbdata);
        //        ////    if (Remarks.Count > 1)
        //        ////    {
        //        ////        RemarksList = string.Join(delimiter, Remarks.ToArray());
        //        ////        FlagsList = string.Join(delimiter, Flags.ToArray());
        //        ////    }
        //        ////    else
        //        ////    {
        //        ////        RemarksList = Remarks[0];
        //        ////        FlagsList = Flags[0];
        //        ////    }
        //        ////    if (PossibleWords.Count >= 1)
        //        ////    {
        //        ////        PossibleWords[PossibleWords.Count - 1].Flag = FlagsList;
        //        ////        PossibleWords[PossibleWords.Count - 1].Remarks = RemarksList;
        //        ////    }

        //        ////}
        //    }
        //    catch (Exception ex)
        //    { }
        //    finally
        //    {
        //        Tempdt.Dispose();
        //        Remarks = null;
        //        Flags = null;
        //    }
        //    return matchFound;
        //}
        private void Search_SI_WholePg(clsCncMetaData ObjMetaData, int intCurrPageNumber, List<clsCnCWord> PossibleWords)
        {
            string strline;
            clsCnCLine cncstrline = new clsCnCLine();
            clsCnCWord Linetoword = new clsCnCWord();
            try
            {
                for (int iLine = 1; iLine <= ObjMetaData.Page[intCurrPageNumber].LineCount; iLine++)
                {
                    //cncstrline = ObjMetaData.Page[intCurrPageNumber].Line[iLine];
                    Linetoword = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[1];
                    strline = ObjMetaData.Page[intCurrPageNumber].Line[iLine].strLine.Replace("  ", "").Replace(" ", "");
                    //Linetoword = strLineToCNCword(Linetoword, strline);
                    Linetoword = ModifiedStringToCNC(Linetoword, strline);
                    MatchingSIwithDatabasevalues(Linetoword, intCurrPageNumber, PossibleWords, Module1.dtAccessorials);
                }
            }
            catch (Exception ex)
            { }

        }
        private clsCnCWord MultipleSI(List<clsCnCWord> PossibleWords)
        {
            string ConcatSI = string.Empty;
            string flag = string.Empty;
            string Remarks = string.Empty;
            clsCnCWord tempcnc = PossibleWords[0];
            for (int loop = 0; loop < PossibleWords.Count; loop++)
            {
                ConcatSI = ConcatSI + "~" + PossibleWords[loop].strWord;
                flag = flag + "~" + PossibleWords[loop].Flag;
                Remarks = Remarks + "~" + PossibleWords[loop].Remarks;
            }
            //tildachk = retundatavalue.Substring(0, 1);
            if (ConcatSI.Substring(0, 1).Trim() == "~")
            {
                ConcatSI = ConcatSI.Remove(0, 1);
            }
            if (flag.Substring(0, 1).Trim() == "~")
            {
                flag = flag.Remove(0, 1);
            }
            if (Remarks.Substring(0, 1).Trim() == "~")
            {
                Remarks = Remarks.Remove(0, 1);
            }
            //PossibleWords.Clear();
            tempcnc = ModifiedStringToCNC(tempcnc, ConcatSI);
            tempcnc.Flag = flag;
            tempcnc.Remarks = Remarks;
            //PossibleWords.Add(tempcnc);
            return tempcnc;
        }
        private Boolean IgnoreShipperMatching(string Imgshippername)
        {
            bool shippermatch = false;
            //clsCNCSBR ObjSbr = new clsCNCSBR();
            string db_shippername = string.Empty;
            if (!string.IsNullOrEmpty(Imgshippername))
            {
                if (Module1.Dt_IgnoreShipperDetails.Rows.Count > 0)
                {
                    foreach (DataRow dr in Module1.Dt_IgnoreShipperDetails.Rows)
                    {
                        db_shippername = dr["ShipperName"].ToString().Trim();
                        RetStructIQSBR050 Obj50 = Module1.ObjSbr.IQSBR050(Imgshippername, db_shippername);
                        if (Obj50.PercentageMatch >= 80)
                        {
                            shippermatch = true;
                            break;
                        }
                    }
                }
            }
            return shippermatch;
        }
        private DataTable Addrowintodttable(DataTable dt, string Matchrow, string MatchCode)
        {
            DataRow Dr = dt.NewRow();
            Dr["Instruction"] = Matchrow;
            Dr["Code"] = MatchCode;
            dt.Rows.Add(Dr);
            return dt;
        }
        private DataTable New_Addrowintodttable(DataTable dt, string Matchrow, string MatchCode, int LineNo, int WordNo, string DbSI, int SILength)
        {
            DataRow Dr = dt.NewRow();
            Dr["Instruction"] = Matchrow;
            Dr["Code"] = MatchCode;
            Dr["LineNo"] = LineNo;
            Dr["WordNo"] = WordNo;
            Dr["DbSI"] = DbSI;
            Dr["NoOfWord"] = SILength;
            dt.Rows.Add(Dr);
            return dt;
        }
        private bool MatchingSIwithDatabasevalues(clsCnCWord clsWord, int intCurrPageNumber, List<clsCnCWord> PossibleWords, DataTable dt)
        {
            string dbSI = string.Empty;
            string dbSICode = string.Empty;
            string CodeWithoutInstructions = string.Empty;
            string Strword = clsWord.strWord.Replace(" ", "");
            Strword = clsWord.strWord.Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "");
            Strword = Strword.ToUpper();
            string[] DbSIArray = new string[100];
            string retundatavalue = "";
            string MergdataCode = "";
            string tildachk = "";
            string datacopy = "";
            bool matchFound = false;
            clsCnCWord clsDbdata = new clsCnCWord();
            List<string> Remarks = new List<string>();
            List<string> Flags = new List<string>();
            string delimiter = "~";
            string RemarksList = string.Empty;
            string FlagsList = string.Empty;
            DataTable Tempdt = dt.Copy();
            string[] ExcludeFromPercentMatching = { "Saia Guaranteed 5", "Saia Guaranteed 12" };
            string[] NegativeInstruction = {"No Inside Delivery", "NO DELIVERY APPOINTMENT","No Liftgate", "No Liftgate Needed", "Call For Appt: NO",
                "Limited Access: None", "NO LIFTGATE REQUIRED", "No need for a lift gate", "NO DELIVERY APPT REQUIRED",
                "NO DEL APPT REQ","DO NOT PERFORM LIFTGATE/INSIDE DELIVERY","NO SORT AND SEG","NO DELIVERY APPOINTMENT REQUIRED","NOT A SHIPPER LOAD & COUNT",
                "Shipper does not authorize inside delivery or liftgate charges","LIFT GATE DELIVERY FEE IS NOT AUTHORIZED","Liftgate Delivery is not included",
                "Inside Delivery is not included","Do not perform inside delivery","NOT REQUIRE A DELIVERY APPT","Delivery Appt (Not Available)",
                "DO NOT CALL FOR DELIVERY APPOINTMENT","LIFT GATE IS NOT AUTHORIZED","inside Deliveries or Hand Unloads are NOT authorized",
                "NO LARGE TRUCKS APPOINTMENT REQUIRED" };

            string[] SameInstrWithDiffCode = { "Liftgate Pickup", "Do not remove from pallet or break shrink wrap", "FRAGILE, HANDLE WITH CARE" };
            string[] ignoreInstrWithDiffCode = { "Liftgate", "Do not remove from pallet", "HANDLE WITH CARE" };
            bool NegatveSIexist = false;

            //For same instruction with differnt code
            bool SameSIwithDiffCodeexist = false;
            DataTable SortedDt = dt.Copy();
            SortedDt.Columns.Add("WordCount", typeof(int));

            foreach (DataRow r in SortedDt.Rows)
            {
                r["WordCount"] = r["Description"].ToString().Split(new char[] { ' ' }).Length;
            }

            SortedDt.Columns.Add("LenOfCol", typeof(int), "len(Description)");
            SortedDt.AcceptChanges();
            SortedDt.DefaultView.Sort = "WordCount DESC, LenOfCol DESC";

            DataView view = new DataView(SortedDt);
            view.Sort = "WordCount DESC, LenOfCol DESC";
            Tempdt = view.ToTable();


            //For same instruction with differnt code


            try
            {
                if (!string.IsNullOrEmpty(Strword))
                {
                    //foreach (DataRow dr in Tempdt.Rows)
                    foreach (DataRow dr in Tempdt.Rows)
                    {
                        if (Tempdt.Columns.Contains("Description"))
                        {
                            dbSI = dr["Description"].ToString().Replace("  ", " ").Trim();
                        }
                        else
                        {
                            dbSI = dr["Instruction"].ToString().Replace("  ", " ").Trim();
                        }

                        CodeWithoutInstructions = dr["CodeWithoutInstructions"].ToString().Replace("  ", " ").Trim().ToUpper();
                        // string[] splitDbSI = dbSI.Split(' ');

                        datacopy = dbSI.ToUpper().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "").Trim();
                        dbSICode = dr["Code"].ToString().Trim();
                        NegatveSIexist = CheckforNegativeSI(clsWord, NegativeInstruction, datacopy);
                        if (NegatveSIexist == false)
                        {
                            if (Strword.Contains(datacopy))
                            {
                                //For adding distinct values for common instructions
                                bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => datacopy == row.Field<String>("Instruction"));
                                bool containsCode = Module1.ComnInstrConatainer.AsEnumerable().Any(row => dbSICode == row.Field<String>("Code"));
                                if (containsCode == false)
                                {
                                    if (!contains)
                                    {
                                        SameSIwithDiffCodeexist = func_SameInstructDiffCode(Module1.ComnInstrConatainer, SameInstrWithDiffCode, ignoreInstrWithDiffCode, datacopy);
                                        if (SameSIwithDiffCodeexist == false)
                                        {
                                            Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
                                            //Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, splitDbSI.Length);
                                            if (CodeWithoutInstructions == "X")
                                            {
                                                MergdataCode = "#" + dbSICode;
                                            }
                                            else
                                            {
                                                MergdataCode = dbSI.Trim() + "#" + dbSICode;
                                            }

                                            //MergdataCode = dbSI.Trim() + "#" + dbSICode;
                                            retundatavalue = retundatavalue + "~" + MergdataCode;
                                            tildachk = retundatavalue.Substring(0, 1);
                                            if (tildachk == "~")
                                            {
                                                retundatavalue = retundatavalue.Remove(0, 1);
                                            }
                                            Remarks.Add("2F");
                                            Flags.Add("1");
                                        }

                                    }
                                }

                            }
                            else if (!datacopy.Any(char.IsDigit))
                            {
                                if (datacopy.Length > 10)
                                {
                                    RetStructIQSBR050 Obj50 = Module1.ObjSbr.IQSBR050(datacopy, Strword);
                                    //if (Obj50.PercentageMatch > 80)
                                    //if (Obj50.PercentageMatch > 80 && !datacopy.Contains("SAIAGUARANTEED5") && !datacopy.Contains("SAIAGUARANTEED12"))
                                    //if (Obj50.PercentageMatch > 80 && !ExcludeFromPercentMatching.ToString().ToUpper().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "").Trim().Any(datacopy.Contains))
                                    if (Obj50.PercentageMatch > 87)
                                    {
                                        //For adding distinct values for common instructions
                                        bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => datacopy == row.Field<String>("Instruction"));
                                        bool containsCode = Module1.ComnInstrConatainer.AsEnumerable().Any(row => dbSICode == row.Field<String>("Code"));
                                        if (containsCode == false)
                                        {
                                            if (!contains)
                                            {
                                                SameSIwithDiffCodeexist = func_SameInstructDiffCode(Module1.ComnInstrConatainer, SameInstrWithDiffCode, ignoreInstrWithDiffCode, datacopy);
                                                if (SameSIwithDiffCodeexist == false)
                                                {
                                                    Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
                                                    //Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, splitDbSI.Length);

                                                    if (CodeWithoutInstructions == "X")
                                                    {
                                                        MergdataCode = "#" + dbSICode;
                                                    }
                                                    else
                                                    {
                                                        MergdataCode = dbSI.Trim() + "#" + dbSICode;
                                                    }

                                                    //MergdataCode = dbSI.Trim() + "#" + dbSICode;
                                                    retundatavalue = retundatavalue + "~" + MergdataCode;
                                                    tildachk = retundatavalue.Substring(0, 1);
                                                    if (tildachk == "~")
                                                    {
                                                        retundatavalue = retundatavalue.Remove(0, 1);
                                                    }
                                                    Remarks.Add("2F");
                                                    Flags.Add("1");
                                                }

                                                ////Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode);
                                                //////Module1.ComnInstrConatainer = Addrowintodttable(Module1.ComnInstrConatainer, datacopy, dbSICode, LineNo, WordNo, splitDbSI.Length);

                                                ////MergdataCode = dbSI.Trim() + "#" + dbSICode;
                                                ////retundatavalue = retundatavalue + "~" + MergdataCode;
                                                ////tildachk = retundatavalue.Substring(0, 1);
                                                ////if (tildachk == "~")
                                                ////{
                                                ////    retundatavalue = retundatavalue.Remove(0, 1);
                                                ////}
                                                ////Remarks.Add("2F");
                                                ////Flags.Add("1");
                                            }
                                        }

                                    }
                                }

                            }
                        }

                    }
                }

                if (!string.IsNullOrEmpty(retundatavalue))
                {
                    matchFound = true;
                    clsDbdata = ModifiedStringToCNC(clsWord, retundatavalue.ToUpper());
                    PossibleWords.Add(clsDbdata);
                    if (Remarks.Count > 1)
                    {
                        RemarksList = string.Join(delimiter, Remarks.ToArray());
                        FlagsList = string.Join(delimiter, Flags.ToArray());
                    }
                    else
                    {
                        RemarksList = Remarks[0];
                        FlagsList = Flags[0];
                    }

                    // PossibleWords[0].Flag = "1";
                    //if (PossibleWords.Count==1)
                    //{
                    //    PossibleWords[0].Flag = FlagsList;
                    //    PossibleWords[0].Remarks = RemarksList;
                    //}
                    if (PossibleWords.Count >= 1)
                    {
                        PossibleWords[PossibleWords.Count - 1].Flag = FlagsList;
                        PossibleWords[PossibleWords.Count - 1].Remarks = RemarksList;
                    }

                }
            }
            catch (Exception ex)
            { }
            finally
            {
                Tempdt.Dispose();
                Remarks = null;
                Flags = null;
            }
            return matchFound;
        }
        private void RemoveSameSIwithDifferentCode(DataTable dt)
        {
            for (int i = 0; i <= dt.Rows.Count; i++)
            {

            }
        }
        private bool CheckforNegativeSI(clsCnCWord clsWord, string[] NegativeSIarry, string DBdata)
        {
            bool NegativeSIexist = false;
            string Nstring = string.Empty;
            for (int indx = 0; indx < NegativeSIarry.Length; indx++)
            {

                //string strword= clsWord.strWord.ToUpper().Trim().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "");
                //Nstring = NegativeSIarry[indx].ToUpper().Trim().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "") ;
                string strword = Regex.Replace(clsWord.strWord, @"[^0-9a-zA-Z]+", "").ToUpper().Trim();
                Nstring = Regex.Replace(NegativeSIarry[indx], @"[^0-9a-zA-Z]+", "").ToUpper().Trim();
                string DBdataAfterRplace= Regex.Replace(DBdata, @"[^0-9a-zA-Z]+", "").ToUpper().Trim();
                if (strword.Contains(Nstring))
                {

                    ////if (Nstring.Contains(DBdata))
                    if (Nstring.Contains(DBdataAfterRplace))
                    {
                        NegativeSIexist = true;
                    }
                }
            }

            return NegativeSIexist;
        }
        private bool func_SameInstructDiffCode(DataTable dt, string[] arraycheck, string[] ignorearray, string DBdata)
        {
            bool istructPresent = false;
            string ignor_strword = string.Empty;
            string chk_strword = string.Empty;
            for (int i = 0; i < ignorearray.Length; i++)
            {
                ignor_strword = ignorearray[i].ToUpper().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "").Trim();
                if (ignor_strword == DBdata)
                {
                    chk_strword = arraycheck[i].ToUpper().Replace("  ", "").Replace(" ", "").Replace("-", "").Replace("(", "").Replace(")", "").Replace(",", "").Replace("'", "").Trim();
                    bool contains = Module1.ComnInstrConatainer.AsEnumerable().Any(row => chk_strword == row.Field<String>("Instruction"));
                    if (contains == true)
                    {
                        istructPresent = true;
                    }
                }

            }

            return istructPresent;
        }
        #endregion

        #region Supporting Quote No
        private clsCnCWord vilidateQuoteNo(clsCnCWord clsWord)
        {
            clsCnCWord NewCNC = new clsCnCWord();
            string[] splitCNC;
            // clsCnCWord orig_Cnc = clsWord;
            try
            {
                if (clsWord.strWord.Contains("-"))
                {
                    splitCNC = clsWord.strWord.Split('-');
                    //clsWord.strWord = splitCNC[1];
                    for (int indx = 0; indx < splitCNC.Length; indx++)
                    {
                        if ((Module1.reg1.IsMatch(splitCNC[indx].Trim()) || Module1.reg2.IsMatch(splitCNC[indx].Trim()))
                             && !Char.IsLetter(splitCNC[indx][splitCNC[indx].Length - 1]))
                        {
                            //NewCNC = ModifiedStringToCNC(orig_Cnc, splitCNC[indx]);
                            NewCNC = ModifiedStringToCNC(clsWord, splitCNC[indx]);
                        }
                    }
                    //if ((Module1.reg1.IsMatch(clsWord.strWord) || Module1.reg2.IsMatch(clsWord.strWord)))
                    //{
                    //    NewCNC = ModifiedStringToCNC(orig_Cnc, clsWord.strWord);
                    //}
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                // orig_Cnc = null;
                splitCNC = null;
            }



            return NewCNC;
        }
        #endregion

        #region Common supporting Funcn
        private clsCnCWord GetJoinKeywordWord(ClsCNC objCNC, clsCncMetaData ObjMetaData, int intCurrPageNo, int LineNo, int wordno, int Splitlength)
        {
            clsCnCWord CombinedKeyword = new clsCnCWord();
            clsCnCWord[] KeywordArray = new clsCnCWord[10];
            try
            {
                for (int t = 0; t <= Splitlength - 1; t++)
                {
                    KeywordArray[t] = ObjMetaData.Page[intCurrPageNo].Line[LineNo].Word[wordno + t];
                }
                CombinedKeyword = objCNC.MergeWords(KeywordArray);
            }
            catch (Exception ex)
            {
                //MsgBox(ex.Message, MsgBoxStyle.Critical, "F3:GetNextJoinWord");

            }
            finally
            {
                KeywordArray = null;
            }

            return CombinedKeyword;
        }
        public clsCnCWord StringToCNC(int pageNo, string str)
        {
            clsCnCWord objWord = new clsCnCWord();

            objWord.X1Char = "5";
            objWord.Y1Char = "5";
            objWord.X2Char = "10";
            objWord.Y2Char = "10";
            objWord.Confidence = 90;
            objWord.LineNo = 1;
            objWord.PageNo = pageNo;
            objWord.Left = 10;
            objWord.Right = 20;
            objWord.Top = 10;
            objWord.Bottom = 20;
            objWord.strWord = str.ToString();
            objWord.ConfString = "9".PadLeft(str.Length, '9');

            return objWord;

        }
        public clsCnCWord ModifiedStringToCNC(clsCnCWord cncword, string str)
        {
            clsCnCWord objWord = new clsCnCWord();

            objWord.X1Char = cncword.X1Char;
            objWord.Y1Char = cncword.Y1Char;
            objWord.X2Char = cncword.X2Char;
            objWord.Y2Char = cncword.Y2Char;
            objWord.Confidence = cncword.Confidence;
            objWord.LineNo = cncword.LineNo;
            objWord.PageNo = cncword.PageNo;
            objWord.Left = cncword.Left;
            objWord.Right = cncword.Right;
            objWord.Top = cncword.Top;
            objWord.Bottom = cncword.Bottom;
            objWord.strWord = str.ToString();
            objWord.ConfString = cncword.ConfString + "9";

            return objWord;

        }
        private RetStructF3 reurnPossibleWordData(clsCncMetaData ObjMetaData, RetStructF3 returnZones, List<clsCnCWord> PossibleWords, List<int> conflevel, int iRecursiveCallCount, int intCurrPageNumber, clsCnCWord cWord, int Id)
        {

            if ((Id == 6 && Module1.pgSearchInstrut == true && PossibleWords.Count > 1) || (Id == 4 && PossibleWords.Count > 1))
            {
                clsCnCWord combWord = MultipleSI(PossibleWords);
                PossibleWords.Clear();
                PossibleWords.Add(combWord);
            }
            if (PossibleWords.Count > 3)
            {
                PossibleWords = GetTopThreeValue(PossibleWords);
                CompairConfidence(PossibleWords);
            }
            if (PossibleWords.Count > 1)
            {
                CompairConfidence(PossibleWords);
            }
            if (PossibleWords.Count == 1)
            {
                Boolean ConfidenceFlag = HighestConfidence(PossibleWords[0]);

                if (Id == 4 && PossibleWords[0].strWord.Length >= 7 && PossibleWords[0].strWord.Length <= 12)
                {
                    if (Module1.MarkRedSpecialCharacters.Any( PossibleWords[0].strWord.Contains))
                    {
                        PossibleWords[0].Flag = "0";
                        PossibleWords[0].Remarks = "NA";
                    }
                    if (!char.IsDigit(PossibleWords[0].strWord[PossibleWords[0].strWord.Length - 1]))
                    {
                        PossibleWords[0].Flag = "0";
                        PossibleWords[0].Remarks = "NA";
                    }
                }
                if ((Id == 2) || (Id == 1))
                {
                    if (Id == 1) // Pieces
                    {
                        if (PossibleWords[0].strWord.Length >1)
                        {
                            PossibleWords[0].Flag = "0";
                            PossibleWords[0].Remarks = "NA";
                        }
                    }
                    if (Id == 2) //Total weight Explicitly Marked RED 03/08/2021
                    {
                        PossibleWords[0].Flag = "0";
                        PossibleWords[0].Remarks = "NA";
                    }
                    if (PossibleWords[0].Flag == null)
                    {
                       
                        if (ObjMetaData.PageCount == 1 && !Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains) && ConfidenceFlag==true)
                        {
                            PossibleWords[0].Flag = "1";
                            PossibleWords[0].Remarks = "2F";
                        }
                        else
                        {
                            PossibleWords[0].Flag = "0";
                            PossibleWords[0].Remarks = "NA";
                        }
                        // //new added
                        //else if (ObjMetaData.PageCount > 1 ||Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                        // {
                        //     PossibleWords[0].Flag = "0";
                        //     PossibleWords[0].Remarks = "NA";
                        // }
                        // // new added

                    }
                }
                if (Id==3) //COD Explicitly Marked RED
                {
                    PossibleWords[0].Flag = "0";
                    PossibleWords[0].Remarks = "NA";
                }
                //if ((Id == 5)) // without future date
                if ((Id == 5) && Module1.isFuturedate == false)
                {
                    if (PossibleWords[0].Flag == null)
                    {
                        PossibleWords[0].Flag = "0";
                        PossibleWords[0].Remarks = "NA";
                        PossibleWords[0].Confidence = 80;
                        Module1.FuturePDTRADF = false;
                    }
                    else
                    {
                        Module1.FuturePDTRADF = true;
                    }
                }
                else if ((Id == 5) && Module1.isFuturedate == true)
                {
                    Module1.FuturePDTRADF = true;
                }

                if (ConfidenceFlag == true)// if (PossibleWords[0].Confidence >= 90)
                {
                    ////PossibleWords[0].Confidence = 100;
                    //New added
                    if ((Id != 5 ) && Module1.MarkRedSpecialCharacters.Any(PossibleWords[0].strWord.Contains))
                    {
                        PossibleWords[0].Flag = "0";
                        PossibleWords[0].Remarks = "NA";
                    }
                    //New added

                    else if (PossibleWords[0].Flag == null)
                    {
                        PossibleWords[0].Flag = "1";
                        PossibleWords[0].Remarks = "1F";
                    }
                }
                else
                {
                    if (PossibleWords[0].Flag == null)
                    {
                        PossibleWords[0].Flag = "0";
                        PossibleWords[0].Remarks = "NA";
                    }
                }
                returnZones.Words = PossibleWords;
                returnZones.NoOfieldsSuspects = PossibleWords.Count;
                conflevel.Add(90);
                returnZones.ConfidenceLevelofSuspect = conflevel;
                returnZones.Flag = "Y";
                if (iRecursiveCallCount == 1)
                {
                    returnZones.ManualConfirmation = "N";
                    PossibleWords[0].Flag = "0";
                    PossibleWords[0].Remarks = "NA";
                }
                else
                {
                    returnZones.ManualConfirmation = "N";
                }

                // returnZones.ManualConfirmation = "N"
                returnZones.Status = "S";
                //return returnZones;
            }
            if (PossibleWords.Count == 0)
            {
                returnZones = StarReturn(returnZones, PossibleWords, conflevel, intCurrPageNumber, cWord);
            }
            return returnZones;
        }
        private RetStructF3 StarReturn(RetStructF3 returnZones, List<clsCnCWord> PossibleWords, List<int> conflevel, int intCurrPageNumber, clsCnCWord cWord)
        {
            cWord = StringToCNC(intCurrPageNumber, "*****");
            PossibleWords.Add(cWord);
            returnZones.Words = PossibleWords;
            conflevel.Add(90);
            returnZones.Flag = "Y";
            returnZones.ManualConfirmation = "N";
            returnZones.ConfidenceLevelofSuspect = conflevel;
            returnZones.Status = "S";
            returnZones.NoOfieldsSuspects = PossibleWords.Count;
            return returnZones;
        }
        private void CompairConfidence(List<clsCnCWord> PossibleWords)
        {
            double temp_conf = 0;
            double highestconf = 0;
            int ActualIndex = 0;
            clsCnCWord cWord = new clsCnCWord();
            for (int index = 0; index <= PossibleWords.Count - 1; index++)
            {
                temp_conf = PossibleWords[index].Confidence;
                if (temp_conf > highestconf)
                {
                    highestconf = temp_conf;
                    ActualIndex = index;
                    cWord = PossibleWords[index];
                    cWord.Flag = PossibleWords[index].Flag;
                    cWord.Remarks = PossibleWords[index].Remarks;
                }

            }
            PossibleWords.Clear();
            PossibleWords.Add(cWord);

        }
        private Boolean HighestConfidence(clsCnCWord CurrentWord)
        {

            // check the character confidence 
            var NineDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
            var EightDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
            var ZeroDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
            if (CurrentWord.Confidence >= 85 && (CurrentWord.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
            {
                return true;
            }
            else
            {
                return false;
                //ClsOCRValues.ShipperNameOCRWithHighConfidence = "N";
            }
        }
        private Boolean HighConfidenceforMergedWords(clsCnCWord Currentword, string str)
        {
            Boolean flagHighestConfidence = true;
            string confstring = Currentword.ConfString;
            int index = Currentword.strWord.IndexOf(str);
            if (index == -1)  //If not found index then takes confidence of whole value with keyword..this will work only for single merged keywords with value, not for multivalues.
            {
                flagHighestConfidence = HighestConfidence(Currentword);
            }
            else
            {
                for (int i = index; i < Currentword.strWord.Length; i++)
                {
                    if (confstring[i] < '8')
                    {
                        flagHighestConfidence = false;
                        break;
                    }
                    //else
                    //{ return false; }
                }
            }
            
            return flagHighestConfidence;
        }
        private Boolean HighestConfidence_Quote(clsCnCWord CurrentWord, Boolean flag)
        {
            Boolean confidancecheck = false;
            if (flag == true) // for quote length >7
            {
                // check the character confidence 
                var NineDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                
                if (CurrentWord.Confidence >= 90 && (CurrentWord.ConfString.Length == NineDigitcount ))
                {
                    confidancecheck= true;
                }
                else
                {
                    confidancecheck= false;
                }
            }
            else if (flag == false) // for original Rule
            {
                // check the character confidence 
                var NineDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
                var EightDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
                var ZeroDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
                var SevenDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '7').ToArray().Length;
                if (CurrentWord.Confidence >= 80 && (CurrentWord.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount + SevenDigitcount))
                {
                    confidancecheck= true;
                }
                else
                {
                    confidancecheck= false;
                }
            }
            return confidancecheck;
        }
        private DataRow[] RemoveDuplicate(DataRow[] foundRows, DataTable dt, string sKeyword)
        {
            int Line = 0, tempLine = 0, NoofWord = 0, tempNoofWord = 0, counter = 0;

            if (foundRows.Length > 1)
            {
                foreach (DataRow drDataRow in foundRows)
                {
                    Line = Convert.ToInt32(drDataRow.ItemArray[2]);
                    NoofWord = Convert.ToInt32(drDataRow.ItemArray[8]);
                    if (tempLine == Line)
                    {
                        counter++;
                        if (counter >= 1)
                        {
                            if (tempNoofWord != NoofWord)
                            {
                                drDataRow.Delete();
                            }

                        }
                        //Word = Convert.ToInt32(drDataRow.ItemArray[3]);
                        //if (tempWord - Word == 1 || tempWord - Word == -1 || tempWord - Word == 0)
                        //{
                        //    drDataRow.Delete();
                        //}
                    }
                    else
                    {
                        counter = 0;
                    }
                    tempLine = Line;
                    tempNoofWord = NoofWord;
                }
                dt.AcceptChanges();
                foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
            }
            return foundRows;
        }
        private bool CheckMultiple_dateMatch(int pageNo, string str)
        {
            //str = str.Replace("~"," ");
            //Regex reg = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])?(\/|-|\s*)((January|February|March|April|May|June|July|August|September|October|November|December|Jan|Feb|Mar|Apr|May|Jun|Jul|Aug|Sep|Oct|Nov|Dec)|(1[012]|0?[1-9]))(\/|,||-|\s*)(\s*)?(\d{4})");//(?:/[A-Za-z])?");
            Regex reg = new Regex(@"(0[1-9]|1[012])[- /.](0[1-9]|[12][0-9]|3[01])[- /.](19|20)[0-9]{2}");
            int m = reg.Matches(str).Count;
            int counter = 0;
            bool isdaterange = false;
            //clsCnCWord date= new clsCnCWord();
            if (m == 2)
            {

                foreach (Match match in reg.Matches(str))
                {
                    counter++;
                    if (counter == 2)
                    {
                        if (match.Success)
                        {
                            Module1.MABDDate = StringToCNC(pageNo, match.Value.ToString());
                            isdaterange = true;
                        }
                    }

                }
            }
            return isdaterange;
        }
        public clsCnCWord strLineToCNCword(clsCnCLine cncLine, string str)
        {
            clsCnCWord objWord = new clsCnCWord();

            objWord.X1Char = cncLine.LineX1Char;
            objWord.Y1Char = cncLine.LineY1Char;
            objWord.X2Char = cncLine.LineX2Char;
            objWord.Y2Char = cncLine.LineY2Char;
            objWord.Confidence = cncLine.Confidence;
            objWord.LineNo = cncLine.LineNo;
            objWord.PageNo = cncLine.PageNo;
            objWord.Left = cncLine.LineLeft;
            objWord.Right = cncLine.lineRight;
            objWord.Top = cncLine.LineTop;
            objWord.Bottom = cncLine.lineBottom;
            objWord.strWord = str.ToString();
            objWord.ConfString = cncLine.LineConfString + "9";

            return objWord;

        }
        public bool MABDKeywordMatch(string Keyword, string[] MABDKywArry)
        {
            bool isMABDKeywordFound = false;
            for (int iarr = 0; iarr < MABDKywArry.Length; iarr++)
            {
                if (Keyword.ToUpper() == MABDKywArry[iarr].ToUpper())
                {
                    isMABDKeywordFound = true;
                }
            }
            return isMABDKeywordFound;
        }
        #endregion

        private DataTable DataMergeWithKeyword(DataRow[] foundRows, string sKeyword, clsCncMetaData ObjMetaData, int intCurrPageNumber)
        {
            string keywordFound = "";
            int lineno = 0;
            int wordno = 0;
            // DataTable MergeKwordDt = new DataTable();
            sKeyword = sKeyword.Replace("'", "");
            string[] KeyArry = sKeyword.Split(',');
            //MergeKwordDt = func_MergeKwordDt(MergeKwordDt);
            Module1.Dt_KeywordMerge = func_MergeKwordDt(Module1.Dt_KeywordMerge);
            if (foundRows.Length >= 1)
            {
                foreach (DataRow drDataRow in foundRows)
                {
                    keywordFound = Convert.ToString(drDataRow.ItemArray[0]);
                    //pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                    lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                    wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                    string[] sPlitkeyword = keywordFound.Split(' ');
                    //for (int Kword = 0; Kword <= sPlitkeyword.Length-1; Kword++)
                    // {
                    for (int i = 0; i <= KeyArry.Length - 1; i++)
                    {
                        string setArray = KeyArry[i].Trim();
                        string[] splitArray = setArray.Split(' ');

                        if (sPlitkeyword[0] == splitArray[0])
                        {
                            if (sPlitkeyword[sPlitkeyword.Length - 1].ToUpper().StartsWith(splitArray[splitArray.Length - 1].ToUpper()))
                            {
                                if (sPlitkeyword[sPlitkeyword.Length - 1].Length > splitArray[splitArray.Length - 1].Length && sPlitkeyword[sPlitkeyword.Length - 1].Any(char.IsDigit))
                                {
                                    DataRow dr = Module1.Dt_KeywordMerge.NewRow();
                                    dr["ExactKeyword"] = keywordFound;
                                    dr["Keyword"] = setArray;
                                    dr["Line No"] = lineno;
                                    dr["Word No"] = wordno;
                                    Module1.Dt_KeywordMerge.Rows.Add(dr);

                                }
                            }

                        }
                    }
                    //}
                }

            }
            return Module1.Dt_KeywordMerge;
        }

        private DataTable func_MergeKwordDt(DataTable dtKeyword)
        {
            //var dtKeyword = new DataTable();
            dtKeyword.Columns.Add("ExactKeyword", typeof(string));
            dtKeyword.Columns.Add("Keyword", typeof(string));
            dtKeyword.Columns.Add("Line No", typeof(string));
            dtKeyword.Columns.Add("Word No", typeof(string));
            return dtKeyword;
        }

        private string ReplaceStartEndSpecialchar(string strWord)
        {
            //String WordAlphanumeric = Regex.Replace(strWord, @"[^0-9a-zA-Z]+", "");
            //if ((!string.IsNullOrEmpty(WordAlphanumeric)) && (!Module1.SpecialCharacters.Contains(strWord)))
            if ((!string.IsNullOrEmpty(strWord)) && (!Module1.SpecialCharacters.Contains(strWord)))
            {
                strWord = strWord.Replace("'", "");
            LabelStart: if ((!string.IsNullOrEmpty(strWord)) && (Module1.SpecialCharacters.Contains(strWord[0].ToString())))
                {
                    strWord = strWord.Remove(0, 1);
                    goto LabelStart;
                }
            LabelEnd: if ((!string.IsNullOrEmpty(strWord)) && (Module1.SpecialCharacters.Any(strWord.EndsWith)))
                {
                    strWord = strWord.Remove(strWord.Length - 1, 1);
                    goto LabelEnd;
                }
            }
            return strWord;
        }

        private clsCnCWord ContainsKeyword(clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCWord[] TopKeywords, clsCnCWord Topword, int field, int iRecursiveCallCount, string keyword, int keywordLength)
        {

            clsCnCWord KeywordwithValue = new clsCnCWord();
            clsCnCWord ExtractedValue = new clsCnCWord();
            KeywordwithValue = Topword;
            string word;

            if ((TopKeywords[keywordLength - 1].strWord.Contains(":")))
            {
                string[] splitword = TopKeywords[keywordLength - 1].strWord.Split(':');
                word = splitword[1];
            }
            else if (TopKeywords[keywordLength - 1].strWord.Contains("="))
            {
                string[] splitword = TopKeywords[keywordLength - 1].strWord.Split('=');
                word = splitword[1];
            }
            else
            {
                //string replacedString = string.Empty;
                if (Topword.strWord.ToUpper().StartsWith(keyword.ToUpper()))
                {
                    Regex regex = new Regex(keyword.ToUpper());
                    word = regex.Replace(Topword.strWord.ToUpper(), "", 1);
                    //word = Topword.strWord.ToUpper().Replace(keyword.ToUpper(), "");
                }
                else
                {
                    word = Topword.strWord.ToUpper();
                }
            }

            string WordToChckLength = string.Empty;
            if (word.Length > 2)
            {
                word = ReplaceStartEndSpecialchar(word);

            }
            ExtractedValue = ModifiedStringToCNC(Topword, word);
            ////RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");
            ////if (Obj008.Status == "S")
            ////{
            ////    ////if (word.All(char.IsLetterOrDigit))
            ////    ////{
            ////    ////    Obj008.Status = "F";
            ////    ////}
            ////}
            ////if ((!Module1.SpecialCharacters.Contains(word)))
            ////{
            ////    if (!Module1.Time.IsMatch(word))
            ////    {
            ////        WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
            ////    }
            ////}
            if (field == 5)
            {
                //isValidDate = CheckForValidDate(Module1.ObjBoundingWord.RightWord.strWord);
                //if (isValidDate)
                //{
                //    if (!FoundWord(PossibleWords, Module1.ObjBoundingWord.RightWord))
                //    {
                //        // isMABDKeywordFound = MABDKeywordMatch(keywordFound, MABDArray);
                //        isMABDKeywordFound = Common_KeywordMatch(keywordFound, MABDArray);
                //        isDLV_O_KeywordFound = Common_KeywordMatch(keywordFound, Dlv_O_Keywords);

                //        if ((!isMABDKeywordFound && !isTwodatePresent) || (isMABDKeywordFound && isTwodatePresent) || (!isMABDKeywordFound && isTwodatePresent))
                //        {
                //            PossibleWords.Add(Module1.ObjBoundingWord.RightWord);
                //        }
                //        if (isMABDKeywordFound && !isTwodatePresent)
                //        {
                //            Module1.MABDDate = Module1.ObjBoundingWord.RightWord;
                //        }

                //        if (isDLV_O_KeywordFound && !isTwodatePresent && !isMABDKeywordFound)
                //        {
                //            DataRow Dr = dt_O_Dlv.NewRow();
                //            Dr["keyword"] = keywordFound;
                //            Dr["Line No"] = lineno;
                //            Dr["Word No"] = Module1.ObjBoundingWord.RightWord.WordNumber;
                //            Dr["flag"] = 1;
                //            dt_O_Dlv.Rows.Add(Dr);
                //        }
                //    }
                //}

            }

            ////                }

            ////            }

            ////        }

            ////    }
            return ExtractedValue;
        }

        private Boolean Keyword_value_Heightdiff(clsCnCWord CombinedKeyword, clsCnCWord value)
        {
            int keywordTop = CombinedKeyword.Top;
            int keywordBottom = CombinedKeyword.Bottom;
            int KeywordHeight = keywordBottom - keywordTop;
            int TopValue = value.Top;
            int DistanceDiff = TopValue - keywordBottom;
            if (DistanceDiff <= KeywordHeight * 3)
            {
                return true;
            }
            else
            {
                return false;
            }
        }
    }
}