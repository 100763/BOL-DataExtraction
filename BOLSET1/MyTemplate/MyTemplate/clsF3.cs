using iQDataProvider;
using iQPUBLIC;
using System;
using System.Collections.Generic;
using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data;
using System.Collections;
using System.Linq;
using System.IO;
using System.Text;
//using BOL_SET1;

namespace BOLSET1IQ
{
    public class clsF3
    {
        //BOL_SET1.clsF3 objBOLF3;
        //List<string> ContainsSeparator = new List<string>() {",","-","/" };

        public RetStructF3 F3(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strRoutineNo, string strArg)
        {
            string isEDI = "0";
            if (strArg.Contains('#'))
            {
                isEDI = strArg.Split('#')[1].Trim();
                strArg = strArg.Split('#')[0];
                
            }

            string[] strCordinates = strArg.Split('~');
            Module1.MX1 = Convert.ToInt32(strCordinates[0]);
            Module1.MY1 = Convert.ToInt32(strCordinates[1]);
            Module1.MX2 = Convert.ToInt32(strCordinates[2]);
            Module1.MY2 = Convert.ToInt32(strCordinates[3]);

            


            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;

           Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;

            //objBOLF3 = new BOL_SET1.clsF3();

            if (Module1.imageName.ToLower() != Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath).ToLower())
            {
                Module1.bFoundSearchKeywordTable = false;
               // Module1.func_ReadBarcode();
                Module1.imageName = Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath);
                Module1.ProNo = Path.GetFileNameWithoutExtension(iQPUBLIC.PublicComponents.PrimaryImagePath);
                Module1.ProNo = Module1.ProNo.Split('_')[Module1.ProNo.Split('_').Length - 1];
            }


            if (isEDI == "0")
            {

                switch (strRoutineNo)
                {
                    case "CKM202": //FreightBillNumber

                        //objRetStructF3 = F3_AutoLocate_BillOfLading(ObjMetaData, intCurrPageNumber, strArg);
                        //Module1.objRetStructF3 = FileName(ObjMetaData, intCurrPageNumber, strArg);
                        break;

                    case "CKM213": //BillOfLading

                        //objRetStructF3 = F3_AutoLocate_BillOfLading(ObjMetaData, intCurrPageNumber, strArg);
                        //Module1.objRetStructF3 = BOLNumber(ObjMetaData, intCurrPageNumber, strArg);
                        Module1.objRetStructF3 = BOLNumber(ObjMetaData, intCurrPageNumber, strArg);
                        break;

                    case "CKM213_1": //Load#

                        //Module1.objRetStructF3 = LoadId(ObjMetaData, intCurrPageNumber, strArg);
                        Module1.objRetStructF3 = LoadId(ObjMetaData, intCurrPageNumber, strArg);
                        break;


                    case "CKM214"://PurchaseOrder

                        //Module1.objRetStructF3 = PONumber(ObjMetaData, intCurrPageNumber, strArg);
                        Module1.objRetStructF3 = PONumber(ObjMetaData, intCurrPageNumber, strArg);

                        break;

                    case "CKM215"://ShipperNumber

                        //Module1.objRetStructF3 = ShipperNumber(ObjMetaData, intCurrPageNumber, strArg);
                        Module1.objRetStructF3 = ShipperNumber(ObjMetaData, intCurrPageNumber, strArg);
                        break;

                    case "CKM218"://HazmatTelephone

                        //Module1.objRetStructF3 = Tel_HazMat(ObjMetaData, intCurrPageNumber, strArg);
                        Module1.objRetStructF3 = Tel_HazMat(ObjMetaData, intCurrPageNumber, strArg);

                        break;

                    case "CKM219"://IndividualTelephone

                        // Module1.objRetStructF3 = Tel_Indiv(ObjMetaData, intCurrPageNumber, strArg);

                        break;

                    case "CKM218_1"://ERTCode

                        //Module1.objRetStructF3 = ERTCode(ObjMetaData, intCurrPageNumber, strArg);
                        Module1.objRetStructF3 = ERTCode(ObjMetaData, intCurrPageNumber, strArg);

                        break;

                }
            }
            else
            {
                Module1.PossibleWords.Add(StringToCncWord_EDI(1, "*****"));

                Module1.objRetStructF3.Words = Module1.PossibleWords;
                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                Module1.conflevel.Add(90);
                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                Module1.objRetStructF3.Flag = "Y";
                Module1.objRetStructF3.ManualConfirmation = "N";
                Module1.objRetStructF3.Status = "S";
            }

            //return Module1.objRetStructF3;
            return Module1.objRetStructF3;

        }

        //private RetStructF3 F3_AutoLocate_FreightBillNumber(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        //{

        //    RetStructF3 objRetStructF3 = new RetStructF3();
        //    objRetStructF3.Status = "F";
        //    objRetStructF3.NoOfieldsSuspects = 0;


        //    try
        //    {

        //        objRetStructF3 = FileName(ObjMetaData, intCurrPageNumber, strArg);

        //    }


        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("F3: F3_AutoLocate_BillOfLading " + ex.Message);
        //    }


        //    return objRetStructF3;

        //}

        //private RetStructF3 F3_AutoLocate_BillOfLading(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        //{

        //    RetStructF3 objRetStructF3 = new RetStructF3();
        //    objRetStructF3.Status = "F";
        //    objRetStructF3.NoOfieldsSuspects = 0;


        //    try
        //    {

        //        objRetStructF3 = BOLNumber(ObjMetaData, intCurrPageNumber, strArg);

        //    }


        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("F3: F3_AutoLocate_BillOfLading " + ex.Message);
        //    }


        //    return objRetStructF3;

        //}

        //private RetStructF3 F3_AutoLocate_PurchaseOrder(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        //{

        //    RetStructF3 objRetStructF3 = new RetStructF3();
        //    objRetStructF3.Status = "F";
        //    objRetStructF3.NoOfieldsSuspects = 0;


        //    try
        //    {

        //        objRetStructF3 = PONumber(ObjMetaData, intCurrPageNumber, strArg);

        //    }


        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("F3: F3_AutoLocate_PurchaseOrder " + ex.Message);
        //    }


        //    return objRetStructF3;

        //}

        //private RetStructF3 F3_AutoLocate_ShipperNumber(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        //{

        //    RetStructF3 objRetStructF3 = new RetStructF3();
        //    objRetStructF3.Status = "F";
        //    objRetStructF3.NoOfieldsSuspects = 0;

        //    try
        //    {

        //        objRetStructF3 = ShipperNumber(ObjMetaData, intCurrPageNumber, strArg);

        //    }


        //    catch (Exception ex)
        //    {
        //        MessageBox.Show("F3: F3_AutoLocate_ShipperNumber " + ex.Message);
        //    }


        //    return objRetStructF3;

        //}

        //#region Freight Bill Number (File name)
        //private RetStructF3 FileName(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        //{
        //    //RetStructF3 returnZones = new RetStructF3();
        //    List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
        //    Module1.objRetStructF3.Status = "F";
        //    Module1.objRetStructF3.ManualConfirmation = "Y";
        //    Module1.objRetStructF3.Flag = "Y";
        //    //ObjF3.Status = "F";
        //    //ObjF3.NoOfieldsSuspects = 0;
        //    //ClsCNC objCNC = new ClsCNC();
        //    //clsCNCSBR ObjSbr = new clsCNCSBR();

        //    List<int> conflevel = new List<int>();

        //    try
        //    {
        //        string FileName = string.Empty;

        //        //string sPath = Path.GetFileNameWithoutExtension(iQPUBLIC.PublicComponents.InputBatchPath.ToString().Trim());
        //        //FileName = Path.GetFileName(iQPUBLIC.PublicComponents.PrimaryImagePath);
        //        FileName = Path.GetFileNameWithoutExtension(iQPUBLIC.PublicComponents.PrimaryImagePath);
        //        //FileName = FileName.Replace(".TIF", "");
        //        //FileName = FileName.Replace(".tif", "");

        //        clsCnCWord ObjWord;

        //        ObjWord = StringToCncWord(intCurrPageNumber, FileName);
        //        PossibleWords.Add(ObjWord);

        //        if (PossibleWords.Count > 0)
        //        {
        //            Module1.objRetStructF3.Words = PossibleWords;
        //            Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
        //            conflevel.Add(90);
        //            Module1.objRetStructF3.ConfidenceLevelofSuspect = conflevel;
        //            Module1.objRetStructF3.Flag = "Y";
        //            Module1.objRetStructF3.ManualConfirmation = "N";
        //            // returnZones.ManualConfirmation = "N"
        //            Module1.objRetStructF3.Status = "S";
        //            //return returnZones;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //       // MessageBox.Show(ex.Message);
        //    }
        //    return Module1.objRetStructF3;
        //}
        private clsCnCWord StringToCncWord(int pageNo, string StringToConvert)
        {
            clsCnCWord ObjWord = new clsCnCWord();
            ObjWord.X1Char = "5";
            ObjWord.Y1Char = "5";
            ObjWord.X2Char = "10";
            ObjWord.Y2Char = "10";
            ObjWord.Confidence = 90;
            ObjWord.PageNo = pageNo;
            ObjWord.Left = 10;
            ObjWord.Right = 20;
            ObjWord.Top = 10;
            ObjWord.Bottom = 20;
            ObjWord.strWord = StringToConvert.ToString();
            ObjWord.ConfString = "9".PadLeft(StringToConvert.Length, '9');

            return ObjWord;
        }

        private clsCnCWord StringToCncWord_EDI(int pageNo, string StringToConvert)
        {
            clsCnCWord ObjWord = new clsCnCWord();
            ObjWord.X1Char = "5";
            ObjWord.Y1Char = "5";
            ObjWord.X2Char = "10";
            ObjWord.Y2Char = "10";
            ObjWord.Confidence = 90;
            ObjWord.PageNo = pageNo;
            ObjWord.Left = 10;
            ObjWord.Right = 20;
            ObjWord.Top = 10;
            ObjWord.Bottom = 20;
            ObjWord.Flag = "0";
            ObjWord.Remarks = "EDI";
            ObjWord.strWord = StringToConvert.ToString();
            ObjWord.ConfString = "9".PadLeft(StringToConvert.Length, '9');

            return ObjWord;
        }
        //#endregion

        #region Bill of Lading Number       
        private RetStructF3 BOLNumber(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        {
            DataSet ds = Module1.MakeDs(oMeta, intCurrPageNumber);

            //RetStructF3 objRetStructF3 = new RetStructF3();

            DataTable dt = new DataTable();
            //List<clsCnCWord> PossibleWords1 = new List<clsCnCWord>();
            //List<int> conflevel = new List<int>();
            Module1.PossibleWords.Clear();
            Module1.conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword;
            int flag = 0;

            System.Data.DataRow[] foundRows;
            clsCnCWord[] TopKeywords = new clsCnCWord[50];
            clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();

            string[] FilterKeywords = { "CONTAINS","CONTAIN","NOTE","NOTES","INFORMATION" };

            try
            {
                //Line1:
                //    if (iRecursiveCallCount == 0)
                //    {
                //        // -------------------------------------------------BILL OF LADING NUMBER KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------               
                //        sKeyword = "'RGA#','RA#','RA','RMA','RGA','RETURN AUTHORIZATION','Bill Of Lading #','Bill of Lading#','BHI OF LADING #','Bill of Lading Number','Bill of Lading No','Bill of L ading No','Bil of Lading Number:','BILLOF LADING NUMBER','SHIPPER BILL OF LADING NUMBER:','Bill of LadingNumber','BILL OF LADINGNUMBER','BILL OF LADING REF NO','BILL OFLADING NO','BDL OF LADING NO''B/L Document No.','B/L NUMBER','BILL OF LADING NUMBER/SHIPMENT NUMBER','ORDER/BILL OF LADING NO','Shippers Bill of Ladias No','BIE OF LADING NO','BE OF LADING NO','BIL OF LADING NO','Shippers bill of lading','BOL No','BOLNUMBER','B/L NO.','BOL#','BOL NBR','BILL OF LADING - ME','BOL','B/L #','BDL NO','BOL ID','BOLNBR','BOL NO (LD#)','B L #:','BL#','BOL/SHIPMENT NO','Bill Number','B/L#','B0L#','B L#:','BOLNO','BOL/Order#:','DO/BOL','B0 L #','BILL OF LADING / BOOKING NUMBER','SID/BOL#:','BOL/SHIPMENT#','BILL OF LADING / PACKING SLIP','HBL','MBL'";
                //        //sKeyword = "'RA','RMA','RGA','RETURN AUTHORIZATION','RGA#','RA#','Bill of Lading Number','Bill of Lading No','Shippers bill of lading','Bill of L ading No','Shippers Bill of Ladias No','BOL Number','Bill of LadingNumber','SHIPPER BILL OF LADING NUMBER:','B/L Document No.','B/L NUMBER','BOL No','B/L NO.','BOL NBR','BOL#','BOL', 'B/L', 'Manifest','BOL ID','Bill Number','HBL','MBL','Manifest Id','Bill Of Lading #','BOL/SHIPMENT NO','BE OF LADING NO','BIL OF LADING NO','Bil of Lading Number:','Bill Number','Bill of Lading#','B/L#','B0L#','B L#:','BOLNO','BOL/Order#:','DO/BOL'";
                //    }
                //    else
                //    {
                //        sKeyword = "'Bill of Lading','BILLOFLADING'";
                //    }

                //try
                //{
                //Module1.objRetStructF3.Status = "F";
                //Module1.objRetStructF3.ManualConfirmation = "Y";
                //Module1.objRetStructF3.Flag = "Y";

                Module1.objRetStructF3.Status = "F";
                Module1.objRetStructF3.ManualConfirmation = "Y";
                Module1.objRetStructF3.Flag = "Y";

                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSet1_keyword"))
                {
                    Module1.dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSet1_keyword"];
                }
                if (Module1.dsKeys != null)
                {

                    dt = Module1.dsKeys.Tables[0];

                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]="+intCurrPageNumber);

                    //-----For Single Page
                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    //---For multiPage
                    for (int iPage = 1; iPage <= oMeta.PageCount; iPage++)
                    //for (int iPage = intCurrPageNumber; iPage <= intCurrPageNumber; iPage++)
                    {
                        iRecursiveCallCount = 0;
                    Line1:
                        if (iRecursiveCallCount == 0)
                        {
                            // -------------------------------------------------BILL OF LADING NUMBER KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------               
                            sKeyword = "'RGA#','RA#','RMA#','RA','RMA','RGA','RETURN AUTHORIZATION','Bill Of Lading #','Bill of Lading#','BHI OF LADING #','Bill of Lading Number','Bill of Lading Num','Bill Of Lading Nos','Bill of Lading No','Bill of L ading No','Bil of Lading Number:','BILLOF LADING NUMBER','SHIPPER BILL OF LADING NUMBER:','Bill of LadingNumber','BILL OF LADINGNUMBER','BM OF LADING NUMBER','BIOF LADING NUMBER','BLOF LADING NUMBER','BILL OF LAD ING NUMBER','BILL OF LADING REF NO','BILL OFLADING NO','BDL OF LADING NO','B/L Document No.','B/L NUMBER','BILL OF LADING NUMBER/SHIPMENT NUMBER','ORDER/BILL OF LADING NO','Shippers Bill of Ladias No','BIE OF LADING NO','BE OF LADING NO','BIL OF LADING NO','Shippers bill of lading','BOL No','BOLNUMBER','B/L NO.','BOL#','BOL NBR','BOL NUM','BILL OF LADING - ME','B/L #','BDL NO','BOL ID','BOLNBR','BOL NUMBER','BOL NO (LD#)','BOL','B L #:','BL#','BOL/SHIPMENT NO','Bill Number','B/L#','B0L#','B L#:','BOLNO','BOL/Order#:','DO/BOL','B0 L #','BILL OF LADING / BOOKING NUMBER','SID/BOL#:','BOL/SHIPMENT#','BILL OF LADING / PACKING SLIP','Underlying BOLs','PRO/BOL Number'";
                            //sKeyword = "'RA','RMA','RGA','RETURN AUTHORIZATION','RGA#','RA#','Bill of Lading Number','Bill of Lading No','Shippers bill of lading','Bill of L ading No','Shippers Bill of Ladias No','BOL Number','Bill of LadingNumber','SHIPPER BILL OF LADING NUMBER:','B/L Document No.','B/L NUMBER','BOL No','B/L NO.','BOL NBR','BOL#','BOL', 'B/L', 'Manifest','BOL ID','Bill Number','HBL','MBL','Manifest Id','Bill Of Lading #','BOL/SHIPMENT NO','BE OF LADING NO','BIL OF LADING NO','Bil of Lading Number:','Bill Number','Bill of Lading#','B/L#','B0L#','B L#:','BOLNO','BOL/Order#:','DO/BOL'";
                        }
                        else
                        {
                            sKeyword = "'Bill of Lading','BILLOFLADING'";
                        }

                        //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND  [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                        //var orderedRows = foundRows.OrderBy(item => item.ItemArray[2]);//oreder by only Line No
                        foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + iPage + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                        if (foundRows.Length > 1)
                        {
                            DataRow[] SortedRows = sortFoundRowsonKeywordPreference(sKeyword, foundRows, dt);
                            if (SortedRows.Length == foundRows.Length)
                            {
                                foundRows = SortedRows;
                            }
                        }

                        var orderedRows = foundRows.OrderBy(item => item.ItemArray[1]).ThenBy(item => item.ItemArray[2]);//oreder by only Page No then by Line no
                        foundRows = orderedRows.ToArray();

                        

                        //if(foundRows.Length>1)
                        //{
                        //    foundRows = sortFoundRowsonKeywordPreference(sKeyword, foundRows, dt);
                        //}

                        // var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[8]);
                        //var orderedRows = foundRows.OrderBy(item => item.ItemArray[2]);
                        //foundRows = orderedRows.ToArray();
                        string keyword = string.Empty;
                        int pgno = 0;
                        int lineno = 0;
                        int wordno = 0;
                        int LineWordNo = 0;
                        int x1 = 0;
                        int y1 = 0;
                        int x2 = 0;
                        int y2 = 0;
                        //int h = 0;
                        int w = 0;

                        int Count = 0;       //To check how many RA related keywords present

                        //string[] FilterRightWord = { ":", ".", "#", "#:", "NUMBER", "NO.", "NO", "NUMBER:", "NO:","|", "•","-","NUMBER(S)" };
                        string[] FilterRightWord = { ":", ".", "#", "#:", "#.", "#·", "|", "•", "-", "NUMBER(S)", "/", "'", "NUMBER", "NO.", "NO", "NUMBER:", "NO:", "_", "(LD#):", "'-","NUM","NUM:","ID", "ID:" };
                        //ClsCNC oClsCnc = new ClsCNC();
                        //clsCnCWord Words = new clsCnCWord();

                        //clsCnCWord sJoinWord = new clsCnCWord();
                        //regex = new Regex("^([a-zA-Z]*[-/]?[a-zA-Z]*[0-9]+[-/]?[a-zA-Z]*[-/]?)*$");
                        if (foundRows.Length == 0 && iRecursiveCallCount == 0)
                        {
                            iRecursiveCallCount = iRecursiveCallCount + 1;
                            if (iRecursiveCallCount == 1)
                            {
                                goto Line1;
                            }
                        }
                        if (foundRows.Length == 1)
                        {
                            Array.Clear(TopKeywords, 0, TopKeywords.Length);
                            keyword = (foundRows[0].ItemArray[0]).ToString();
                            pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                            lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                            wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                            LineWordNo = Convert.ToInt32(foundRows[0].ItemArray[4]);
                            x1 = Convert.ToInt32(foundRows[0].ItemArray[5]);
                            y1 = Convert.ToInt32(foundRows[0].ItemArray[6]);
                            x2 = Convert.ToInt32(foundRows[0].ItemArray[7]);
                            y2 = Convert.ToInt32(foundRows[0].ItemArray[8]);
                            w = oMeta.Page[pgno].ImageWidth;


                            if ((Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75)) < 1000)
                            {
                                y2 = y2;
                            }
                            else
                            {
                                y2 = Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75);
                            }

                            clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);


                            if (Lines != null)
                            {
                                if (Lines.Length <= lineno)
                                {
                                    iRecursiveCallCount = iRecursiveCallCount + 1;
                                    if (iRecursiveCallCount == 1)
                                    {
                                        goto Line1;
                                    }
                                    //Module1.objRetStructF3.Status = "F";
                                    //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                    //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                    //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                    //Module1.conflevel.Add(90);
                                    //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                    //Module1.objRetStructF3.Status = "S";
                                    ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                    //Module1.objRetStructF3.ManualConfirmation = "N";
                                    //Module1.objRetStructF3.Flag = "Y";
                                    //return Module1.objRetStructF3;
                                    continue;
                                }
                            }
                            else
                            {
                                iRecursiveCallCount = iRecursiveCallCount + 1;
                                if (iRecursiveCallCount == 1)
                                {
                                    goto Line1;
                                }
                                //Module1.objRetStructF3.Status = "F";
                                //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                //Module1.conflevel.Add(90);
                                //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                //Module1.objRetStructF3.Status = "S";
                                ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                //Module1.objRetStructF3.ManualConfirmation = "N";
                                //Module1.objRetStructF3.Flag = "Y";
                                //return Module1.objRetStructF3;
                                continue;
                            }


                            //string[] SplitKeyword = keyword.Split();
                            int i;
                            int KeywordLength = keyword.Split().Length;
                            string[] SplitKeyword = keyword.Split();
                            string lastWordOfKeyword = SplitKeyword[KeywordLength - 1];

                            for (i = 0; i < KeywordLength; i++)
                            {
                                TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                            }

                            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);
                            clsCnCBoundingWords BoundingwordForLeft = Module1.oClsCnc.GetBoundingWords(TopKeywords[0], oMeta, pgno);

                            if (keyword.ToUpper() == "RA")
                            {
                                if (!TopWord.strWord.All(char.IsUpper))
                                {
                                    //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                    //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                    //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                    //Module1.conflevel.Add(90);
                                    //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                    //Module1.objRetStructF3.Status = "S";
                                    ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                    //Module1.objRetStructF3.ManualConfirmation = "N";
                                    //Module1.objRetStructF3.Flag = "Y";
                                    //return Module1.objRetStructF3;
                                    continue;
                                }
                            }

                            //if (BoundingwordForLeft.LeftWord != null && BoundingwordForLeft.LeftWord.strWord.ToUpper() == "MASTER")
                            //{
                            //    //Module1.objRetStructF3.Status = "F";
                            //    //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                            //    //Module1.objRetStructF3.Words = Module1.PossibleWords;
                            //    //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            //    //Module1.conflevel.Add(90);
                            //    //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            //    //Module1.objRetStructF3.Status = "S";
                            //    ////Module1.objRetStructF3.ManualConfirmation = "Y";
                            //    //Module1.objRetStructF3.ManualConfirmation = "N";
                            //    //Module1.objRetStructF3.Flag = "Y";
                            //    //return Module1.objRetStructF3;
                            //    continue;
                            //}

                            //if (TopWord.strWord.Any(char.IsDigit) && TopWord.strWord.Length <= 3)
                            if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                            {
                                ContainsKeyword(oMeta, pgno, TopKeywords, TopWord, 1, iRecursiveCallCount, keyword, KeywordLength);
                            }
                            else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1 && TopKeywords[KeywordLength - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                            {
                                //if (TopKeywords[KeywordLength - 1].strWord.Contains(":"))
                                //{
                                //    string[] splitword = TopKeywords[KeywordLength - 1].strWord.Split(':');
                                //    clsCnCWord oWord = TopKeywords[KeywordLength - 1];
                                //    oWord.strWord = splitword[1];
                                //    ContainsKeyword(oMeta, intCurrPageNumber, TopKeywords, oWord, 1, iRecursiveCallCount, keyword.Split()[KeywordLength - 1], KeywordLength);
                                //}
                                //else
                                //{
                                ContainsKeyword(oMeta, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 1, iRecursiveCallCount, keyword.Split()[KeywordLength - 1], KeywordLength);
                                //}
                            }
                            else
                            {
                                if (keyword == "B/L")
                                {
                                    if (objBoundingWord.RightWord != null)
                                    {
                                        if (!FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                        {
                                            //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                            //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                            //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                            //Module1.conflevel.Add(90);
                                            //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                            //Module1.objRetStructF3.Status = "S";
                                            ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                            //Module1.objRetStructF3.ManualConfirmation = "N";
                                            //Module1.objRetStructF3.Flag = "Y";
                                            //return Module1.objRetStructF3;
                                            continue;
                                        }
                                    }
                                    else
                                    {
                                        //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                        //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                        //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                        //Module1.conflevel.Add(90);
                                        //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                        //Module1.objRetStructF3.Status = "S";
                                        ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                        //Module1.objRetStructF3.ManualConfirmation = "N";
                                        //Module1.objRetStructF3.Flag = "Y";
                                        //return Module1.objRetStructF3;
                                        continue;
                                    }
                                }
                                flag = 0;

                            Label1: if (objBoundingWord.RightWord != null)
                                {
                                    if (!(FilterKeywords.Contains(objBoundingWord.RightWord.strWord.ToUpper())))
                                    {
                                        if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                        {
                                            if (objBoundingWord.RightWord.strWord == "#" || objBoundingWord.RightWord.strWord == "#:" || objBoundingWord.RightWord.strWord == "#.")
                                            {
                                                flag = 1;
                                            }
                                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                            goto Label1;
                                        }
                                        //CheckRightWord(oMeta, objBoundingWord, pgno,Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                        CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 1, iRecursiveCallCount, keyword);
                                    }
                                }
                                //else if (LineWordNo < oMeta.Page[intCurrPageNumber].Line[lineno].WordCount)
                                else
                                {
                                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                    if (objBoundingWord.BottomWords != null)
                                    {
                                        //CheckBottomWord(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                        CheckBottomWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 1, iRecursiveCallCount, keyword);
                                    }
                                    else
                                    {
                                        clsCnCWord BOL = FindBottomfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], oMeta.Page[pgno].Line[lineno].Word[LineWordNo], 4, 1, TopKeywords, iRecursiveCallCount, keyword);
                                        if (BOL.strWord != null)
                                        {
                                            Boolean exist = CheckforExistWord(Module1.PossibleWords, BOL);
                                            if (exist == false)
                                            {
                                                Module1.PossibleWords.Add(BOL);
                                            }
                                        }
                                    }
                                    //if (objBoundingWord.TopWords != null)
                                    //{
                                    //    CheckTopWord(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                    //}
                                    //else
                                    //{
                                    //    BOL = FindTopfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], oMeta.Page[pgno].Line[lineno].Word[LineWordNo]);
                                    //    if (BOL.strWord != null)
                                    //    {
                                    //        Boolean exist = CheckforExistWord(Module1.PossibleWords, BOL);
                                    //        if (exist == false)
                                    //        {
                                    //            Module1.PossibleWords.Add(BOL);
                                    //        }
                                    //    }
                                    //}

                                }
                            }
                            if (Module1.PossibleWords.Count == 0)
                            {
                                if (keyword == "RA")
                                {
                                    iRecursiveCallCount = iRecursiveCallCount + 1;
                                    if (iRecursiveCallCount == 1)
                                    {
                                        goto Line1;
                                    }
                                }
                                //else
                                //{
                                //    Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                //}
                            }

                            //Module1.objRetStructF3 = GetValues(oMeta, pgno, 1, iRecursiveCallCount, TopKeywords);
                            //return Module1.objRetStructF3;


                        }
                        else
                        {
                            foreach (DataRow drDataRow in foundRows)
                            {
                                Array.Clear(TopKeywords, 0, TopKeywords.Length);
                                keyword = (drDataRow.ItemArray[0]).ToString();
                                pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                                lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                                wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                                LineWordNo = Convert.ToInt32(drDataRow.ItemArray[4]);
                                x1 = Convert.ToInt32(drDataRow.ItemArray[5]);
                                y1 = Convert.ToInt32(drDataRow.ItemArray[6]);
                                x2 = Convert.ToInt32(drDataRow.ItemArray[7]);
                                y2 = Convert.ToInt32(drDataRow.ItemArray[8]);

                                w = oMeta.Page[pgno].ImageWidth;

                                if (Module1.RAKeywords.Contains(keyword.ToUpper()))
                                {
                                    Count = Count + 1;
                                }
                                if ((Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75)) < 1000)
                                {
                                    y2 = y2;
                                }
                                else
                                {
                                    y2 = Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75);
                                }

                                clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);
                                if (Lines != null)
                                {
                                    if (Lines.Length <= lineno)
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    continue;
                                }

                                //clsCnCWord[] TopKeywords = new clsCnCWord[50];
                                //string[] SplitKeyword = keyword.Split();
                                int i;

                                int KeywordLength = keyword.Split().Length;
                                string[] SplitKeyword = keyword.Split();
                                string lastWordOfKeyword = SplitKeyword[KeywordLength - 1];

                                for (i = 0; i < KeywordLength; i++)
                                {
                                    TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                                }

                                clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                                objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);
                                clsCnCBoundingWords BoundingwordForLeft = Module1.oClsCnc.GetBoundingWords(TopKeywords[0], oMeta, pgno);

                                if (keyword.ToUpper() == "RA")
                                {
                                    if (!TopWord.strWord.All(char.IsUpper))
                                    {
                                        continue;
                                    }
                                }

                                //if (BoundingwordForLeft.LeftWord != null && BoundingwordForLeft.LeftWord.strWord.ToUpper() == "MASTER")
                                //{
                                //    continue;
                                //}
                                if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                                {
                                    ContainsKeyword(oMeta, pgno, TopKeywords, TopWord, 1, iRecursiveCallCount, keyword, KeywordLength);
                                }
                                else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1 && TopKeywords[KeywordLength - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                                {
                                    ContainsKeyword(oMeta, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 1, iRecursiveCallCount, keyword.Split()[KeywordLength - 1], KeywordLength);
                                }
                                else
                                {
                                    if (keyword == "B/L")
                                    {
                                        if (objBoundingWord.RightWord != null)
                                        {
                                            if (!FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                            {
                                                {
                                                    continue;
                                                }
                                            }
                                        }
                                        else { continue; }
                                    }

                                    flag = 0;
                                //if ((objBoundingWord.RightWord != null) && regex.IsMatch(objBoundingWord.RightWord.strWord.Replace(":", "")) || (objBoundingWord.RightWord != null))
                                Label1: if (objBoundingWord != null && objBoundingWord.RightWord != null)
                                    {
                                        if (!(FilterKeywords.Contains(objBoundingWord.RightWord.strWord.ToUpper())))
                                        {
                                            if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                            {
                                                if (objBoundingWord.RightWord.strWord == "#" || objBoundingWord.RightWord.strWord == "#:" || objBoundingWord.RightWord.strWord == "#.")
                                                {
                                                    flag = 1;
                                                }
                                                objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                                goto Label1;
                                            }
                                            CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 1, iRecursiveCallCount, keyword);
                                        }
                                    }

                                    else
                                    {
                                        //clsCnCWord BOL = new clsCnCWord();
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                        if (objBoundingWord.BottomWords != null)
                                        {
                                            CheckBottomWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 1, iRecursiveCallCount, keyword);
                                        }
                                        else
                                        {
                                            clsCnCWord BOL = FindBottomfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], oMeta.Page[pgno].Line[lineno].Word[LineWordNo], 4, 1, TopKeywords, iRecursiveCallCount, keyword);
                                            if (BOL.strWord != null)
                                            {
                                                Boolean exist = CheckforExistWord(Module1.PossibleWords, BOL);
                                                if (exist == false)
                                                {
                                                    Module1.PossibleWords.Add(BOL);

                                                }
                                            }

                                        }
                                        //if (objBoundingWord.TopWords != null)
                                        //{
                                        //    CheckTopWord(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                        //}
                                        //else
                                        //{
                                        //    BOL = FindTopfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], oMeta.Page[pgno].Line[lineno].Word[LineWordNo]);
                                        //    if (BOL.strWord != null)
                                        //    {
                                        //        Boolean exist = CheckforExistWord(Module1.PossibleWords, BOL);
                                        //        if (exist == false)
                                        //        {
                                        //            Module1.PossibleWords.Add(BOL);
                                        //        }
                                        //    }
                                        //}

                                    }
                                }
                            } // end of for 

                            if (Module1.PossibleWords.Count == 0)
                            {
                                if (Count == foundRows.Length)
                                {
                                    iRecursiveCallCount = iRecursiveCallCount + 1;
                                    if (iRecursiveCallCount == 1)
                                    {
                                        goto Line1;
                                    }
                                }

                                //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                            }

                            //Module1.objRetStructF3 = GetValues(oMeta, pgno, 1, iRecursiveCallCount, TopKeywords);
                            //return Module1.objRetStructF3;

                        }                    // end of the F3
                    }
                    if (Module1.PossibleWords.Count == 0)
                    {
                        Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                    }
                    Module1.objRetStructF3 = GetValues(oMeta, intCurrPageNumber, 1, iRecursiveCallCount, TopKeywords);
                    return Module1.objRetStructF3;
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                //throw;
            }
            finally
            {
                //ds.Clear();                
                ds.Dispose();
                ds = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                sKeyword = null;
                foundRows = null;
                TopKeywords = null;
                objBoundingWord = null;
            }

            return Module1.objRetStructF3;
        }
        //private void CheckRightWord(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord)
        //{
        //    ClsCNC objCNC = new ClsCNC();
        //    clsCNCSBR ObjSbr = new clsCNCSBR();
        //    string[] FilterRightWord = { ":", ".", "#", "#:", "#." };
        //    int flag = 0;
        //    //Regex regex = new Regex("^([a-zA-Z]*[-,.#/]?[0-9]+[-,.#/]?[a-zA-Z]*[-,.#/]?)*$");
        //    //Regex regex1 = new Regex("^([a-zA-Z]*[-,./#%]?[a-zA-Z]*[-,./#%]?[0-9]+[-,./#%]?)*$");

        //    if (ObjBoundingWord.RightWord != null)
        //    {
        //        //if (FilterRightWord.Contains(ObjBoundingWord.RightWord.strWord.ToUpper()))
        //        //{
        //        //    if(ObjBoundingWord.RightWord.strWord=="#"|| ObjBoundingWord.RightWord.strWord == "#:"|| ObjBoundingWord.RightWord.strWord == "#.")
        //        //    {
        //        //        flag = 1;
        //        //    }
        //        //    ObjBoundingWord = objCNC.GetBoundingWords(ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
        //        //    goto Label3;
        //        //}

        //        string word = ObjBoundingWord.RightWord.strWord;

        //        RetStructIQSBR008 Obj008 = ObjSbr.IQSBR008(word, "E");

        //        //if ((Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.InvalidDate.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && Obj008.Status == "F" && ObjBoundingWord.RightWord.strWord.Length >= 5)
        //        if ((Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.InvalidDate.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && Obj008.Status == "F")
        //        {
        //            if (word.All(char.IsDigit))
        //            {
        //                clsCnCWord CncBLWord = ObjBoundingWord.RightWord;
        //                if (word.Length >= 4 && CncBLWord.strWord.EndsWith(","))
        //                {
        //                    CncBLWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncBLWord,",");
        //                }
        //                else
        //                {
        //                    CncBLWord = FindNextWordforNumeric(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncBLWord);
        //                    string strCncBLWord = CncBLWord.strWord.Replace(" ", "");
        //                    if (strCncBLWord.Length >= 4)
        //                    {
        //                        Boolean exist = CheckforExistWord(PWords, CncBLWord);
        //                        if (exist == false)
        //                        {
        //                            PWords.Add(CncBLWord);
        //                        }
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                if (word.Length >= 4)
        //                {
        //                    clsCnCWord CncBLWord = ObjBoundingWord.RightWord;
        //                    if (CncBLWord.strWord.EndsWith(","))
        //                    {
        //                        CncBLWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncBLWord,",");
        //                    }
        //                    Boolean exist = CheckforExistWord(PWords, CncBLWord);
        //                    if (exist == false)
        //                    {
        //                        PWords.Add(CncBLWord);
        //                    }
        //                }
        //            }
        //        }

        //        else
        //        {
        //            clsCnCWord TopWord = objCNC.MergeWords(TopKeywords);
        //            ObjBoundingWord = objCNC.GetBoundingWords(TopWord, ObjMetaData, intCurrPageNumber);
        //            clsCnCBoundingWords NewObjBoundingWord;
        //            if (ObjBoundingWord.BottomWords != null)
        //            {
        //                for (int t = 0; t < ObjBoundingWord.BottomWords.Length; t++)
        //                {
        //                    clsCnCWord CncSNWord = ObjBoundingWord.BottomWords[t];
        //                    word = CncSNWord.strWord;

        //                    Obj008 = ObjSbr.IQSBR008(word, "E");

        //                    if ((Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.InvalidDate.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && Obj008.Status == "F" && word.Length >= 5)
        //                    {
        //                        NewObjBoundingWord = objCNC.GetBoundingWords(CncSNWord, ObjMetaData, intCurrPageNumber);
        //                        if (NewObjBoundingWord!=null && NewObjBoundingWord.LeftWord != null)
        //                        {
        //                            if ((NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':'))
        //                            {
        //                                Boolean exist = CheckforExistWord(PWords, CncSNWord);
        //                                if (exist == false)
        //                                {
        //                                    PWords.Add(CncSNWord);
        //                                    //j = ObjBoundingWord.BottomWords.Length;
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            Boolean exist = CheckforExistWord(PWords, CncSNWord);
        //                            if (exist == false)
        //                            {
        //                                PWords.Add(CncSNWord);
        //                                // j = ObjBoundingWord.BottomWords.Length;
        //                            }
        //                        }
        //                        //Boolean exist = CheckforExistWord(PWords, CncSNWord);
        //                        //if (exist == false)
        //                        //{
        //                        //    PWords.Add(CncSNWord);
        //                        //    t = ObjBoundingWord.BottomWords.Length;
        //                        //}
        //                    }

        //                }
        //            }

        //            else
        //            {
        //                clsCnCWord BOL = FindBottomfromCoordinates(ObjMetaData, intCurrPageNumber, ObjMetaData.Page[intCurrPageNumber].Line[iLine], ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord], 5, 1,TopKeywords);
        //                if (BOL.strWord != null)
        //                {
        //                    NewObjBoundingWord = objCNC.GetBoundingWords(BOL, ObjMetaData, intCurrPageNumber);
        //                    if (NewObjBoundingWord!=null && NewObjBoundingWord.LeftWord != null)
        //                    {
        //                        if ((NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':'))
        //                        {
        //                            Boolean exist = CheckforExistWord(PWords, BOL);
        //                            if (exist == false)
        //                            {
        //                                PWords.Add(BOL);
        //                                // j = ObjBoundingWord.BottomWords.Length;
        //                            }
        //                        }
        //                    }
        //                    else
        //                    {
        //                        Boolean exist = CheckforExistWord(PWords, BOL);
        //                        if (exist == false)
        //                        {
        //                            PWords.Add(BOL);
        //                            //j = ObjBoundingWord.BottomWords.Length;
        //                        }
        //                    }

        //                    //Boolean exist = CheckforExistWord(PWords, BOL);
        //                    //if (exist == false)
        //                    //{
        //                    //    PWords.Add(BOL);

        //                    //}
        //                }
        //            }
        //        }
        //    }

        //    else
        //    {
        //        clsCnCWord TopWord = objCNC.MergeWords(TopKeywords);
        //        ObjBoundingWord = objCNC.GetBoundingWords(TopWord, ObjMetaData, intCurrPageNumber);
        //        clsCnCBoundingWords NewObjBoundingWord;
        //        if (ObjBoundingWord.BottomWords != null)
        //        {
        //            CheckBottomWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, PWords, TopKeywords, iLine, iWord);
        //        }
        //        else
        //        {
        //            clsCnCWord BOL = FindBottomfromCoordinates(ObjMetaData, intCurrPageNumber, ObjMetaData.Page[intCurrPageNumber].Line[iLine], ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord], 5, 1,TopKeywords);
        //            if (BOL.strWord != null)
        //            {
        //                NewObjBoundingWord = objCNC.GetBoundingWords(BOL, ObjMetaData, intCurrPageNumber);
        //                if (NewObjBoundingWord!=null &&  NewObjBoundingWord.LeftWord != null)
        //                {
        //                    if ((NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':'))
        //                    {
        //                        Boolean exist = CheckforExistWord(PWords, BOL);
        //                        if (exist == false)
        //                        {
        //                            PWords.Add(BOL);
        //                            // j = ObjBoundingWord.BottomWords.Length;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    Boolean exist = CheckforExistWord(PWords, BOL);
        //                    if (exist == false)
        //                    {
        //                        PWords.Add(BOL);
        //                        //j = ObjBoundingWord.BottomWords.Length;
        //                    }
        //                }

        //            }
        //        }
        //    }
        //}

        //private void CheckBottomWord(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord)
        //{
        //    ClsCNC objCNC = new ClsCNC();
        //    clsCNCSBR ObjSbr = new clsCNCSBR();
        //    Regex regex = new Regex("^([a-zA-Z]*[-,.#/]?[0-9]+[-,.#/]?[a-zA-Z]*[-,.#/]?)*$");
        //    Regex regex1 = new Regex("^([a-zA-Z]*[-,./#%]?[a-zA-Z]*[-,./#%]?[0-9]+[-,./#%]?)*$");
        //    //Regex InvalidDate = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.]*(Map|MAP|mAP|MAp)(\d{4}|(\d{3}[a-zA-Z])|(\d{1}[a-zA-Z]\d{2}))");
        //    //Regex Time = new Regex(@"^(?:[01]?[0-9]|2[0-3]):[0-5][0-9]$");
        //    clsCnCBoundingWords NewObjBoundingWord;
        //    for (int j = 0; j < ObjBoundingWord.BottomWords.Length; j++)
        //    {
        //        clsCnCWord CncSNWord = ObjBoundingWord.BottomWords[j];
        //        string word = CncSNWord.strWord;

        //        RetStructIQSBR008 Obj008 = ObjSbr.IQSBR008(word, "E");

        //        if ((regex.IsMatch(word) || regex1.IsMatch(word)) && (!Module1.InvalidDate.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && Obj008.Status == "F" && word.Length >= 5)
        //        {
        //            NewObjBoundingWord = objCNC.GetBoundingWords(CncSNWord, ObjMetaData, intCurrPageNumber);
        //            if (NewObjBoundingWord!=null&&NewObjBoundingWord.LeftWord != null)
        //            {
        //                if (NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':')
        //                {
        //                    Boolean exist = CheckforExistWord(PWords, CncSNWord);
        //                    if (exist == false)
        //                    {
        //                        PWords.Add(CncSNWord);
        //                        j = ObjBoundingWord.BottomWords.Length;
        //                    }
        //                }
        //            }
        //            else
        //            {
        //                Boolean exist = CheckforExistWord(PWords, CncSNWord);
        //                if (exist == false)
        //                {
        //                    PWords.Add(CncSNWord);
        //                    j = ObjBoundingWord.BottomWords.Length;
        //                }
        //            }
        //        }

        //        else
        //        {
        //            clsCnCWord BOL = FindBottomfromCoordinates(ObjMetaData, intCurrPageNumber, ObjMetaData.Page[intCurrPageNumber].Line[iLine], ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord], 5, 1,TopKeywords);
        //            if (BOL.strWord != null)
        //            {
        //                NewObjBoundingWord = objCNC.GetBoundingWords(BOL, ObjMetaData, intCurrPageNumber);
        //                if (NewObjBoundingWord!=null && NewObjBoundingWord.LeftWord != null)
        //                {
        //                    if ((NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':'))
        //                    {
        //                        Boolean exist = CheckforExistWord(PWords, BOL);
        //                        if (exist == false)
        //                        {
        //                            PWords.Add(BOL);
        //                            // j = ObjBoundingWord.BottomWords.Length;
        //                        }
        //                    }
        //                }
        //                else
        //                {
        //                    Boolean exist = CheckforExistWord(PWords, BOL);
        //                    if (exist == false)
        //                    {
        //                        PWords.Add(BOL);
        //                        //j = ObjBoundingWord.BottomWords.Length;
        //                    }
        //                }
        //            }

        //            //Boolean exist = CheckforExistWord(PWords, BOL);
        //            //    if (exist == false)
        //            //    {
        //            //        PWords.Add(BOL);

        //            //    }

        //        }
        //    }
        //}
        #endregion

        #region Purchase Order Number
        private RetStructF3 PONumber(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            DataSet ds = Module1.MakeDs(oMeta, intCurrPageNumber);
            DataTable dt = new DataTable();

            Module1.conflevel.Clear();
            Module1.PossibleWords.Clear();

            int iRecursiveCallCount = 0;
            string sKeyword;
            //clsCnCWord[] PoNumber = new clsCnCWord[2];
            clsCnCWord PoNoMerged = new clsCnCWord();
            string[] FilterKeywords = { "BOX","B0X", "OFFICE", "TYPE", "BOX#", "TERMS","LISTED" };

            System.Data.DataRow[] foundRows;
            clsCnCWord[] TopKeywords = new clsCnCWord[10];
            clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();

            try
            {

            Line1:

                if (iRecursiveCallCount == 0)
                {
                    // -------------------------------------------------PURCHASE ORDER NUMBER KEYWORDS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                    sKeyword = "'Purchase Order Number','Purchase Order No','cust po#','po#', 'Customer PO Number', 'Customer PO No','purchase order','customer p.o. no','customer p.o. number', 'PO Number','P. O. NUMBER','Customers PO','Customer PO','P.O. #','Consisnees Refer.noe/po No','P.O. No','Purchase Order #','PURCHASEORDER#','Cust P.O. No','No/Customer PO#','PO No.(s)','P.O. NUMBER(S)','P.O.NUMBER(S)','Cust.P.O.#','PURCHASE-ORDER-NO','CUSTOMER PURCHASE ORDER','CUSTOMERPURCHASEORDER#','P/O #','CustomerPO#','PONumber','PO Reference','PO#(s)','Customer PO#','Customer POs:','P O #:','P/O#','P O. NUMBER(S)','PO NUM BER','PO NUMBE R:','SALES ORDER NBR/CUSTOMER P.O','CUSTOMER PO INFORMATION','CUSTOMER P O. NO.','P O NO','ORDER ID / PO','P0/ORDER NUMBERS','DELIVERY/PO NUMBER','CUSTOME R P.O.','SPECIAL INSTRUCTIONS / PURCHASE ORDER NUMBERS','PO/REFERENCE NO','P0 #','CUSTOMER P0','Purchase Ader','p/o','Po','Customer Order Number','Customer Order No','CUSTOMERORDERNUMBER','CUSTOMER ORDER NUM','CUSTOMER ORDER NO(S)','CUSTOMERORDER NUM','CUSTOMER ORDER N','CUST. ORDER NO.','CUSTOMER ORDER #','Cust order'";
                }
                // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                else
                {
                    sKeyword = "'Order', 'Order #','Order Number'";
                }

                //try
                //{
                Module1.objRetStructF3.Status = "F";
                Module1.objRetStructF3.ManualConfirmation = "Y";
                Module1.objRetStructF3.Flag = "Y";
                string[] FilterRightWord = { ":", ".", "#", "#:", "#.", "#·", "NUMBER", "NO.", "NO", "NUMBER:", "NO:", "PO", "NO.(S)", "|", "•", "-", "NUMBER(S)", "/","IS:","IS", "NUM", "NUM:" };
                string[] ExcludeLeftWords = { "SHIPPERS", "SHIPPER'S", "SHIPPING" };

                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSet1_keyword"))
                {
                    Module1.dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSet1_keyword"];
                }
                if (Module1.dsKeys != null)
                {
                    dt = Module1.dsKeys.Tables[0];

                    //System.Data.DataRow[] foundRows;
                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber);
                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    //---For multiPage
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND  [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    //var orderedRows = foundRows.OrderBy(item => item.ItemArray[2]);
                    var orderedRows = foundRows.OrderBy(item => item.ItemArray[1]).ThenBy(item => item.ItemArray[2]).ThenBy(item => item.ItemArray[3]);//oreder by only Page No then by Line no
                    foundRows = orderedRows.ToArray();
                    //var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();
                    if (foundRows.Length > 1)
                    {
                        DataRow[] SortedRows = sortFoundRowsonKeywordPreference(sKeyword, foundRows, dt);
                        if(SortedRows.Length==foundRows.Length)
                        {
                            foundRows = SortedRows;
                        }
                    }

                    //var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    //foundRows = orderedRows.ToArray();
                    string keyword = string.Empty;
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int LineWordNo = 0;
                    int x1 = 0;
                    int y1 = 0;
                    int x2 = 0;
                    int y2 = 0;
                    //int h = 0;
                    int w = 0;



                    //if (foundRows.Length == 0 && iRecursiveCallCount == 0)
                    //{
                    //    iRecursiveCallCount = iRecursiveCallCount + 1;
                    //    if (iRecursiveCallCount == 1)
                    //    {
                    //        goto Line1;
                    //    }
                    //}
                    //else if (foundRows.Length == 1)
                    if (foundRows.Length == 1)
                    {
                        Array.Clear(TopKeywords, 0, TopKeywords.Length);
                        //Array.Clear(PoNumber, 0, PoNumber.Length);
                        keyword = (foundRows[0].ItemArray[0]).ToString();
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        LineWordNo = Convert.ToInt32(foundRows[0].ItemArray[4]);
                        x1 = Convert.ToInt32(foundRows[0].ItemArray[5]);
                        y1 = Convert.ToInt32(foundRows[0].ItemArray[6]);
                        x2 = Convert.ToInt32(foundRows[0].ItemArray[7]);
                        y2 = Convert.ToInt32(foundRows[0].ItemArray[8]);
                        w = oMeta.Page[pgno].ImageWidth;


                        if ((Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75)) < 1000)
                        {
                            y2 = y2;
                        }
                        else
                        {
                            y2 = Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.90);
                        }

                        //Words = oClsCnc.GetDataForROI(oMeta.Page, pgno, 0, 0, w, y2);
                        //clsCnCWord[] Words =Module1.oClsCnc.GetDataForROI(oMeta.Page, pgno, 0, 0, 9999, 9999);
                        clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);
                        if (Lines != null)
                        {
                            if (Lines.Length <= lineno)
                            {
                                //iRecursiveCallCount = iRecursiveCallCount + 1;
                                //if (iRecursiveCallCount == 1)
                                //{
                                //    goto Line1;
                                //}
                                //Module1.objRetStructF3.Status = "F";
                                Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                Module1.objRetStructF3.Words = Module1.PossibleWords;
                                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                Module1.conflevel.Add(90);
                                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                Module1.objRetStructF3.Status = "S";
                                //Module1.objRetStructF3.ManualConfirmation = "Y";
                                Module1.objRetStructF3.ManualConfirmation = "N";
                                Module1.objRetStructF3.Flag = "Y";
                                return Module1.objRetStructF3;
                            }
                        }
                        else
                        {
                            //iRecursiveCallCount = iRecursiveCallCount + 1;
                            //if (iRecursiveCallCount == 1)
                            //{
                            //    goto Line1;
                            //}
                            //Module1.objRetStructF3.Status = "F";
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                            Module1.objRetStructF3.Words = Module1.PossibleWords;
                            Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            Module1.conflevel.Add(90);
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            Module1.objRetStructF3.Status = "S";
                            //Module1.objRetStructF3.ManualConfirmation = "Y";
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Flag = "Y";
                            return Module1.objRetStructF3;
                        }


                        //clsCnCWord[] TopKeywords = new clsCnCWord[10];
                        //Array.Clear(Module1.TopKeywords, 0, Module1.TopKeywords.Length);
                        string[] SplitKeyword = keyword.Split();
                        int i;
                        int KeywordLength = keyword.Split().Length;
                        string lastWordOfKeyword = SplitKeyword[KeywordLength - 1];
                        for (i = 0; i < KeywordLength; i++)
                        {
                            TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                        }
                        clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);

                        if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                        {
                            ContainsKeyword(oMeta, pgno, TopKeywords, TopWord, 2, iRecursiveCallCount, keyword,KeywordLength);
                        }
                        else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1&&TopKeywords[KeywordLength-1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                        {
                            ContainsKeyword(oMeta, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 2, iRecursiveCallCount, keyword.Split()[KeywordLength - 1],KeywordLength);
                        }
                        else
                        {

                            if (keyword.ToUpper() == "ORDER" && objBoundingWord.LeftWord != null && ExcludeLeftWords.Contains(objBoundingWord.LeftWord.strWord.ToUpper()))
                            {
                                //Module1.objRetStructF3.Status = "F";
                                Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                Module1.objRetStructF3.Words = Module1.PossibleWords;
                                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                Module1.conflevel.Add(90);
                                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                Module1.objRetStructF3.Status = "S";
                                //Module1.objRetStructF3.ManualConfirmation = "Y";
                                Module1.objRetStructF3.ManualConfirmation = "N";
                                Module1.objRetStructF3.Flag = "Y";
                                return Module1.objRetStructF3;
                            }


                        Label1: if (objBoundingWord != null && objBoundingWord.RightWord != null)
                            {
                                if (!(FilterKeywords.Contains(objBoundingWord.RightWord.strWord.ToUpper())))
                                {
                                    string word = objBoundingWord.RightWord.strWord;
                                    if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                    {
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                        goto Label1;
                                    }

                                    //CheckRightWord(oMeta, objBoundingWord, intCurrPageNumber, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                    CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 2, iRecursiveCallCount, keyword);
                                }
                                else if (FilterKeywords.Contains(objBoundingWord.RightWord.strWord.ToUpper()) && (objBoundingWord.RightWord.strWord.All(char.IsUpper))&&(objBoundingWord.RightWord.strWord.ToUpper()=="OFFICE"))
                                {
                                    string word = objBoundingWord.RightWord.strWord;
                                    if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                    {
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                        goto Label1;
                                    }

                                    //CheckRightWord(oMeta, objBoundingWord, intCurrPageNumber, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                    CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 2, iRecursiveCallCount, keyword);
                                }
                                //else
                                //{
                                //    iRecursiveCallCount = iRecursiveCallCount + 1;
                                //    if (iRecursiveCallCount == 1)
                                //    {
                                //        goto Line1;
                                //    }
                                //}
                            }
                            else if (LineWordNo <= oMeta.Page[pgno].Line[lineno].WordCount)
                            {
                                // clsCnCWord BOL = new clsCnCWord();
                                objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                if (objBoundingWord.BottomWords != null)
                                {
                                    CheckBottomWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 2, iRecursiveCallCount,keyword);
                                }
                                else
                                {
                                    PoNoMerged = FindBottomfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], TopWord, 4, 2, TopKeywords, iRecursiveCallCount,keyword);
                                    if (PoNoMerged.strWord != null)
                                    {
                                        Boolean exist = CheckforExistWord(Module1.PossibleWords, PoNoMerged);
                                        if (exist == false)
                                        {
                                            Module1.PossibleWords.Add(PoNoMerged);
                                        }
                                    }
                                }

                            }
                        }
                        if (Module1.PossibleWords.Count == 0)
                        {
                            //iRecursiveCallCount = iRecursiveCallCount + 1;
                            //if (iRecursiveCallCount == 1)
                            //{
                            //    goto Line1;
                            //}

                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                            //Module1.objRetStructF3.Words = PossibleWords;
                            //Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                            //conflevel.Add(90);
                            //Module1.objRetStructF3.ConfidenceLevelofSuspect = conflevel;
                            //Module1.objRetStructF3.Status = "S";
                            //Module1.objRetStructF3.ManualConfirmation = "N";
                            //Module1.objRetStructF3.Flag = "Y";
                        }

                        Module1.objRetStructF3 = GetValues(oMeta,pgno, 2, iRecursiveCallCount,TopKeywords);
                        return Module1.objRetStructF3;

                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            Array.Clear(TopKeywords, 0, TopKeywords.Length);
                            //Array.Clear(PoNumber, 0, PoNumber.Length);
                            keyword = (drDataRow.ItemArray[0]).ToString();
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            LineWordNo = Convert.ToInt32(drDataRow.ItemArray[4]);
                            x1 = Convert.ToInt32(drDataRow.ItemArray[5]);
                            y1 = Convert.ToInt32(drDataRow.ItemArray[6]);
                            x2 = Convert.ToInt32(drDataRow.ItemArray[7]);
                            y2 = Convert.ToInt32(drDataRow.ItemArray[8]);

                            w = oMeta.Page[pgno].ImageWidth;

                            //if (iQPUBLIC.PublicComponents.ImageHeight < 1000)
                            //{
                            //    y2 = y2;
                            //}
                            //else
                            //{
                            //    y2 = iQPUBLIC.PublicComponents.ImageHeight;
                            //}

                            if ((Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75)) < 1000)
                            {
                                y2 = y2;
                            }
                            else
                            {
                                y2 = Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.90);
                            }

                            clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);

                            if ((Lines != null))
                            {
                                if (Lines.Length+1 < lineno)
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                continue;
                            }
                            //clsCnCWord[] TopKeywords = new clsCnCWord[10];
                            string[] SplitKeyword = keyword.Split();
                            int KeywordLength = keyword.Split().Length;
                            string lastWordOfKeyword = SplitKeyword[KeywordLength - 1];
                            int i;
                            for (i = 0; i < KeywordLength; i++)
                            {
                                TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];

                            }

                            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);

                            if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                            {
                                ContainsKeyword(oMeta, pgno, TopKeywords, TopWord, 2, iRecursiveCallCount, keyword,KeywordLength);
                            }
                            else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1 && TopKeywords[KeywordLength - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                            {
                                ContainsKeyword(oMeta, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 2, iRecursiveCallCount, keyword.Split()[KeywordLength - 1],KeywordLength);
                            }
                            else
                            {
                                if (keyword.ToUpper() == "ORDER" && objBoundingWord.LeftWord != null && ExcludeLeftWords.Contains(objBoundingWord.LeftWord.strWord.ToUpper()))
                                {
                                    continue;
                                }


                            //if ((objBoundingWord.RightWord != null) && regex.IsMatch(objBoundingWord.RightWord.strWord.Replace(":", "")) || (objBoundingWord.RightWord != null))
                            Label1: if (objBoundingWord != null && objBoundingWord.RightWord != null)
                                {
                                    if (!(FilterKeywords.Contains(objBoundingWord.RightWord.strWord.ToUpper())))
                                    {
                                        string word = objBoundingWord.RightWord.strWord;
                                        if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                        {
                                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                            goto Label1;
                                        }

                                        // CheckRightWord(oMeta, objBoundingWord, intCurrPageNumber, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                        CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 2, iRecursiveCallCount,keyword);
                                    }
                                }
                                else if (LineWordNo <= oMeta.Page[pgno].Line[lineno].WordCount)
                                {
                                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                    if (objBoundingWord.BottomWords != null)
                                    {
                                        CheckBottomWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 2, iRecursiveCallCount,keyword);
                                    }
                                    else
                                    {
                                        PoNoMerged = FindBottomfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], TopWord, 4, 2, TopKeywords, iRecursiveCallCount,keyword);
                                        if (PoNoMerged.strWord != null)
                                        {
                                            Boolean exist = CheckforExistWord(Module1.PossibleWords, PoNoMerged);
                                            if (exist == false)
                                            {
                                                Module1.PossibleWords.Add(PoNoMerged);
                                            }
                                        }
                                    }

                                    //if (objBoundingWord.TopWords != null)
                                    //{
                                    //    CheckTopWord(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo);
                                    //}


                                }
                            }
                        } // end of for 

                        if (Module1.PossibleWords.Count == 0)
                        {
                            //iRecursiveCallCount = iRecursiveCallCount + 1;
                            //if (iRecursiveCallCount == 1)
                            //{
                            //    goto Line1;
                            //}
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));

                        }

                        Module1.objRetStructF3 = GetValues(oMeta,pgno ,2, iRecursiveCallCount,TopKeywords);
                        return Module1.objRetStructF3;

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
                //dt.Clear();
                dt.Dispose();
                dt = null;
                sKeyword = null;
                foundRows = null;
                TopKeywords = null;
                PoNoMerged = null;
                objBoundingWord = null;
            }

            return Module1.objRetStructF3;
        }
        // private void CheckRightWordForPO(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord, clsCnCWord[] PoNumber)
        private void CheckRightWordForPO(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord, int field, int iRecursiveCallCount,string keyword)
        {
            //ClsCNC objCNC = new ClsCNC();
            //clsCNCSBR ObjSbr = new clsCNCSBR();
            //clsCnCWord PoNoMerged = new clsCnCWord();
            List<string> Separators = new List<string>() { ",", "&", "AND", "/" };
            char ch = ',';
            string[] FilterKeywords = { "BOX","B0X", "OFFICE", "TYPE", "BOX#" };
            string[] FilterleftWord = { ":", "#", "#:", "#.", "#·" };
            clsCnCWord CncKeyword = Module1.oClsCnc.MergeWords(TopKeywords);
            List<string> FilterkeywordsforAlpha = new List<string>() { "CUST ORDER", "CUSTOMER ORDER #" };
            List<string> WordNotStartWith = new List<string>() {"DATE","PAGE","PIECE","SALE" };
            List<string> WordNotContains = new List<string>() { "DIGIT","PAGE"};

            if (ObjBoundingWord.RightWord != null)
            {
                string word = ObjBoundingWord.RightWord.strWord;
                string WordToChckLength = string.Empty;
                word = word.Replace("'", "");
                Boolean FilterwordsPO = false;
                Boolean FlagShipment = false;
                if (word.Length > 2)
                {
                    if (Module1.SpecialCharacters.Contains(word[0].ToString()) && (!word.StartsWith("(")))
                    {
                        word = word.Remove(0, 1);
                    }
                    if (Module1.SpecialCharacters.Any(word.EndsWith) && (!word.EndsWith(")")))
                    {
                        word = word.Remove(word.Length - 1, 1);
                    }
                }
                
                clsCnCBoundingWords LeftobjBoundingWord = Module1.oClsCnc.GetBoundingWords(ObjBoundingWord.RightWord, ObjMetaData, intCurrPageNumber);
                int DiffHash = 0, DiffbetValueandkeyword = 0;
                if (LeftobjBoundingWord != null && LeftobjBoundingWord.LeftWord.strWord != null)
                {
                     DiffHash = LeftobjBoundingWord.LeftWord.Left - CncKeyword.Right;
                     DiffbetValueandkeyword = ObjBoundingWord.RightWord.Left - LeftobjBoundingWord.LeftWord.Right;
                }
               
                if ((!Module1.SpecialCharacters.Contains(word)))
                {
                    if (!Module1.Time.IsMatch(word))
                    {
                        WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                        word = word.Replace(":","");
                        word = Module1.ReplaceBraces(word);
                        int NoofDash = word.Count(c => !char.IsLetterOrDigit(c));
                        if (word.Count(c => (c == '-')) > 1)
                        {
                            word = word.Replace("-", "");
                            // word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                        }
                        //word = word.Replace("(", "");
                        //word = word.Replace(")", "");
                    }
                }
                RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");
                if (Obj008.Status == "S")
                {
                    if (word.All(char.IsLetterOrDigit))
                    {
                        Obj008.Status = "F";
                    }
                }

                if (field == 2)
                {
                    if ((FilterKeywords.Any(word.ToUpper().StartsWith))&&(!word.All(char.IsUpper)&&(!word.ToUpper().StartsWith("OFFICE"))))
                    {
                        FilterwordsPO = true;
                    }
                }
                if(field==3&&(keyword.ToUpper()=="SHIPMENT"|| keyword.ToUpper() == "SHIPMENTS"))
                {
                    if (word.Count(char.IsLetter) >= word.Count(char.IsDigit))
                    {
                        FlagShipment = true;
                    }
                }

                int NumberOfSpecialCharaters = word.Count(c => !char.IsLetterOrDigit(c));
                int NumberofCharandNum = word.Count(c => char.IsLetterOrDigit(c));
                //Boolean A = ((FilterleftWord.Contains(LeftobjBoundingWord.LeftWord.strWord.ToUpper())) || (CncKeyword.strWord.EndsWith(":") || CncKeyword.strWord.EndsWith("#")));
                //Boolean B = (DiffHash <= 40);
                //Boolean C = (iRecursiveCallCount != 1);
                //Boolean D = (WordToChckLength.Length >= 2);
                //Boolean E = (FlagShipment == false);
                //Boolean F = (DiffbetValueandkeyword <= 250);
                //Boolean G = (!FilterkeywordsforAlpha.Contains(keyword.ToUpper()));
                //Boolean H = (WordToChckLength.All(char.IsUpper));
                //Boolean I = (!Module1.RAKeywords.Contains(keyword.ToUpper()));

                Boolean A = (!string.IsNullOrEmpty(word));
                Boolean B = (!(Module1.NewRegexForAlphanumeric.IsMatch(word)));
                Boolean C = (!WordToChckLength.All(char.IsLetter));
                Boolean D = (!Module1.Time.IsMatch(word));
                Boolean E = (!Module1.RegTel.IsMatch(word));
                Boolean F = (Obj008.Status == "F");
                Boolean G = (word != Module1.ProNo);
                Boolean H = (WordToChckLength.Length >= 2);
                Boolean I = (FlagShipment == false);
                Boolean J = (NumberOfSpecialCharaters < NumberofCharandNum);
                Boolean K = (!WordNotStartWith.Any(word.ToUpper().StartsWith));
                Boolean L = (!WordNotContains.Any(word.ToUpper().Contains));
                if (FilterwordsPO == false)
                {
                    //if (!string.IsNullOrEmpty(word))
                    //{
                    //if ((!string.IsNullOrEmpty(word)) && (Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && (Obj008.Status == "F") && (word != Module1.ProNo) &&( WordToChckLength.Length >= 2) &&( FlagShipment == false)&&(NumberOfSpecialCharaters<NumberofCharandNum))
                    if ((!string.IsNullOrEmpty(word)) && (!(Module1.NewRegexForAlphanumeric.IsMatch(word)))&&(!WordToChckLength.All(char.IsLetter)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(ObjBoundingWord.RightWord.strWord)) && (Obj008.Status == "F") && (word != Module1.ProNo) && (WordToChckLength.Length >= 2) && (FlagShipment == false) && (NumberOfSpecialCharaters < NumberofCharandNum)&&(!WordNotStartWith.Any(word.ToUpper().StartsWith))&&(!WordNotContains.Any(word.ToUpper().Contains)))
                    {

                        //if (word.All(char.IsDigit))
                        if (ObjBoundingWord.RightWord.strWord.All(char.IsDigit))
                        {
                            clsCnCWord CncPoWord = ObjBoundingWord.RightWord;
                            clsCnCBoundingWords Newbounding = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                            int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                            int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                            if (WordToChckLength.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1 && Newbounding != null && Newbounding.RightWord != null)
                            {
                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, ",", iRecursiveCallCount,field);
                                Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                if (exist == false)
                                {
                                    PWords.Add(CncPoWord);
                                }
                            }
                            else if (WordToChckLength.Length >= 4 && CncPoWord.strWord.EndsWith("/") && SlashCount == 1 && Newbounding != null && Newbounding.RightWord != null)
                            {
                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, "/", iRecursiveCallCount,field);
                                Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                if (exist == false)
                                {
                                    PWords.Add(CncPoWord);
                                }
                            }
                            else if (WordToChckLength.Length >= 4 && Newbounding != null && Newbounding.RightWord != null && Separators.Contains(Newbounding.RightWord.strWord.ToUpper()))
                            {
                                foreach (string sepratorStr in Separators)
                                {
                                    if (Newbounding.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                    {
                                        CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, sepratorStr, iRecursiveCallCount,field);
                                        string StrKeyword = Regex.Replace(CncKeyword.strWord, @"[^0-9a-zA-Z]+", "");
                                        if (Module1.RAKeywords.Contains(keyword.ToUpper()))     //give prefix for RA,RGA,RMA
                                        {
                                            CncPoWord = AddPrefix(CncPoWord, keyword);
                                        }

                                        CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                                        Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                        if (exist == false)
                                        {
                                            PWords.Add(CncPoWord);
                                        }
                                        break;
                                    }
                                }

                            }
                            else
                            {
                                clsCnCWord  MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord,keyword);
                                //if(Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ","")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                {
                                    CncPoWord = MergedCncPoWord;
                                }
                                string strCncWord = CncPoWord.strWord.Replace(" ", "");
                                if (strCncWord.Length > 1)
                                {
                                    if (Module1.SpecialCharacters.Contains(strCncWord[0].ToString()))
                                    {
                                        strCncWord = strCncWord.Remove(0, 1);
                                    }
                                    if (Module1.SpecialCharacters.Any(strCncWord.EndsWith))
                                    {
                                        strCncWord = strCncWord.Remove(strCncWord.Length - 1, 1);
                                    }
                                }
                                if (strCncWord.Length >= 2)
                                {
                                    //CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord);
                                    if (Module1.SpecialCharacters.Contains(CncPoWord.strWord[0].ToString()) && (!CncPoWord.strWord.StartsWith("(")))
                                    {
                                        CncPoWord = ModifyWord(CncPoWord, CncPoWord.strWord.Remove(0, 1));
                                    }
                                    string StrKeyword = Regex.Replace(CncKeyword.strWord, @"[^0-9a-zA-Z]+", "");
                                    if (Module1.RAKeywords.Contains(keyword.ToUpper()))     //give prefix for RA,RGA,RMA
                                    {
                                        //CncPoWord = AddPrefix(CncPoWord, CncKeyword.strWord);
                                        CncPoWord = AddPrefix(CncPoWord, keyword);
                                    }

                                    CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                                    if (iRecursiveCallCount == 1)
                                    {
                                        CncPoWord = SecondaryKeywordFlag(ObjMetaData, intCurrPageNumber, TopKeywords, CncPoWord,field);
                                    }
                                    Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                    if (exist == false)
                                    {
                                        PWords.Add(CncPoWord);
                                    }
                                }

                            }
                        }
                        else
                        {
                            //if (word.Count(char.IsLetter) <= word.Count(char.IsDigit))
                            //{
                            //if (word.Count(char.IsNumber) != 1)
                            //{
                            word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                            if (word.Length >= 2)
                            {
                                clsCnCWord CncPoWord = ObjBoundingWord.RightWord;
                                string strCncWord = CncPoWord.strWord;
                                clsCnCBoundingWords Newbounding = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                if (WordToChckLength.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1 && Newbounding != null && Newbounding.RightWord != null)
                                {
                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, ",", iRecursiveCallCount,field);
                                }
                                else if (WordToChckLength.Length >= 4 && CncPoWord.strWord.EndsWith("/") && SlashCount == 1 && Newbounding != null && Newbounding.RightWord != null)
                                {
                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, "/", iRecursiveCallCount,field);
                                }
                                else if (WordToChckLength.Length >= 4 && Newbounding != null && Newbounding.RightWord != null && Separators.Contains(Newbounding.RightWord.strWord.ToUpper()))
                                {
                                    foreach (string sepratorStr in Separators)
                                    {
                                        if (Newbounding.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                        {
                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, sepratorStr, iRecursiveCallCount,field);
                                           string Keyword = Regex.Replace(CncKeyword.strWord, @"[^0-9a-zA-Z]+", "");
                                            if (Module1.RAKeywords.Contains(keyword.ToUpper()))     //give prefix for RA,RGA,RMA
                                            {
                                                CncPoWord = AddPrefix(CncPoWord, keyword);
                                            }

                                            CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);

                                            Boolean exist1 = CheckforExistWord(PWords, CncPoWord);
                                            if (exist1 == false)
                                            {
                                                PWords.Add(CncPoWord);
                                            }
                                            break;
                                        }
                                    }

                                }
                                else if (WordToChckLength.Length >= 4 && strCncWord.Contains(",") || strCncWord.Contains("-") || strCncWord.Contains("/") || strCncWord.Contains(";"))  // 07/12/2020
                                {
                                    if (CncPoWord.strWord.Contains(","))
                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ",", iRecursiveCallCount); }
                                    else if (CncPoWord.strWord.Contains("-"))
                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "-", iRecursiveCallCount); }
                                    else if (CncPoWord.strWord.Contains("/"))
                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "/", iRecursiveCallCount); }
                                    else if (CncPoWord.strWord.Contains(";"))
                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ";", iRecursiveCallCount); }
                                    //if (Convert.ToInt32(CncPoWord.WordNumber) != 0)
                                    if(!CncPoWord.strWord.Contains("~"))
                                    {
                                        CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord,keyword);
                                        CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                                        if (iRecursiveCallCount == 1)
                                        {
                                            CncPoWord = SecondaryKeywordFlag(ObjMetaData, intCurrPageNumber, TopKeywords, CncPoWord,field);
                                        }
                                    }
                                }
                                else
                                {
                                   clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord,keyword);
                                    string strMergedWord = MergedCncPoWord.strWord;
                                    // MergedCncPoWord.strWord = Module1.ReplaceBraces(MergedCncPoWord.strWord);
                                    strMergedWord= Module1.ReplaceBraces(strMergedWord);

                                    //if (Module1.regex.IsMatch(strMergedWord.Replace(" ", "")) || Module1.regex1.IsMatch(strMergedWord.Replace(" ", "")))
                                    if (!Module1.NewRegexForAlphanumeric.IsMatch(strMergedWord.Replace(" ", "")))
                                    {
                                        CncPoWord = MergedCncPoWord;
                                    }
                                    CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                                    
                                    if (iRecursiveCallCount == 1)
                                    {
                                      CncPoWord=  SecondaryKeywordFlag(ObjMetaData, intCurrPageNumber, TopKeywords, CncPoWord,field);
                                    }
                                }

                                if (Module1.SpecialCharacters.Contains(CncPoWord.strWord[0].ToString()) && (!CncPoWord.strWord.StartsWith("(")))
                                {
                                    CncPoWord = ModifyWord(CncPoWord, CncPoWord.strWord.Remove(0, 1));
                                }

                                CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                                string StrKeyword = Regex.Replace(CncKeyword.strWord, @"[^0-9a-zA-Z]+", "");
                                if (Module1.RAKeywords.Contains(keyword.ToUpper()))   //give prefix for RA,RGA,RMA
                                {
                                    CncPoWord = AddPrefix(CncPoWord, keyword);
                                }
                                Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                if (exist == false)
                                {
                                    PWords.Add(CncPoWord);
                                }
                            }

                            //}
                        }
                    }

                    //else if ((FilterleftWord.Contains(LeftobjBoundingWord.LeftWord.strWord.ToUpper())) || (CncKeyword.strWord.EndsWith(":") || CncKeyword.strWord.EndsWith("#") || CncKeyword.strWord.EndsWith("#:"))&& (LeftobjBoundingWord.LeftWord.Left - CncKeyword.Right <= 40)&&(iRecursiveCallCount!=1))
                    else if (((LeftobjBoundingWord!=null&&LeftobjBoundingWord.LeftWord.strWord!=null&&(FilterleftWord.Contains(LeftobjBoundingWord.LeftWord.strWord.ToUpper())))|| (CncKeyword.strWord.EndsWith(":") || CncKeyword.strWord.EndsWith("#"))) && (DiffHash <= 40) && (iRecursiveCallCount != 1) && (WordToChckLength.Length >= 2) && (FlagShipment == false) && (DiffbetValueandkeyword <= 250) && (!FilterkeywordsforAlpha.Contains(keyword.ToUpper())) && (WordToChckLength.All(char.IsUpper))&&(!Module1.RAKeywords.Contains(keyword.ToUpper())))                    
                    {
                        Boolean BOLFlag = false;
                        if (keyword.ToUpper() == "BOL" &&(! LeftobjBoundingWord.LeftWord.strWord.Contains("#")))
                        {
                            BOLFlag = true;
                        }
                        if (BOLFlag == false)
                        {
                            clsCnCWord CncPoWord = ObjBoundingWord.RightWord;
                            clsCnCWord MergedCncPoWord = FindNextWordforAlpha(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord);
                            if (MergedCncPoWord.strWord.Replace(" ", "").ToUpper().All(char.IsLetter))
                            {
                                CncPoWord = MergedCncPoWord;
                            }
                            //else if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                            else if (!(Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", ""))))
                            {
                                CncPoWord = MergedCncPoWord;
                            }
                            string strCncWord = CncPoWord.strWord;
                            if (strCncWord.Length >= 2)
                            {
                                //CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord);
                                if (Module1.SpecialCharacters.Contains(CncPoWord.strWord[0].ToString()) && (!CncPoWord.strWord.StartsWith("(")))
                                {
                                    CncPoWord = ModifyWord(CncPoWord, CncPoWord.strWord.Remove(0, 1));
                                }
                                CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                                //string StrKeyword = Regex.Replace(CncKeyword.strWord, @"[^0-9a-zA-Z]+", "");
                                //if (Module1.RAKeywords.Contains(StrKeyword.ToUpper()))     //give prefix for RA,RGA,RMA
                                //{
                                //    CncPoWord = AddPrefix(CncPoWord, CncKeyword.strWord);
                                //}
                                List<clsCnCWord> PossibleValuesforAlpha = new List<clsCnCWord>();
                                PossibleValuesforAlpha.Add(CncPoWord);
                                //To check for bottom
                                clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                                ObjBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, ObjMetaData, intCurrPageNumber);

                                if (ObjBoundingWord.BottomWords != null)
                                {
                                    CheckBottomWordForPO(ObjMetaData, ObjBoundingWord, intCurrPageNumber, PossibleValuesforAlpha, TopKeywords, iLine, iWord, field, iRecursiveCallCount, keyword);
                                }
                                if(PossibleValuesforAlpha.Count>1)
                                {
                                    string strAlpha = CncPoWord.strWord.Replace(" ", "");
                                    if (strAlpha.Length >= 3)
                                    {
                                        CncPoWord = RetWordforAlpha(PossibleValuesforAlpha);
                                    }
                                    else
                                    {
                                        CncPoWord = PossibleValuesforAlpha[1];
                                    }
                                }
                                else
                                {
                                    string strAlpha = CncPoWord.strWord.Replace(" ", "");
                                    if (strAlpha.Length >= 3)
                                    {
                                        CncPoWord = PossibleValuesforAlpha[0];
                                    }
                                }
                                Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                if (exist == false)
                                {
                                    PWords.Add(CncPoWord);
                                }
                            }
                        }
                        else
                        {
                            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            ObjBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, ObjMetaData, intCurrPageNumber);

                            if (ObjBoundingWord.BottomWords != null)
                            {
                                CheckBottomWordForPO(ObjMetaData, ObjBoundingWord, intCurrPageNumber, PWords, TopKeywords, iLine, iWord, field, iRecursiveCallCount, keyword);
                            }
                        }
                    }
                    else
                    {
                        clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            ObjBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, ObjMetaData, intCurrPageNumber);

                        if (ObjBoundingWord.BottomWords != null)
                        {
                            if (keyword != "SHIPMENT")
                            {
                                CheckBottomWordForPO(ObjMetaData, ObjBoundingWord, intCurrPageNumber, PWords, TopKeywords, iLine, iWord, field, iRecursiveCallCount, keyword);
                            }
                        }
                        else
                        {
                            if (keyword != "SHIPMENT")
                            {
                                clsCnCWord CncPOWord = FindBottomfromCoordinates(ObjMetaData, intCurrPageNumber, ObjMetaData.Page[intCurrPageNumber].Line[iLine], ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[iWord], 4, field, TopKeywords, iRecursiveCallCount, keyword);
                                clsCnCBoundingWords NewObjBoundingWord;
                                if (CncPOWord.strWord != null)
                                {
                                    NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPOWord, ObjMetaData, intCurrPageNumber);
                                    int CommaCount = CncPOWord.strWord.Count(c => (c == ch));
                                    if (NewObjBoundingWord != null && NewObjBoundingWord.LeftWord != null)
                                    {
                                        //if ((NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':'))
                                        if (!Module1.IgnoreWordforLeftSplChar.Any(NewObjBoundingWord.LeftWord.strWord.EndsWith))
                                        {
                                            if (CncPOWord.strWord.EndsWith(",") && CommaCount == 1 && NewObjBoundingWord.RightWord != null)
                                            {
                                                CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, ",", iRecursiveCallCount,field);
                                            }
                                            if (Module1.SpecialCharacters.Contains(CncPOWord.strWord[0].ToString()) && (!CncPOWord.strWord.StartsWith("(")))
                                            {
                                                CncPOWord = ModifyWord(CncPOWord, CncPOWord.strWord.Remove(0, 1));
                                            }
                                            //if (!CncPOWord.strWord.Contains("~")&&(!Module1.SpecialCharacters.Any(CncPOWord.strWord.Contains)))
                                            //{
                                            //    if ((CncPOWord.strWord.Count(char.IsLetter) == 1) && (!char.IsLetter(CncPOWord.strWord[0])) && (!char.IsLetter(CncPOWord.strWord[CncPOWord.strWord.Length - 1])))
                                            //    {
                                            //        CncPOWord = Module1.ReplaceAlphanumericWord(CncPOWord);
                                            //    }
                                            //}
                                            CncPOWord.strWord = ReplaceValue(CncPOWord.strWord, field);
                                            Boolean exist = CheckforExistWord(PWords, CncPOWord);
                                            if (exist == false)
                                            {
                                                PWords.Add(CncPOWord);
                                                //j = ObjBoundingWord.BottomWords.Length;
                                            }

                                        }
                                    }
                                    else
                                    {
                                        if (CncPOWord.strWord.EndsWith(",") && CommaCount == 1 && NewObjBoundingWord.RightWord != null)
                                        {
                                            CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, ",", iRecursiveCallCount,field);
                                        }
                                        if (Module1.SpecialCharacters.Contains(CncPOWord.strWord[0].ToString()) && (!CncPOWord.strWord.StartsWith("(")))
                                        {
                                            CncPOWord = ModifyWord(CncPOWord, CncPOWord.strWord.Remove(0, 1));
                                        }
                                        //if (!CncPOWord.strWord.Contains("~")&& (!Module1.SpecialCharacters.Any(CncPOWord.strWord.Contains)))
                                        //{

                                        //    if ((CncPOWord.strWord.Count(char.IsLetter) == 1) && (!char.IsLetter(CncPOWord.strWord[0])) && (!char.IsLetter(CncPOWord.strWord[CncPOWord.strWord.Length - 1])))
                                        //    {
                                        //        CncPOWord = Module1.ReplaceAlphanumericWord(CncPOWord);
                                        //    }
                                        //}
                                        CncPOWord.strWord = ReplaceValue(CncPOWord.strWord, field);
                                        Boolean exist = CheckforExistWord(PWords, CncPOWord);
                                        if (exist == false)
                                        {
                                            PWords.Add(CncPOWord);
                                        }
                                    }
                                }
                            }
                        }
                    }
                    //}
                }
            }

            else
            {
                clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                ObjBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, ObjMetaData, intCurrPageNumber);
                if (ObjBoundingWord.BottomWords != null)
                {
                    if (keyword != "SHIPMENT")
                    {
                        CheckBottomWordForPO(ObjMetaData, ObjBoundingWord, intCurrPageNumber, PWords, TopKeywords, iLine, iWord, field, iRecursiveCallCount, keyword);
                    }
                }
            }
        }
        private void CheckBottomWordForPO(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord, int field, int iRecursiveCallCount,string keyword)
        {
            //ClsCNC objCNC = new ClsCNC();
            //clsCNCSBR ObjSbr = new clsCNCSBR();
            //clsCnCWord[] PoNumber = new clsCnCWord[5];
            // clsCnCWord PoNoMerged = new clsCnCWord();
            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);

            clsCnCWord topWord = Module1.oClsCnc.MergeWords(TopKeywords);
            int KeywordBottom = topWord.Bottom;
            int Keywordheight = TopWord.Bottom - TopWord.Top;
            List<string> Separators = new List<string>() { ",", "&", "AND", "/" };
            List<string> WordNotStartWith = new List<string>() { "DATE", "PAGE", "PIECE" };
            List<string> WordNotContains = new List<string>() { "DIGIT" };
            char ch = ',';
            string WordToChckLength = string.Empty;
            Boolean FlagShipment = false;

            for (int j = 0; j < ObjBoundingWord.BottomWords.Length; j++)
            {
                //int p = 0;
                clsCnCWord CncPOWord = ObjBoundingWord.BottomWords[j];
                string word = CncPOWord.strWord;
                word = word.Replace("'", "");
                clsCnCBoundingWords NewObjBoundingWord;
                if (word.Length > 2)
                {
                    // word = word.Replace("'", "");
                    if (Module1.SpecialCharacters.Contains(word[0].ToString()) && (!word.StartsWith("(")))
                    {
                        word = word.Remove(0, 1);
                    }
                    if (Module1.SpecialCharacters.Any(word.EndsWith) && (!word.EndsWith(")")))
                    {
                        word = word.Remove(word.Length - 1, 1);
                    }
                }
                
                if ((!Module1.SpecialCharacters.Contains(word)))
                {
                    if (!Module1.Time.IsMatch(word))
                    {
                        WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                        word = Module1.ReplaceBraces(word);
                        int NoofDash = word.Count(c => !char.IsLetterOrDigit(c));
                        if (word.Count(c => (c == '-')) > 1)
                        {
                            word=word.Replace("-","");
                           // word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                        }

                    }
                }
                RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");

                if (Obj008.Status == "S")
                {
                    if (word.All(char.IsLetterOrDigit))
                    {
                        Obj008.Status = "F";
                    }
                }
                if (field == 3 && (keyword.ToUpper() == "SHIPMENT" || keyword.ToUpper() == "SHIPMENTS"))
                {

                    if (word.Count(char.IsLetter) >= word.Count(char.IsDigit))
                    {
                        FlagShipment = true;
                    }
                }
                int NumberOfSpecialCharaters = word.Count(c => !char.IsLetterOrDigit(c));
                int NumberofCharandNum = word.Count(c => char.IsLetterOrDigit(c));

                Boolean A = (!(Module1.NewRegexForAlphanumeric.IsMatch(word))) && (!WordToChckLength.All(char.IsLetter));
                Boolean B = (!Module1.Time.IsMatch(word));
                Boolean C = (!Module1.RegTel.IsMatch(word));
                Boolean D = Obj008.Status == "F";
                Boolean E = WordToChckLength.Length >= 3;
                Boolean F = word != Module1.ProNo;
                Boolean G = FlagShipment == false;
                Boolean H = NumberOfSpecialCharaters < NumberofCharandNum;
                Boolean I = (!word.ToUpper().StartsWith("DATE"));
                

                //if ((regex.IsMatch(word) || (regex1.IsMatch(word))) && (!Module1.Time.IsMatch(word)) && Obj008.Status == "F" && (!word.ToUpper().Contains("BOX")) && word.Length >= 4)
                if (!string.IsNullOrEmpty(word))
                {
                    //if ((Module1.regex.IsMatch(word) || (Module1.regex1.IsMatch(word))) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && WordToChckLength.Length >= 3 && word != Module1.ProNo && FlagShipment == false&&NumberOfSpecialCharaters<NumberofCharandNum&& (!word.ToUpper().StartsWith("DATE")))
                    if ((!(Module1.NewRegexForAlphanumeric.IsMatch(word))) && (!WordToChckLength.All(char.IsLetter)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(CncPOWord.strWord)) && Obj008.Status != "S" && WordToChckLength.Length >= 3 && word != Module1.ProNo && FlagShipment == false && NumberOfSpecialCharaters < NumberofCharandNum && (!WordNotStartWith.Any(word.ToUpper().StartsWith)) && (!WordNotContains.Any(word.ToUpper().Contains)))
                    {
                        //if (word.Count(char.IsLetter) <= word.Count(char.IsDigit))
                        //{
                        //if (word.Count(char.IsDigit) != 1)
                        //{
                            int wordTop = CncPOWord.Top;
                            int height = wordTop - KeywordBottom;

                            if (height <= Keywordheight * 3)
                            {
                                NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPOWord, ObjMetaData, intCurrPageNumber);
                                int CommaCount = CncPOWord.strWord.Count(c => (c == ch));
                                int SlashCount = CncPOWord.strWord.Count(c => (c == '/'));
                            if (NewObjBoundingWord != null && NewObjBoundingWord.LeftWord != null)
                            {
                                //if (NewObjBoundingWord.LeftWord.strWord[NewObjBoundingWord.LeftWord.strWord.Length - 1] != ':')
                                if (!Module1.IgnoreWordforLeftSplChar.Any(NewObjBoundingWord.LeftWord.strWord.EndsWith))
                                {
                                    if (CncPOWord.strWord.All(char.IsDigit))
                                    {
                                        // clsCnCWord CncPoWord = NewLine.Word[iword];
                                        //int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                        if (WordToChckLength.Length >= 4 && CncPOWord.strWord.EndsWith(",") && CommaCount == 1)
                                        {
                                            CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, ",", iRecursiveCallCount,field);

                                        }
                                        else if (WordToChckLength.Length >= 4 && CncPOWord.strWord.EndsWith("/") && SlashCount == 1)
                                        {
                                            CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, "/", iRecursiveCallCount,field);

                                        }
                                        else if (WordToChckLength.Length >= 4 && NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                        {
                                            foreach (string sepratorStr in Separators)
                                            {
                                                if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                {
                                                    CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, sepratorStr, iRecursiveCallCount,field);
                                                    Boolean exist1 = CheckforExistWord(PWords, CncPOWord);
                                                    if (exist1 == false)
                                                    {
                                                        PWords.Add(CncPOWord);
                                                        j = ObjBoundingWord.BottomWords.Length;
                                                    }
                                                    break;
                                                }
                                            }

                                        }
                                        else
                                        {
                                            clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord,keyword);
                                            //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                            if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                            {
                                                CncPOWord = MergedCncPoWord;
                                            }
                                            string strCncBLWord = CncPOWord.strWord.Replace(" ", "");
                                            if (strCncBLWord.Length >= 4)
                                            {
                                                //CncPOWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord);
                                                CncPOWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, TopWord, CncPOWord, iRecursiveCallCount,field);
                                            }
                                        }
                                    }
                                    else
                                    {
                                        if (CncPOWord.strWord.EndsWith(",") && CommaCount == 1 && NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null)
                                        {
                                            CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, ",", iRecursiveCallCount,field);
                                        }
                                        else if (CncPOWord.strWord.EndsWith("/") && SlashCount == 1 && NewObjBoundingWord!=null&& NewObjBoundingWord.RightWord != null)
                                        {
                                            CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, "/", iRecursiveCallCount,field);
                                        }
                                        else if (NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                        {
                                            foreach (string sepratorStr in Separators)
                                            {
                                                if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                {
                                                    CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, sepratorStr, iRecursiveCallCount,field);
                                                    Boolean exist1 = CheckforExistWord(PWords, CncPOWord);
                                                    if (exist1 == false)
                                                    {
                                                        PWords.Add(CncPOWord);
                                                        j = ObjBoundingWord.BottomWords.Length;
                                                    }
                                                    break;
                                                }
                                            }

                                        }
                                        else if (CncPOWord.strWord.Contains(",") || CncPOWord.strWord.Contains("/") || CncPOWord.strWord.Contains(";"))  // 07/12/2020
                                        {
                                            if (CncPOWord.strWord.Contains(","))
                                            { CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, ",", iRecursiveCallCount); }
                                            //else if (CncPOWord.strWord.Contains("-"))
                                            //{ CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, "-", iRecursiveCallCount); }
                                            else if (CncPOWord.strWord.Contains("/"))
                                            { CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, "/", iRecursiveCallCount); }
                                            else if (CncPOWord.strWord.Contains(";"))
                                            { CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, ";", iRecursiveCallCount); }

                                            //if (Convert.ToInt32(CncPOWord.WordNumber) != 0)
                                            if (!CncPOWord.strWord.Contains("~"))
                                            {
                                                CncPOWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord,keyword);
                                            }
                                        }
                                        else
                                        {
                                            clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord,keyword);
                                            // if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                            if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                            {
                                                CncPOWord = MergedCncPoWord;
                                            }
                                            CncPOWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, TopWord, CncPOWord, iRecursiveCallCount,field);
                                        }
                                        //if (Module1.SpecialCharacters.Contains(CncPOWord.strWord[0].ToString()) && (!CncPOWord.strWord.StartsWith("(")))
                                        //{
                                        //    CncPOWord = ModifyWord(CncPOWord, CncPOWord.strWord.Remove(0, 1));
                                        //}

                                        //Boolean exist = CheckforExistWord(PWords, CncPOWord);
                                        //if (exist == false)
                                        //{
                                        //    PWords.Add(CncPOWord);
                                        //    j = ObjBoundingWord.BottomWords.Length;
                                        //}
                                    }
                                    if (Module1.SpecialCharacters.Contains(CncPOWord.strWord[0].ToString()) && (!CncPOWord.strWord.StartsWith("(")))
                                    {
                                        CncPOWord = ModifyWord(CncPOWord, CncPOWord.strWord.Remove(0, 1));
                                    }
                                    //if (Module1.RAKeywords.Contains())      //give prefix for RA,RGA,RMA
                                    string StrKeyword = Regex.Replace(topWord.strWord, @"[^0-9a-zA-Z]+", "");
                                    if (Module1.RAKeywords.Contains(keyword.ToUpper()))
                                    {
                                        CncPOWord = AddPrefix(CncPOWord, keyword);
                                    }
                                    CncPOWord.strWord = ReplaceValue(CncPOWord.strWord, field);
                                    Boolean exist = CheckforExistWord(PWords, CncPOWord);
                                    if (exist == false)
                                    {
                                        if (keyword.ToUpper() == "SALES ORDER NBR/CUSTOMER P.O")
                                        {
                                            CncPOWord.Flag = "0";
                                            CncPOWord.Remarks = "NA";
                                        }
                                        PWords.Add(CncPOWord);
                                        j = ObjBoundingWord.BottomWords.Length;
                                    }
                                }
                                else
                                { break; }
                            }
                            else
                            {
                                if (CncPOWord.strWord.All(char.IsDigit))
                                {
                                    // clsCnCWord CncPoWord = NewLine.Word[iword];
                                    //int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                    // if (WordToChckLength.Length >= 4 && CncPOWord.strWord.EndsWith(",") && CommaCount == 1)
                                    if (WordToChckLength.Length >= 4 && CncPOWord.strWord.EndsWith(",") && CommaCount == 1 && NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null)
                                    {
                                        CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, ",", iRecursiveCallCount,field);

                                    }
                                    else if (WordToChckLength.Length >= 4 && CncPOWord.strWord.EndsWith("/") && SlashCount == 1)
                                    {
                                        CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, "/", iRecursiveCallCount,field);

                                    }
                                    else if (WordToChckLength.Length >= 4 && NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                    {
                                        foreach (string sepratorStr in Separators)
                                        {
                                            if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                            {
                                                CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, sepratorStr, iRecursiveCallCount,field);
                                                Boolean exist1 = CheckforExistWord(PWords, CncPOWord);
                                                if (exist1 == false)
                                                {
                                                    PWords.Add(CncPOWord);
                                                    j = ObjBoundingWord.BottomWords.Length;
                                                }
                                                break;
                                            }
                                        }

                                    }
                                    else
                                    {
                                        clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord,keyword);
                                        string ReplaceBracesStr = Module1.ReplaceBraces(MergedCncPoWord.strWord.Replace(" ", ""));
                                        //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                        //if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                        if (!Module1.NewRegexForAlphanumeric.IsMatch(ReplaceBracesStr))
                                        {
                                            CncPOWord = MergedCncPoWord;
                                        }
                                        string strCncBLWord = CncPOWord.strWord.Replace(" ", "");
                                        if (strCncBLWord.Length >= 4)
                                        {
                                            CncPOWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, TopWord, CncPOWord, iRecursiveCallCount,field);
                                        }
                                    }
                                }
                                else
                                {
                                    if (CncPOWord.strWord.EndsWith(",") && CommaCount == 1 && NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null)
                                    {
                                        CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, ",", iRecursiveCallCount,field);
                                    }
                                    else if (CncPOWord.strWord.EndsWith("/") && SlashCount == 1 && NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null)
                                    {
                                        CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, "/", iRecursiveCallCount,field);
                                    }
                                    else if (NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                    {
                                        foreach (string sepratorStr in Separators)
                                        {
                                            if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                            {
                                                CncPOWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord.LineNo, CncPOWord, sepratorStr, iRecursiveCallCount,field);
                                                Boolean exist1 = CheckforExistWord(PWords, CncPOWord);
                                                if (exist1 == false)
                                                {
                                                    PWords.Add(CncPOWord);
                                                    j = ObjBoundingWord.BottomWords.Length;
                                                }
                                                break;
                                            }
                                        }

                                    }
                                    else if (CncPOWord.strWord.Contains(",") || CncPOWord.strWord.Contains("/") || CncPOWord.strWord.Contains(";"))  // 15/12/2020
                                    {
                                        if (CncPOWord.strWord.Contains(","))
                                        { CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, ",", iRecursiveCallCount); }
                                        //else if (CncPOWord.strWord.Contains("-"))
                                        //{ CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, "-", iRecursiveCallCount); }
                                        else if (CncPOWord.strWord.Contains("/"))
                                        { CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, "/", iRecursiveCallCount); }
                                        else if (CncPOWord.strWord.Contains(";"))
                                        { CncPOWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPOWord, ";", iRecursiveCallCount); }

                                        //if (Convert.ToInt32(CncPOWord.WordNumber) != 0)
                                        if (!CncPOWord.strWord.Contains("~"))
                                        {
                                            CncPOWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord,keyword);
                                            CncPOWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, TopWord, CncPOWord, iRecursiveCallCount,field);
                                        }
                                    }
                                    else
                                    {
                                        clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, CncPOWord.LineNo, CncPOWord,keyword);
                                        string ReplaceBracesStr = Module1.ReplaceBraces(MergedCncPoWord.strWord.Replace(" ", ""));
                                        //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                        //if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                        if (!Module1.NewRegexForAlphanumeric.IsMatch(ReplaceBracesStr))
                                        {
                                            CncPOWord = MergedCncPoWord;
                                        }
                                        CncPOWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, TopWord, CncPOWord, iRecursiveCallCount,field);
                                    }
                                }
                                if (Module1.SpecialCharacters.Contains(CncPOWord.strWord[0].ToString()) && (!CncPOWord.strWord.StartsWith("(")))
                                {
                                    CncPOWord = ModifyWord(CncPOWord, CncPOWord.strWord.Remove(0, 1));
                                }
                                string StrKeyword = Regex.Replace(topWord.strWord, @"[^0-9a-zA-Z]+", "");
                                //if (Module1.RAKeywords.Contains(StrKeyword.ToUpper()))     //give prefix for RA,RGA,RMA
                                //{
                                //    CncPOWord = AddPrefix(CncPOWord, topWord.strWord);
                                //}
                                CncPOWord.strWord = ReplaceValue(CncPOWord.strWord, field);
                                Boolean exist = CheckforExistWord(PWords, CncPOWord);
                                if (exist == false)
                                {
                                    if (keyword.ToUpper() == "SALES ORDER NBR/CUSTOMER P.O")
                                    {
                                        CncPOWord.Flag = "0";
                                        CncPOWord.Remarks = "NA";
                                    }
                                    PWords.Add(CncPOWord);
                                    j = ObjBoundingWord.BottomWords.Length;
                                }
                            }
                            }
                        //}
                    }
                    
                    else
                    {
                        if (!Module1.SpecialCharacters.Contains(word))
                        {
                            break;
                        }
                    }
                }

            }
        }

        #endregion

        #region Shipper Number       
        private RetStructF3 ShipperNumber(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        {

            DataSet ds = Module1.MakeDs(oMeta, intCurrPageNumber);
           // RetStructF3 objRetStructF3 = new RetStructF3();

            // RetStructF3 returnZones = new RetStructF3();

            //RetStructIQSBR008 objRetStructIQSBR008 = new RetStructIQSBR008();

            DataTable dt = new DataTable();

            Module1.PossibleWords.Clear();
            Module1.conflevel.Clear();

            int iRecursiveCallCount = 0;
            string sKeyword;
            clsCnCWord ShipperNO = new clsCnCWord();
            //clsCnCWord CncSNWord = new clsCnCWord();

            string[] FilterRightWord = { ":", ".", "#", "#:", "#.", "#·", "|", "•", "-", "/", "NUM", "NUM:" };

            System.Data.DataRow[] foundRows;
            clsCnCWord[] TopKeywords = new clsCnCWord[50];
            clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();

            try
            {
                ////Line1:

                //if (iRecursiveCallCount == 0)
                //{
                //    // -------------------------------------------------SHIPPER NUMBER KEYWORDS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                //    sKeyword = "'Shipper Number','Ship ID','Shipper Ref', 'SHIPPERS NUMBER', 'Shippers No', 'Shipper No','Shipment', 'Shipping Order No','Shipper No/Customer','Shipment ID','Shipment No','Shipment number','Shipment Reference Number','Shippers Release No','Ship Ref','SHIPPER NUM','Shipment numbers','SHIPPER REF NUMBER','Shipper NOS','BOL/SHIPMENT NO','shpmt num','shipperno','Ship ID#:','SH PPER NUMBER','SHIPPER#','SHIPMEN T NUMBER','SHIPMEN T NUMBE R ','Shipper ID','SHIPPERNOS','Shippers #','SHIPPER #','BOL/SHIPMENT#','SHIPPER REF #/ORIG BOL #','SHIP NO','SHIPMENT IDENTIFICATION NO','SHIPPERS REFERENCE','BILL OF LADING NUMBER/SHIPMENT NUMBER','SHIPPER REFERENCE NO'";
                //}
                //// -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                //else
                //{
                //    sKeyword = "'SHIPPER','Shippers'";
                //}

                //try
                //{
                Module1.objRetStructF3.Status = "F";
                Module1.objRetStructF3.ManualConfirmation = "Y";
                Module1.objRetStructF3.Flag = "Y";
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSet1_keyword"))
                {
                    Module1.dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSet1_keyword"];
                }
                if (Module1.dsKeys != null)
                {
                    //Regex regex = new Regex("^[^/w/d]$");                    
                    dt = Module1.dsKeys.Tables[0];
                    for (int iPage = 1; iPage <= oMeta.PageCount; iPage++)
                    //for (int iPage = intCurrPageNumber; iPage <= intCurrPageNumber; iPage++)
                    {
                        //Line1:

                        if (iRecursiveCallCount == 0)
                        {
                            // -------------------------------------------------SHIPPER NUMBER KEYWORDS-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------

                            //Commented list to remove shipment id keyword
                            //sKeyword = "'Shipper Number','Ship ID','Shipper Ref', 'SHIPPERS NUMBER', 'Shippers No', 'Shipper No','Shipment', 'Shipping Order No','Shipper No/Customer','Shipment ID','Shipment No','Shipment number','Shipment Reference Number','Shipment Reference','Shippers Release No','Ship Ref','SHIPPER NUM','SHIPMENT NUMBER(S)','SHIPPER REF NUMBER','Shipper NOS','BOL/SHIPMENT NO','shpmt num','shipperno','Ship ID#:','SH PPER NUMBER','SHIPPER#','SHIPMEN T NUMBER','SHIPMEN T NUMBE R ','Shipper ID','SHIPPERNOS','Shippers #','SHIPPER #','BOL/SHIPMENT#','SHIPPER REF #/ORIG BOL #','SHIP NO','SHIPMENT IDENTIFICATION NO','SHIPPERS REFERENCE','BILL OF LADING NUMBER/SHIPMENT NUMBER','SHIPPER REFERENCE NO','SHIPPERSNO','Shipper Reference','Shipment #'";
                            sKeyword = "'Shipper Number','Shipper Ref', 'SHIPPERS NUMBER', 'Shippers No', 'Shipper No','Shipment', 'Shipping Order No','Shipper No/Customer','Shipment No','Shipment number','Shipment Reference Number','Shipment Reference','Shippers Release No','Ship Ref','SHIPPER NUM','SHIPMENT NUMBER(S)','SHIPPER REF NUMBER','Shipper NOS','BOL/SHIPMENT NO','shpmt num','shipperno','SH PPER NUMBER','SHIPPER#','SHIPMEN T NUMBER','SHIPMEN T NUMBE R ','SHIPPERNOS','Shippers #','SHIPPER #','BOL/SHIPMENT#','SHIPPER REF #/ORIG BOL #','SHIP NO','SHIPMENT IDENTIFICATION NO','SHIPPERS REFERENCE','BILL OF LADING NUMBER/SHIPMENT NUMBER','SHIPPER REFERENCE NO','SHIPPERSNO','Shipper Reference','Shipment #'";
                        }
                        // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
                        else
                        {
                            sKeyword = "'SHIPPER','Shippers'";
                        }
                        // foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
                        // foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber);
                        // foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                        //---For multiPage
                        //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND  [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                        foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + iPage + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);

                        var orderedRows = foundRows.OrderBy(item => item.ItemArray[1]).ThenByDescending(item => item.ItemArray[0]);//order by only Page No then by keywords
                        foundRows = orderedRows.ToArray();
                        //var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                        //foundRows = orderedRows.ToArray();
                        string keyword = string.Empty;
                        int pgno = 0;
                        int lineno = 0;
                        int wordno = 0;
                        int LineWordNo = 0;
                        int x1 = 0;
                        int y1 = 0;
                        int x2 = 0;
                        int y2 = 0;
                        //int h = 0;
                        int w = 0;


                        //clsCnCWord sJoinWord = new clsCnCWord();

                        if (foundRows.Length == 1)
                        {
                            Array.Clear(TopKeywords, 0, TopKeywords.Length);
                            // Array.Clear(PoNumber, 0, PoNumber.Length);
                            keyword = (foundRows[0].ItemArray[0]).ToString();
                            pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                            lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                            wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                            LineWordNo = Convert.ToInt32(foundRows[0].ItemArray[4]);
                            x1 = Convert.ToInt32(foundRows[0].ItemArray[5]);
                            y1 = Convert.ToInt32(foundRows[0].ItemArray[6]);
                            x2 = Convert.ToInt32(foundRows[0].ItemArray[7]);
                            y2 = Convert.ToInt32(foundRows[0].ItemArray[8]);
                            w = oMeta.Page[pgno].ImageWidth;

                            if ((Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75)) < 1000)
                            {
                                y2 = y2;
                            }
                            else
                            {
                                y2 = Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75);
                            }

                            //Words = oClsCnc.GetDataForROI(oMeta.Page, pgno, 0, 0, w, y2);
                            //clsCnCWord[] Words = oClsCnc.GetDataForROI(oMeta.Page, pgno, 0, 0,9999, 9999);
                            clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);
                            if (Lines != null)
                            {
                                if (Lines.Length <= lineno)
                                {
                                    //Module1.objRetStructF3.Status = "F";
                                    //iRecursiveCallCount = iRecursiveCallCount + 1;
                                    //if (iRecursiveCallCount == 1)
                                    //{
                                    //    //goto Line1;
                                    //}
                                    //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                    //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                    //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                    //Module1.conflevel.Add(90);
                                    //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                    //Module1.objRetStructF3.Status = "S";
                                    ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                    //Module1.objRetStructF3.ManualConfirmation = "N";
                                    //Module1.objRetStructF3.Flag = "Y";
                                    //return Module1.objRetStructF3;
                                    continue;
                                }
                            }
                            else
                            {
                                //Module1.objRetStructF3.Status = "F";
                                //iRecursiveCallCount = iRecursiveCallCount + 1;
                                //if (iRecursiveCallCount == 1)
                                //{
                                //   // goto Line1;
                                //}
                                //Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                //Module1.objRetStructF3.Words = Module1.PossibleWords;
                                //Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                //Module1.conflevel.Add(90);
                                //Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                //Module1.objRetStructF3.Status = "S";
                                ////Module1.objRetStructF3.ManualConfirmation = "Y";
                                //Module1.objRetStructF3.ManualConfirmation = "N";
                                //Module1.objRetStructF3.Flag = "Y";
                                //return Module1.objRetStructF3;
                                continue;
                            }

                            // clsCnCWord[] TopKeywords = new clsCnCWord[10];
                            string[] SplitKeyword = keyword.Split();
                            int i;
                            int KeywordLength = keyword.Split().Length;
                            string lastWordOfKeyword = SplitKeyword[KeywordLength - 1];
                            for (i = 0; i < KeywordLength; i++)
                            {
                                TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                            }

                            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);
                            //if (keyword.ToUpper() == "SHIPMENT")
                            //{
                            //    if (objBoundingWord.RightWord != null)
                            //    {
                            //        string[] ShipmentRight = { "#", "#:", "#." };
                            //        if (!ShipmentRight.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                            //        {
                            //            goto Label4;
                            //        }
                            //    }
                            //    else
                            //    { goto Label4; }
                            //}
                            if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                            {
                                ContainsKeyword(oMeta, pgno, TopKeywords, TopWord, 3, iRecursiveCallCount, keyword, KeywordLength);
                            }
                            else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1 && TopKeywords[KeywordLength - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                            {
                                ContainsKeyword(oMeta, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 3, iRecursiveCallCount, keyword.Split()[KeywordLength - 1], KeywordLength);
                            }
                            else
                            {

                            Label3: if ((objBoundingWord.RightWord != null))
                                {
                                    if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                    {
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                        goto Label3;
                                    }
                                    CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 3, iRecursiveCallCount, keyword);

                                }
                                else
                                {
                                    TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                    if (objBoundingWord.BottomWords != null)
                                    {
                                        if (keyword != "SHIPMENT")
                                        {
                                            CheckBottomWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 3, iRecursiveCallCount, keyword);
                                        }

                                    }
                                    else
                                    {
                                        if (iRecursiveCallCount != 1)
                                        {
                                            ShipperNO = FindBottomfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], oMeta.Page[pgno].Line[lineno].Word[LineWordNo], 4, 3, TopKeywords, iRecursiveCallCount, keyword);
                                            if (ShipperNO.strWord != null)
                                            {
                                                Boolean exist = CheckforExistWord(Module1.PossibleWords, ShipperNO);
                                                if (exist == false)
                                                {
                                                    Module1.PossibleWords.Add(ShipperNO);

                                                }
                                            }
                                        }
                                    }
                                }
                            }

                        //Label4: if (Module1.PossibleWords.Count == 0)
                        //    {
                        //        //   iRecursiveCallCount = iRecursiveCallCount + 1;
                        //        //if (iRecursiveCallCount == 1)
                        //        //{
                        //        //   // goto Line1;
                        //        //}

                        //        Module1.PossibleWords.Add(StringToCncWord(1, "*****"));

                        //    }
                            //Module1.objRetStructF3 = GetValues(oMeta, pgno, 3, iRecursiveCallCount, TopKeywords);
                            //return Module1.objRetStructF3;
                            
                        }
                        else
                        {
                            foreach (DataRow drDataRow in foundRows)
                            {
                                Array.Clear(TopKeywords, 0, TopKeywords.Length);
                                keyword = (drDataRow.ItemArray[0]).ToString();
                                pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                                lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                                wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                                LineWordNo = Convert.ToInt32(drDataRow.ItemArray[4]);
                                x1 = Convert.ToInt32(drDataRow.ItemArray[5]);
                                y1 = Convert.ToInt32(drDataRow.ItemArray[6]);
                                x2 = Convert.ToInt32(drDataRow.ItemArray[7]);
                                y2 = Convert.ToInt32(drDataRow.ItemArray[8]);

                                w = oMeta.Page[pgno].ImageWidth;


                                if ((Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75)) < 1000)
                                {
                                    y2 = y2;
                                }
                                else
                                {
                                    y2 = Convert.ToInt32(oMeta.Page[pgno].ImageHeight * 0.75);
                                }

                                //clsCnCWord[] TopKeywords = new clsCnCWord[10];
                                string[] SplitKeyword = keyword.Split();
                                int KeywordLength = keyword.Split().Length;
                                string lastWordOfKeyword = SplitKeyword[KeywordLength - 1];
                                int i;
                                for (i = 0; i < KeywordLength; i++)
                                {
                                    TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];

                                }

                                clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);

                                if (Lines != null)
                                {
                                    if (Lines.Length <= lineno)
                                    {
                                        continue;
                                    }
                                }
                                else
                                {
                                    continue;
                                }

                                clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                                objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);
                                //if ((objBoundingWord.RightWord != null) && regex.IsMatch(objBoundingWord.RightWord.strWord.Replace(":", "")) || (objBoundingWord.RightWord != null))

                                //if (keyword.ToUpper() == "SHIPMENT")
                                //{
                                //    if (objBoundingWord.RightWord != null)
                                //    {
                                //        string[] ShipmentRight = { "#", "#:", "#.", "#·" };
                                //        if (!ShipmentRight.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                //        {
                                //            continue;
                                //        }
                                //    }
                                //    else { continue; }
                                //}

                                if (TopWord.strWord.Any(char.IsDigit) && KeywordLength == 1 && TopWord.strWord.Count(char.IsNumber) > 1)
                                {
                                    ContainsKeyword(oMeta, pgno, TopKeywords, TopWord, 3, iRecursiveCallCount, keyword, KeywordLength);
                                }
                                else if (TopKeywords[KeywordLength - 1].strWord.Any(char.IsDigit) && TopKeywords[KeywordLength - 1].strWord.Count(char.IsNumber) > 1&& TopKeywords[KeywordLength - 1].strWord.ToUpper().StartsWith(lastWordOfKeyword))
                                {
                                    ContainsKeyword(oMeta, pgno, TopKeywords, TopKeywords[KeywordLength - 1], 3, iRecursiveCallCount, keyword.Split()[KeywordLength - 1], KeywordLength);
                                }
                                else
                                {
                                Label3: if (objBoundingWord != null && objBoundingWord.RightWord != null)
                                    {
                                        if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                        {
                                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
                                            goto Label3;
                                        }

                                        CheckRightWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 3, iRecursiveCallCount, keyword);

                                    }
                                    else
                                    {
                                        TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                        if (objBoundingWord.BottomWords != null)
                                        {
                                            if (keyword != "SHIPMENT")
                                            {
                                                CheckBottomWordForPO(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 3, iRecursiveCallCount, keyword);
                                            }
                                        }
                                        else
                                        {
                                            if (iRecursiveCallCount != 1)
                                            {
                                                ShipperNO = FindBottomfromCoordinates(oMeta, pgno, oMeta.Page[pgno].Line[lineno], oMeta.Page[pgno].Line[lineno].Word[LineWordNo], 4, 3, TopKeywords, iRecursiveCallCount, keyword);
                                                if (ShipperNO.strWord != null)
                                                {
                                                    Boolean exist = CheckforExistWord(Module1.PossibleWords, ShipperNO);
                                                    if (exist == false)
                                                    {
                                                        Module1.PossibleWords.Add(ShipperNO);

                                                    }
                                                }
                                            }
                                        }


                                    }
                                }
                            } // end of for 


                            //if (Module1.PossibleWords.Count == 0)
                            //{
                            //    //iRecursiveCallCount = iRecursiveCallCount + 1;
                            //    //if (iRecursiveCallCount == 1)
                            //    //{
                            //    //    //goto Line1;
                            //    //}

                            //    Module1.PossibleWords.Add(StringToCncWord(1, "*****"));

                            //}

                            //Module1.objRetStructF3 = GetValues(oMeta, pgno, 3, iRecursiveCallCount, TopKeywords);
                            //return Module1.objRetStructF3;

                            
                        }


                        // end of the F3
                    }
                    if (Module1.PossibleWords.Count == 0)
                    {
                        //iRecursiveCallCount = iRecursiveCallCount + 1;
                        //if (iRecursiveCallCount == 1)
                        //{
                        //    //goto Line1;
                        //}

                        Module1.PossibleWords.Add(StringToCncWord(1, "*****"));

                    }
                    Module1.objRetStructF3 = GetValues(oMeta, intCurrPageNumber, 3, iRecursiveCallCount, TopKeywords);
                    return Module1.objRetStructF3;
                }
            }
            catch (Exception)
            {

                //throw;
            }
            finally
            {
                // ds.Clear();
                ds.Dispose();
                ds = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                sKeyword = null;
                foundRows = null;
                TopKeywords = null;
                objBoundingWord = null;
                ShipperNO = null;
            }

            return Module1.objRetStructF3;
        }
        #endregion

        #region Hazmat Telephone
        private RetStructF3 Tel_HazMat(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            DataSet ds = Module1.MakeDs(oMeta, intCurrPageNumber);

            DataTable dt = new DataTable();

            Module1.PossibleWords.Clear();
            Module1.conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword = string.Empty;
            string vendorSKeyword = string.Empty;

            System.Data.DataRow[] foundRows;
            clsCnCWord[] TopKeywords = new clsCnCWord[50];
            clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();

            List<string> VendorEmerNo = new List<string>() { "8004249300", "8005355053", "8002553924", "18004249300", "18005355053", "18002553924" };
            try
            {

                if (iRecursiveCallCount == 0)
                {
                    // -------------------------------------------------HAZMAT TELEPHONE NUMBER KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------                
                    sKeyword = "'Emergency','Emergency Phone Numbers','Emergency Response Phone#','Emergency Contact','Emergency Response Phone Number','EMERGENCYCONTACTNUMBER:','24hrEMERGENCY PHONENUMBERS','CHEM EMER #','EMERGENCY CONTACT CALL:','EMERGENCY PHONE','chemical Emergencies','Emergency Response','Emergency TEL','Emergency Contact Phone','Emergencies'";
                    vendorSKeyword = "'Chemtrec','Infotrac','Chemtel','Chem tel'";
                }


                //try
                //{
                //    Module1.objRetStructF3.Status = "F";
                Module1.objRetStructF3.ManualConfirmation = "Y";
                Module1.objRetStructF3.Flag = "Y";

                List<string> lstEmerVendors = new List<string>() { "CHEMTREC", "CHEMTEL", "INFOTRAC", "CHEM TEL" };
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSet1_keyword"))
                {
                    Module1.dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSet1_keyword"];
                }
                if (Module1.dsKeys != null)
                {

                    dt = Module1.dsKeys.Tables[0];
                    //System.Data.DataRow[] foundRows;
                    //foundRows= dt.Select("Keyword IN " + "(" + vendorSKeyword + ") AND [Page No]=1");
                    //foundRows = dt.Select("Keyword IN " + "(" + vendorSKeyword + ") AND [Page No]="+intCurrPageNumber);
                    foundRows = dt.Select("Keyword IN " + "(" + vendorSKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    if (foundRows.Length == 0)
                    {
                        //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber);
                        foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                        //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
                    }
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();

                    string keyword = string.Empty;
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int LineWordNo = 0;
                    int x1 = 0;
                    int y1 = 0;
                    int x2 = 0;
                    int y2 = 0;
                    //int h = 0;
                    int w = 0;


                    if (foundRows.Length == 1)
                    {
                        Array.Clear(TopKeywords, 0, TopKeywords.Length);
                        keyword = (foundRows[0].ItemArray[0]).ToString();
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        LineWordNo = Convert.ToInt32(foundRows[0].ItemArray[4]);
                        x1 = Convert.ToInt32(foundRows[0].ItemArray[5]);
                        y1 = Convert.ToInt32(foundRows[0].ItemArray[6]);
                        x2 = Convert.ToInt32(foundRows[0].ItemArray[7]);
                        y2 = Convert.ToInt32(foundRows[0].ItemArray[8]);
                        w = oMeta.Page[intCurrPageNumber].ImageWidth;


                        if (oMeta.Page[intCurrPageNumber].ImageHeight < 1000)
                        {
                            y2 = y2;
                        }
                        else
                        {
                            y2 = oMeta.Page[intCurrPageNumber].ImageHeight;
                        }

                        clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);

                        if (Lines != null)
                        {
                            if (Lines.Length < lineno)
                            {
                                //Module1.objRetStructF3.Status = "F";
                                Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                Module1.objRetStructF3.Words = Module1.PossibleWords;
                                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                Module1.conflevel.Add(90);
                                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                Module1.objRetStructF3.Status = "S";
                                //Module1.objRetStructF3.ManualConfirmation = "Y";
                                Module1.objRetStructF3.ManualConfirmation = "N";
                                Module1.objRetStructF3.Flag = "Y";
                                return Module1.objRetStructF3;
                            }
                        }
                        else
                        {
                            //Module1.objRetStructF3.Status = "F";
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                            Module1.objRetStructF3.Words = Module1.PossibleWords;
                            Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            Module1.conflevel.Add(90);
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            Module1.objRetStructF3.Status = "S";
                            //Module1.objRetStructF3.ManualConfirmation = "Y";
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Flag = "Y";
                            return Module1.objRetStructF3;
                        }

                        // clsCnCWord[] TopKeywords = new clsCnCWord[50];
                        string[] SplitKeyword = keyword.Split();
                        int i;
                        int KeywordLength = keyword.Split().Length;
                        for (i = 0; i < KeywordLength; i++)
                        {
                            TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                        }

                        clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);


                        if (lstEmerVendors.Contains(TopWord.strWord.ToUpper()))
                        {
                            String EmerTel = string.Empty;
                            if (TopWord.strWord.ToUpper() == "CHEMTREC")
                            {
                                EmerTel = "18004249300";
                            }
                            else if (TopWord.strWord.ToUpper() == "INFOTRAC")
                            {
                                EmerTel = "18005355053";
                            }
                            else if (TopWord.strWord.ToUpper() == "CHEMTEL" || TopWord.strWord.ToUpper() == "CHEM TEL")
                            {
                                EmerTel = "18002553924";
                            }
                            //clsCnCWord CncTel = StringToCncWord(intCurrPageNumber, EmerTel);
                            clsCnCWord CncTel = ModifyWord(TopWord, EmerTel);
                            CncTel.Flag = "1";
                            CncTel.Remarks = "3F";
                            CncTel.Confidence = 100;
                            Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                            if (exist == false)
                            {
                                Module1.PossibleWords.Add(CncTel);
                            }
                            // Module1.PossibleWords.Add(StringToCncWord(intCurrPageNumber, EmerTel));
                        }
                        else
                        {
                            //string tel = FindTel(oMeta, pgno, lineno);                      
                            string Tel;

                        Label1: if (objBoundingWord.RightWord != null)
                            {
                                if (Module1.ExcludeWords.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                {
                                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, intCurrPageNumber);
                                    goto Label1;
                                }
                                Tel = objBoundingWord.RightWord.strWord.Replace(",", "");
                                Tel = objBoundingWord.RightWord.strWord.Replace("#", "");
                                if (Module1.RegTel.IsMatch(Tel))
                                {
                                    clsCnCWord CncTel = objBoundingWord.RightWord;
                                    Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                    if (exist == false)
                                    {
                                        Module1.PossibleWords.Add(objBoundingWord.RightWord);
                                    }
                                }
                                else
                                {
                                    clsCnCWord ReturnTel = FindTel(oMeta, pgno, lineno, TopKeywords, KeywordLength);
                                    if (ReturnTel.strWord != null)
                                    {
                                        Boolean exist = CheckforExistWord(Module1.PossibleWords, ReturnTel);
                                        if (exist == false)
                                        {
                                            Module1.PossibleWords.Add(ReturnTel);
                                        }
                                    }
                                }

                                //else
                                //{                                
                                //    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                //    if (objBoundingWord.BottomWords != null)
                                //    {
                                //        for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
                                //        {
                                //            clsCnCWord CncTel = objBoundingWord.BottomWords[t];
                                //            Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
                                //            Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

                                //            if (Module1.RegTel.IsMatch(Tel))
                                //            {
                                //                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                //                if (exist == false)
                                //                {
                                //                    Module1.PossibleWords.Add(CncTel);
                                //                    t = objBoundingWord.BottomWords.Length;
                                //                }
                                //            }

                                //        }
                                //    }
                                //}
                            }
                            else
                            {
                                //clsCnCWord BOL = new clsCnCWord();
                                //objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                //if (objBoundingWord.BottomWords != null)
                                //{
                                //    for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
                                //    {
                                //        clsCnCWord CncTel = objBoundingWord.BottomWords[t];
                                //        Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
                                //        Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

                                //        if (Module1.RegTel.IsMatch(Tel))
                                //        {
                                //            Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                //            if (exist == false)
                                //            {
                                //                Module1.PossibleWords.Add(CncTel);
                                //                t = objBoundingWord.BottomWords.Length;
                                //            }
                                //        }

                                //    }
                                //}
                            }
                            if (Module1.PossibleWords.Count == 0)
                            {
                                objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                if (objBoundingWord.BottomWords != null)
                                {
                                    for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
                                    {
                                        clsCnCWord CncTel = objBoundingWord.BottomWords[t];
                                        Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
                                        Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

                                        if (Module1.RegTel.IsMatch(Tel))
                                        {
                                            Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                            if (exist == false)
                                            {
                                                Module1.PossibleWords.Add(CncTel);
                                                t = objBoundingWord.BottomWords.Length;
                                            }
                                        }

                                    }
                                }

                                //clsCnCWord ReturnTel = FindTel(oMeta, pgno, lineno, TopKeywords,KeywordLength);
                                //if (ReturnTel.strWord != null)
                                //{
                                //    Boolean exist = CheckforExistWord(Module1.PossibleWords, ReturnTel);
                                //    if (exist == false)
                                //    {
                                //        Module1.PossibleWords.Add(ReturnTel);
                                //    }
                                //}
                            }

                        }
                        if (Module1.PossibleWords.Count == 0)
                        {
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                        }

                        if (Module1.PossibleWords.Count == 1)
                        {
                            if (Module1.PossibleWords[0].Flag == null)
                            {
                                //Boolean ConfidenceFlag = HighestConfidence(Module1.PossibleWords[0]);
                                Boolean ConfidenceFlag = HighestConfidenceforHazTel(Module1.PossibleWords[0]);
                                if (ConfidenceFlag == true)
                                {
                                    Module1.PossibleWords[0].Confidence = 100;
                                    Module1.PossibleWords[0].Flag = "1";
                                    Module1.PossibleWords[0].Remarks = "1F";
                                }
                                else
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                
                                //Module1.PossibleWords[0].Flag = "1";
                                //Module1.PossibleWords[0].Remarks = "1F";
                            }
                            Module1.objRetStructF3.Words = Module1.PossibleWords;
                            Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            Module1.conflevel.Add(90);
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            Module1.objRetStructF3.Flag = "Y";
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Status = "S";
                            //Module1.EMerTel = Module1.PossibleWords[0].strWord;
                            return Module1.objRetStructF3;
                        }
                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            Array.Clear(TopKeywords, 0, TopKeywords.Length);
                            keyword = (drDataRow.ItemArray[0]).ToString();
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            LineWordNo = Convert.ToInt32(drDataRow.ItemArray[4]);
                            x1 = Convert.ToInt32(drDataRow.ItemArray[5]);
                            y1 = Convert.ToInt32(drDataRow.ItemArray[6]);
                            x2 = Convert.ToInt32(drDataRow.ItemArray[7]);
                            y2 = Convert.ToInt32(drDataRow.ItemArray[8]);

                            w = oMeta.Page[intCurrPageNumber].ImageWidth;

                            if (oMeta.Page[intCurrPageNumber].ImageHeight < 1000)
                            {
                                y2 = y2;
                            }
                            else
                            {
                                y2 = oMeta.Page[intCurrPageNumber].ImageHeight;
                            }
                            clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);
                            if (Lines != null)
                            {
                                if (Lines.Length < lineno)
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                continue;
                            }

                            string[] SplitKeyword = keyword.Split();
                            int i;
                            int KeywordLength = keyword.Split().Length;
                            for (i = 0; i < KeywordLength; i++)
                            {
                                TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                            }
                            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);
                            if (lstEmerVendors.Contains(TopWord.strWord.ToUpper()))
                            {
                                String EmerTel = string.Empty;
                                if (TopWord.strWord.ToUpper() == "CHEMTREC")
                                {
                                    EmerTel = "18004249300";
                                }
                                else if (TopWord.strWord.ToUpper() == "INFOTRAC")
                                {
                                    EmerTel = "18005355053";
                                }
                                else if (TopWord.strWord.ToUpper() == "CHEMTEL" || TopWord.strWord.ToUpper() == "CHEM TEL")
                                {
                                    EmerTel = "18002553924";
                                }
                                // clsCnCWord CncTel = StringToCncWord(intCurrPageNumber, EmerTel);
                                clsCnCWord CncTel = ModifyWord(TopWord, EmerTel);
                                CncTel.Flag = "1";
                                CncTel.Remarks = "3F";
                                CncTel.Confidence = 100;
                                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                if (exist == false)
                                {
                                    Module1.PossibleWords.Add(CncTel);
                                }

                            }

                            else
                            {
                                string Tel;

                            Label1: if (objBoundingWord.RightWord != null)
                                {
                                    if (Module1.ExcludeWords.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                    {
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, intCurrPageNumber);
                                        goto Label1;
                                    }
                                    Tel = Module1.FuncReplace(objBoundingWord.RightWord.strWord);
                                    Tel = objBoundingWord.RightWord.strWord.Replace("#", "");
                                    if (Module1.RegTel.IsMatch(Tel))
                                    {
                                        clsCnCWord CncTel = objBoundingWord.RightWord;

                                        Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                        if (exist == false)
                                        {
                                            Module1.PossibleWords.Add(CncTel);
                                        }

                                        // PossibleWords.Add(objBoundingWord.RightWord);
                                    }
                                    else
                                    {
                                        clsCnCWord ReturnTel = FindTel(oMeta, pgno, lineno, TopKeywords, KeywordLength);
                                        if (ReturnTel.strWord != null)
                                        {
                                            Boolean exist = CheckforExistWord(Module1.PossibleWords, ReturnTel);
                                            if (exist == false)
                                            {
                                                Module1.PossibleWords.Add(ReturnTel);
                                            }
                                        }

                                    }
                                }

                                if (Module1.PossibleWords.Count == 0)
                                {
                                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
                                    if (objBoundingWord.BottomWords != null)
                                    {
                                        for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
                                        {
                                            clsCnCWord CncTel = objBoundingWord.BottomWords[t];
                                            Tel = Module1.FuncReplace(CncTel.strWord);
                                            Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");
                                            if (Module1.RegTel.IsMatch(Tel))
                                            {
                                                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
                                                if (exist == false)
                                                {
                                                    Module1.PossibleWords.Add(CncTel);
                                                    t = objBoundingWord.BottomWords.Length;
                                                }
                                            }

                                        }
                                        //clsCnCWord ReturnTel = FindTel(oMeta, pgno, lineno, TopKeywords, KeywordLength);
                                        //if (ReturnTel.strWord != null)
                                        //{
                                        //    if (ReturnTel.strWord != null)
                                        //    {
                                        //        Boolean exist = CheckforExistWord(Module1.PossibleWords, ReturnTel);
                                        //        if (exist == false)
                                        //        {
                                        //            Module1.PossibleWords.Add(ReturnTel);
                                        //        }
                                        //    }
                                        //}
                                    }
                                }
                            }
                        } // end of for 

                        if (Module1.PossibleWords.Count == 0)
                        {
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                        }

                        if (Module1.PossibleWords.Count == 1)
                        {
                            if (Module1.PossibleWords[0].Flag == null)
                            {
                                //Boolean ConfidenceFlag = HighestConfidence(Module1.PossibleWords[0]);
                                Boolean ConfidenceFlag = HighestConfidenceforHazTel(Module1.PossibleWords[0]);
                                if (ConfidenceFlag == true)
                                {
                                    Module1.PossibleWords[0].Confidence = 100;
                                    Module1.PossibleWords[0].Flag = "1";
                                    Module1.PossibleWords[0].Remarks = "1F";
                                }
                                else
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                //Module1.PossibleWords[0].Confidence = 100;

                                //Module1.PossibleWords[0].Flag = "1";
                                //Module1.PossibleWords[0].Remarks = "1F";
                            }
                            Module1.objRetStructF3.Words = Module1.PossibleWords;
                            Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            Module1.conflevel.Add(90);
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            Module1.objRetStructF3.Flag = "Y";
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Status = "S";
                            return Module1.objRetStructF3;
                        }
                        else
                        {
                            //List<double> ConfList = new List<double>();
                            clsCnCWord RetWord = new clsCnCWord();
                            foreach (clsCnCWord Word in Module1.PossibleWords)
                            {
                                // string strWord = Word.strWord;
                                string strWord = Regex.Replace(Word.strWord, @"[^0-9a-zA-Z]+", "");
                                if (VendorEmerNo.Contains(strWord))
                                {
                                    RetWord = Word;
                                    RetWord.Confidence = 100;
                                    RetWord.Flag = "1";
                                    RetWord.Remarks = "3F";
                                    break;
                                }
                            }
                            if (RetWord.strWord == null)
                            {
                                double HighestConf = Module1.PossibleWords[0].Confidence;
                                RetWord = Module1.PossibleWords[0];
                                for (int index = 0; index < Module1.PossibleWords.Count; index++)
                                {
                                    //clsCnCWord oWord = Module1.PossibleWords[index];
                                    if (Module1.PossibleWords[index].Confidence > HighestConf)
                                    {
                                        HighestConf = Module1.PossibleWords[index].Confidence;
                                        RetWord = Module1.PossibleWords[index];
                                    }
                                }
                                if (HighestConf >= 90)
                                {
                                    RetWord.Confidence = 100;
                                    RetWord.Flag = "1";
                                    RetWord.Remarks = "1F";
                                }
                                else
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                            }
                            Module1.PossibleWords.Clear();
                            Module1.PossibleWords.Add(RetWord);
                            Module1.objRetStructF3.Words = Module1.PossibleWords;
                            Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            Module1.conflevel.Add(90);
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            Module1.objRetStructF3.Flag = "Y";
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Status = "S";
                            return Module1.objRetStructF3;

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
                // ds.Clear();
                ds.Dispose();
                ds = null;
                //dt.Clear();
                dt.Dispose();
                dt = null;
                sKeyword = null;
                vendorSKeyword = null;
                foundRows = null;
                TopKeywords = null;
                objBoundingWord = null;
                VendorEmerNo = null;
            }

            return Module1.objRetStructF3;
        }
        private clsCnCWord FindTel(clsCncMetaData ObjMetaData, int intCurrPageNumber, int LineNo, clsCnCWord[] TopKeywords, int KeywordLength)
        {
            clsCnCWord cncTel = new clsCnCWord();
            string strTel = string.Empty;
            string ReplaceStrTel = Module1.FuncReplace(ObjMetaData.Page[intCurrPageNumber].Line[LineNo].strLine);
            clsCnCWord[] mergeWords = new clsCnCWord[3];

            try
            {
                if (Module1.RegTel.IsMatch(ReplaceStrTel))
                {
                    Match Tel = Module1.RegTel.Match(ReplaceStrTel);
                    strTel = Tel.ToString();
                    // ht.Add(iline, StrDate)
                }
                if (!string.IsNullOrEmpty(strTel))
                {
                    String[] splitTel = strTel.Split();
                    for (int iword = 1; iword <= ObjMetaData.Page[intCurrPageNumber].Line[LineNo].WordCount; iword++)
                    {
                        string Tel = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iword].strWord;
                        Tel = Module1.FuncReplace(Tel);
                        if (Tel == splitTel[0])
                        {
                            int counter = 0;

                            clsCnCBoundingWords objBoundingWord = Module1.oClsCnc.GetBoundingWords(ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iword], ObjMetaData, intCurrPageNumber);
                            for (int j = KeywordLength - 1; j >= 0; j--)
                            {
                            Label2: if (objBoundingWord.LeftWord != null)
                                {
                                    if (Module1.ExcludeWords.Contains(objBoundingWord.LeftWord.strWord.ToUpper()))
                                    {
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.LeftWord, ObjMetaData, intCurrPageNumber);
                                        goto Label2;
                                    }
                                    if (objBoundingWord.LeftWord.strWord.ToUpper() == TopKeywords[j].strWord.ToUpper())
                                    {
                                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.LeftWord, ObjMetaData, intCurrPageNumber);
                                        counter++;
                                        if (counter == KeywordLength)
                                        {
                                            for (int i = 0; i < splitTel.Length; i++)
                                            {
                                                mergeWords[i] = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iword + i];
                                            }
                                            //clsCnCWord[] mergeWords = { ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword], ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword + 1], ObjMetaData.Page[intCurrPageNumber].Line[line].Word[iword + 2] };
                                            cncTel = Module1.oClsCnc.MergeWords(mergeWords);
                                            j = 0;
                                            iword = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].WordCount;
                                        }

                                    }

                                }
                                else
                                {
                                    break;
                                }
                            }

                            if (cncTel.strWord == null)
                            {
                                for (int i = 0; i < splitTel.Length; i++)
                                {
                                    mergeWords[i] = ObjMetaData.Page[intCurrPageNumber].Line[LineNo].Word[iword + i];
                                }
                                cncTel = Module1.oClsCnc.MergeWords(mergeWords);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            { }
            finally
            {
                mergeWords = null;
            }
            return cncTel;
        }
        #endregion

        #region EmergencyTelephone Code
        private RetStructF3 ERTCode(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.PossibleWords.Clear();
            Module1.conflevel.Clear();

            try
            {
                string HazMatCode = string.Empty;
                clsCnCWord ObjCode;
                //string TelHazmat = iQPUBLIC.Common.GetHeaderDataForCell(7, iQPUBLIC.Common.GetDataTable());
                string TelHazmat = Module1.EMerTel;
                if (!string.IsNullOrEmpty(TelHazmat))
                {
                    if (TelHazmat != "*****")
                    {
                        if (TelHazmat.Contains("8004249300"))
                        {
                            HazMatCode = "CHEMTREC#CTREC";
                        }
                        else if (TelHazmat.Contains("8005355053"))
                        {
                            HazMatCode = "INFOTRAC#INFO";
                        }
                        else if (TelHazmat.Contains("8002553924"))
                        {
                            HazMatCode = "CHEMTEL#CHEMT";
                        }
                        else if (TelHazmat.Contains("8005804632"))
                        {
                            HazMatCode = "SAIA#SAIA TO SAIA";
                        }
                        else
                        {
                            HazMatCode = "EMERGENCY#EMER";
                        }

                        ObjCode = StringToCncWord(intCurrPageNumber, HazMatCode);
                        ObjCode.Confidence = 100;
                        ObjCode.Flag = "1";
                        ObjCode.Remarks = "3F";
                        Module1.PossibleWords.Add(ObjCode);
                    }

                }

                if (Module1.PossibleWords.Count == 0)
                {
                    Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                }

                if (Module1.PossibleWords.Count == 1)
                {
                    Module1.objRetStructF3.Words = Module1.PossibleWords;
                    Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                    Module1.conflevel.Add(90);
                    Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                    Module1.objRetStructF3.Flag = "Y";
                    Module1.objRetStructF3.ManualConfirmation = "N";
                    Module1.objRetStructF3.Status = "S";
                    // Module1.EMerTel = "";
                    Module1.EMerTel = null;
                    // return Module1.objRetStructF3;
                }

                TelHazmat = null;
                ObjCode = null;
            }
            catch (Exception)
            {
                // throw;
            }

            return Module1.objRetStructF3;
        }
        #endregion

        #region Load#
        private RetStructF3 LoadId(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            DataSet ds = Module1.MakeDs(oMeta, intCurrPageNumber);
            DataTable dt = new DataTable();

            Module1.PossibleWords.Clear();
            Module1.conflevel.Clear();
            int iRecursiveCallCount = 0;
            string sKeyword = string.Empty;

            System.Data.DataRow[] foundRows;
            clsCnCWord[] TopKeywords = new clsCnCWord[10];
            int flag = 0;
            clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();

            try
            {
            Line1:
                if (iRecursiveCallCount == 0)
                {
                    // -------------------------------------------------BILL OF LADING NUMBER KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------               
                    sKeyword = "'Load','Load ID','Load Number','Load No','LOAD#'";
                }
                //else
                //{
                //    sKeyword = "'Bill of Lading'";
                //}

                //try
                //{
                Module1.objRetStructF3.Status = "F";
                Module1.objRetStructF3.ManualConfirmation = "Y";
                Module1.objRetStructF3.Flag = "Y";
                if (iQPUBLIC.PublicComponents.htMyVariable.Contains("BOLSet1_keyword"))
                {
                    Module1.dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["BOLSet1_keyword"];
                }
                if (Module1.dsKeys != null)
                {
                    dt = Module1.dsKeys.Tables[0];

                    //foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
                    // foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber);
                    foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=" + intCurrPageNumber + "AND [X1]>=" + Module1.MX1 + " AND [Y1]>=" + Module1.MY1 + " AND [X2]<=" + Module1.MX2 + " AND [Y2]<=" + Module1.MY2);
                    var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
                    foundRows = orderedRows.ToArray();
                    string keyword = string.Empty;
                    int pgno = 0;
                    int lineno = 0;
                    int wordno = 0;
                    int LineWordNo = 0;
                    int x1 = 0;
                    int y1 = 0;
                    int x2 = 0;
                    int y2 = 0;
                    //int h = 0;
                    int w = 0;

                    string[] FilterRightWord = { ":", ".", "#", "#:", "|" };

                    //ClsCNC oClsCnc = new ClsCNC();
                    //clsCnCWord Words = new clsCnCWord();

                    //clsCnCWord sJoinWord = new clsCnCWord();
                    //regex = new Regex("^([a-zA-Z]*[-/]?[a-zA-Z]*[0-9]+[-/]?[a-zA-Z]*[-/]?)*$");
                    //if (foundRows.Length == 0 && iRecursiveCallCount == 0)
                    //{
                    //    iRecursiveCallCount = iRecursiveCallCount + 1;
                    //    if (iRecursiveCallCount == 1)
                    //    {
                    //        goto Line1;
                    //    }
                    //}
                    /*else*/ if (foundRows.Length == 1)
                    {
                        Array.Clear(TopKeywords, 0, TopKeywords.Length);
                        keyword = (foundRows[0].ItemArray[0]).ToString();
                        pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
                        lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
                        wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
                        LineWordNo = Convert.ToInt32(foundRows[0].ItemArray[4]);
                        x1 = Convert.ToInt32(foundRows[0].ItemArray[5]);
                        y1 = Convert.ToInt32(foundRows[0].ItemArray[6]);
                        x2 = Convert.ToInt32(foundRows[0].ItemArray[7]);
                        y2 = Convert.ToInt32(foundRows[0].ItemArray[8]);
                        w = oMeta.Page[intCurrPageNumber].ImageWidth;


                        if ((Convert.ToInt32(oMeta.Page[intCurrPageNumber].ImageHeight * 0.5)) < 1000)
                        {
                            y2 = y2;
                        }
                        else
                        {
                            y2 = Convert.ToInt32(oMeta.Page[intCurrPageNumber].ImageHeight * 0.5);
                        }

                        clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);


                        if (Lines != null)
                        {
                            if (Lines.Length <= lineno)
                            {
                                //Module1.objRetStructF3.Status = "F";
                                Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                                Module1.objRetStructF3.Words = Module1.PossibleWords;
                                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                                Module1.conflevel.Add(90);
                                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                                Module1.objRetStructF3.Status = "S";
                                //Module1.objRetStructF3.ManualConfirmation = "Y";
                                Module1.objRetStructF3.ManualConfirmation = "N";
                                Module1.objRetStructF3.Flag = "Y";
                                return Module1.objRetStructF3;
                            }
                        }
                        else
                        {
                            //Module1.objRetStructF3.Status = "F";
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                            Module1.objRetStructF3.Words = Module1.PossibleWords;
                            Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                            Module1.conflevel.Add(90);
                            Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                            Module1.objRetStructF3.Status = "S";
                            Module1.objRetStructF3.ManualConfirmation = "N";
                            Module1.objRetStructF3.Flag = "Y";
                            return Module1.objRetStructF3;
                        }

                        //clsCnCWord[] TopKeywords = new clsCnCWord[50];
                        //string[] SplitKeyword = keyword.Split();
                        int i;
                        int KeywordLength = keyword.Split().Length;
                        for (i = 0; i < KeywordLength; i++)
                        {
                            TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                        }

                        clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);

                        flag = 0;

                    Label1: if (objBoundingWord.RightWord != null)
                        {
                            if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                            {
                                if (objBoundingWord.RightWord.strWord == "#" || objBoundingWord.RightWord.strWord == "#:" || objBoundingWord.RightWord.strWord == "#.")
                                {
                                    flag = 1;
                                }
                                objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, intCurrPageNumber);
                                goto Label1;
                            }

                            CheckRightWordForLoad(oMeta, objBoundingWord, intCurrPageNumber, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 1, iRecursiveCallCount);
                        }


                        if (Module1.PossibleWords.Count == 0)
                        {
                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                        }

                        Module1.objRetStructF3 = GetValues(oMeta,pgno, 4, iRecursiveCallCount,TopKeywords);
                        return Module1.objRetStructF3;
                    }
                    else
                    {
                        foreach (DataRow drDataRow in foundRows)
                        {
                            Array.Clear(TopKeywords, 0, TopKeywords.Length);
                            keyword = (drDataRow.ItemArray[0]).ToString();
                            pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
                            lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
                            wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
                            LineWordNo = Convert.ToInt32(drDataRow.ItemArray[4]);
                            x1 = Convert.ToInt32(drDataRow.ItemArray[5]);
                            y1 = Convert.ToInt32(drDataRow.ItemArray[6]);
                            x2 = Convert.ToInt32(drDataRow.ItemArray[7]);
                            y2 = Convert.ToInt32(drDataRow.ItemArray[8]);

                            w = oMeta.Page[intCurrPageNumber].ImageWidth;

                            if ((Convert.ToInt32(oMeta.Page[intCurrPageNumber].ImageHeight * 0.5)) < 1000)
                            {
                                y2 = y2;
                            }
                            else
                            {
                                y2 = Convert.ToInt32(oMeta.Page[intCurrPageNumber].ImageHeight * 0.5);
                            }

                            clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);
                            if (Lines != null)
                            {
                                if (Lines.Length <= lineno)
                                {
                                    continue;
                                }
                            }
                            else
                            {
                                continue;
                            }

                            //clsCnCWord[] TopKeywords = new clsCnCWord[50];
                            //string[] SplitKeyword = keyword.Split();
                            int i;

                            int KeywordLength = keyword.Split().Length;
                            for (i = 0; i < KeywordLength; i++)
                            {
                                TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
                            }

                            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);

                            flag = 0;
                        Label1: if (objBoundingWord.RightWord != null)
                            {
                                if (FilterRightWord.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
                                {
                                    if (objBoundingWord.RightWord.strWord == "#" || objBoundingWord.RightWord.strWord == "#:" || objBoundingWord.RightWord.strWord == "#.")
                                    {
                                        flag = 1;
                                    }
                                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, intCurrPageNumber);
                                    goto Label1;
                                }
                                CheckRightWordForLoad(oMeta, objBoundingWord, pgno, Module1.PossibleWords, TopKeywords, lineno, LineWordNo, 1, iRecursiveCallCount);
                            }

                        } // end of for 

                        if (Module1.PossibleWords.Count == 0)
                        {
                            //iRecursiveCallCount = iRecursiveCallCount + 1;
                            //if (iRecursiveCallCount == 1)
                            //{
                            //    goto Line1;
                            //}

                            Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
                        }

                        Module1.objRetStructF3 = GetValues(oMeta, pgno, 4, iRecursiveCallCount,TopKeywords);
                        return Module1.objRetStructF3;

                    }                    // end of the F3
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
                //dt.Clear();
                dt.Dispose();
                dt = null;
                sKeyword = null;
                foundRows = null;
                TopKeywords = null;
                objBoundingWord = null;
            }
            return Module1.objRetStructF3;
        }

        private void CheckRightWordForLoad(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord, int field, int iRecursiveCallCount)
        {
            //ClsCNC objCNC = new ClsCNC();
            //clsCNCSBR ObjSbr = new clsCNCSBR();
            //clsCnCWord[] PoNumber = new clsCnCWord[5];
            //   clsCnCWord PoNoMerged = new clsCnCWord();
            List<string> Separators = new List<string>() { ",", "&", "AND", "/" };
            char ch = ',';
            string WordToChckLength = string.Empty;
            if (ObjBoundingWord.RightWord != null)
            {
                string word = ObjBoundingWord.RightWord.strWord;
                if (word.Length > 2)
                {
                    WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                    // word = word.Replace("'", "");
                    if (Module1.SpecialCharacters.Contains(word[0].ToString()) && (!word.StartsWith("(")))
                    {
                        word = word.Remove(0, 1);
                    }
                    if (Module1.SpecialCharacters.Any(word.EndsWith) && (!word.EndsWith(")")))
                    {
                        word = word.Remove(word.Length - 1, 1);
                    }
                }
                RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");
                if (Obj008.Status == "S")
                {
                    if (word.All(char.IsLetterOrDigit))
                    {
                        Obj008.Status = "F";
                    }
                }
                if (!string.IsNullOrEmpty(word))
                {
                    if ((!string.IsNullOrEmpty(word)) && (!(Module1.NewRegexForAlphanumeric.IsMatch(word))) && (!WordToChckLength.All(char.IsLetter)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word != Module1.ProNo&&word.Length>=4)
                    {
                        if (ObjBoundingWord.RightWord.strWord.All(char.IsDigit))
                        {
                            clsCnCWord CncPoWord = ObjBoundingWord.RightWord;
                            clsCnCBoundingWords Newbounding = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                            int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                            if (word.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                            {
                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, ",", iRecursiveCallCount,field);
                                Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                if (exist == false)
                                {
                                    PWords.Add(CncPoWord);
                                }
                            }
                            else
                            {
                                //CncPoWord = FindNextWordforNumeric(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord,keyword);
                                string strCncBLWord = CncPoWord.strWord.Replace(" ", "");
                                if (strCncBLWord.Length >= 4)
                                {
                                    if (Module1.SpecialCharacters.Contains(CncPoWord.strWord[0].ToString()) && (!CncPoWord.strWord.StartsWith("(")))
                                    {
                                        CncPoWord = ModifyWord(CncPoWord, CncPoWord.strWord.Remove(0, 1));
                                    }
                                    Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                    if (exist == false)
                                    {
                                        PWords.Add(CncPoWord);
                                    }
                                }

                            }
                        }
                        else
                        {
                            //if (word.Count(char.IsLetter) <= word.Count(char.IsDigit))
                            //{
                            if (word.Length >= 4)
                            {
                                clsCnCWord CncPoWord = ObjBoundingWord.RightWord;
                                string strCncBLWord = CncPoWord.strWord;
                                int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                if (CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                {
                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, iLine, CncPoWord, ",", iRecursiveCallCount,field);
                                }
                                if (Module1.SpecialCharacters.Contains(CncPoWord.strWord[0].ToString()) && (!CncPoWord.strWord.StartsWith("(")))
                                {
                                    CncPoWord = ModifyWord(CncPoWord, CncPoWord.strWord.Remove(0, 1));
                                }
                                //if (!CncPoWord.strWord.Contains("~")&& (!Module1.SpecialCharacters.Any(CncPoWord.strWord.Contains)))
                                //{
                                //    if ((CncPoWord.strWord.Count(char.IsLetter) == 1) && (!char.IsLetter(CncPoWord.strWord[0])) && (!char.IsLetter(word[CncPoWord.strWord.Length - 1])))
                                //    {
                                //        CncPoWord = Module1.ReplaceAlphanumericWord(CncPoWord);
                                //    }
                                //}
                                Boolean exist = CheckforExistWord(PWords, CncPoWord);
                                if (exist == false)
                                {
                                    PWords.Add(CncPoWord);
                                }
                            }

                            //}
                        }
                    }
                }

            }

        }

        #endregion

        //#region IndividualTelephone
        //private RetStructF3 Tel_Indiv(clsCncMetaData oMeta, int intCurrPageNumber, string strArg)
        //{
        //    DataSet ds = Module1.MakeDs(oMeta, intCurrPageNumber);

        //    //List<clsCnCWord> PossibleWords = new List<clsCnCWord>();

        //    DataTable dt = new DataTable();
        //    //List<clsCnCWord> PossibleWords1 = new List<clsCnCWord>();
        //    //List<int> conflevel = new List<int>();
        //    Module1.PossibleWords.Clear();
        //    Module1.conflevel.Clear();
        //    int iRecursiveCallCount = 0;
        //    string sKeyword = string.Empty;
        //    string[] ExcludeWords = { ":", ".", "|", "#", "#:", "PHONE", "PHONE#", "PHONE:", "PHONE#:" };
        //    //  Line1:

        //    if (iRecursiveCallCount == 0)
        //    {
        //        // -------------------------------------------------INDIVIDUAL TELEPHONE NUMBER KEYWORD-----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------                
        //        sKeyword = "'Consignee','Ship To','Destination / Consignee','Deliver to','Destination/Consignee','TO','TO(CONSIGNEE):','CONSIGNEE PHONE NO','Destination','Customer Phone #:'";
        //    }
        //    // sKeyword = "'Invoice No:'"
        //    // -----------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
        //    //else
        //    //{
        //    //    //sKeyword = "'Bill of Lading', 'BI LL 0 F LADING'";
        //    //    //sKeyword = "'Account Number','Invoice No','nvoice no','Document No  :''Account Number','INVOICE','FATTURA','Numero','invoice No.','Cr.Note','TAX NVOICE NO.','TAX INVOICE','Jmoice','Inv.No.  :', 'Inv.No.:'";
        //    //    // sKeyword = "'Invoice No:'"
        //    //}

        //    try
        //    {
        //        Module1.objRetStructF3.Status = "F";
        //        Module1.objRetStructF3.ManualConfirmation = "Y";
        //        Module1.objRetStructF3.Flag = "Y";
        //        if (iQPUBLIC.PublicComponents.htMyVariable.Contains("keyword"))
        //        {
        //            Module1.dsKeys = (DataSet)iQPUBLIC.PublicComponents.htMyVariable["keyword"];
        //        }
        //        if (Module1.dsKeys != null)
        //        {
        //            dt = Module1.dsKeys.Tables[0];
        //            System.Data.DataRow[] foundRows;
        //            foundRows = dt.Select("Keyword IN " + "(" + sKeyword + ") AND [Page No]=1");
        //            var orderedRows = foundRows.OrderByDescending(item => item.ItemArray[0]);
        //            foundRows = orderedRows.ToArray();
        //            string keyword = string.Empty;
        //            int pgno = 0;
        //            int lineno = 0;
        //            int wordno = 0;
        //            int LineWordNo = 0;
        //            int x1 = 0;
        //            int y1 = 0;
        //            int x2 = 0;
        //            int y2 = 0;
        //            //int h = 0;
        //            int w = 0;

        //            //ClsCNC oClsCnc = new ClsCNC();
        //            //clsCnCWord Words = new clsCnCWord();
        //            clsCnCBoundingWords objBoundingWord = new clsCnCBoundingWords();
        //            //clsCnCWord sJoinWord = new clsCnCWord();
        //            //regex = new Regex("^([a-zA-Z]*[-/]?[a-zA-Z]*[0-9]+[-/]?[a-zA-Z]*[-/]?)*$");
        //            if (foundRows.Length == 1)
        //            {
        //                keyword = (foundRows[0].ItemArray[0]).ToString();
        //                pgno = Convert.ToInt32(foundRows[0].ItemArray[1]);
        //                lineno = Convert.ToInt32(foundRows[0].ItemArray[2]);
        //                wordno = Convert.ToInt32(foundRows[0].ItemArray[3]);
        //                LineWordNo = Convert.ToInt32(foundRows[0].ItemArray[4]);
        //                x1 = Convert.ToInt32(foundRows[0].ItemArray[5]);
        //                y1 = Convert.ToInt32(foundRows[0].ItemArray[6]);
        //                x2 = Convert.ToInt32(foundRows[0].ItemArray[7]);
        //                y2 = Convert.ToInt32(foundRows[0].ItemArray[8]);
        //                w = oMeta.Page[intCurrPageNumber].ImageWidth;


        //                if (oMeta.Page[intCurrPageNumber].ImageHeight < 1000)
        //                {
        //                    y2 = y2;
        //                }
        //                else
        //                {
        //                    y2 = oMeta.Page[intCurrPageNumber].ImageHeight / 2;
        //                }

        //                clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);

        //                if (Lines != null)
        //                {
        //                    if (Lines.Length <= lineno)
        //                    {
        //                        //Module1.objRetStructF3.Status = "F";
        //                        Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
        //                        Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                        Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                        Module1.conflevel.Add(90);
        //                        Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                        Module1.objRetStructF3.Status = "S";
        //                        //Module1.objRetStructF3.ManualConfirmation = "Y";
        //                        Module1.objRetStructF3.ManualConfirmation = "N";
        //                        Module1.objRetStructF3.Flag = "Y";
        //                        return Module1.objRetStructF3;
        //                    }
        //                }
        //                else
        //                {
        //                    //Module1.objRetStructF3.Status = "F";
        //                    Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
        //                    Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                    Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                    Module1.conflevel.Add(90);
        //                    Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                    Module1.objRetStructF3.Status = "S";
        //                    //Module1.objRetStructF3.ManualConfirmation = "Y";
        //                    Module1.objRetStructF3.ManualConfirmation = "N";
        //                    Module1.objRetStructF3.Flag = "Y";
        //                    return Module1.objRetStructF3;
        //                }

        //                clsCnCWord[] TopKeywords = new clsCnCWord[50];
        //                string[] SplitKeyword = keyword.Split();
        //                int i;
        //                int KeywordLength = keyword.Split().Length;
        //                for (i = 0; i < KeywordLength; i++)
        //                {
        //                    TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
        //                }

        //                clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
        //                objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);

        //                string Tel;
        //                Label1: if (objBoundingWord.RightWord != null)
        //                {
        //                    if (ExcludeWords.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
        //                    {
        //                        objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, intCurrPageNumber);
        //                        goto Label1;
        //                    }
        //                    Tel = objBoundingWord.RightWord.strWord.Replace(",", "");
        //                    Tel = objBoundingWord.RightWord.strWord.Replace("#", "");
        //                    if (Module1.RegTelwithMob.IsMatch(Tel))
        //                    {
        //                        clsCnCWord CncTel = objBoundingWord.RightWord;
        //                        Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
        //                        if (exist == false)
        //                        {
        //                            Module1.PossibleWords.Add(objBoundingWord.RightWord);
        //                        }
        //                    }

        //                    //else
        //                    //{
        //                    //    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
        //                    //    if (objBoundingWord.BottomWords != null)
        //                    //    {
        //                    //        for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
        //                    //        {
        //                    //            clsCnCWord CncTel = objBoundingWord.BottomWords[t];
        //                    //            Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
        //                    //            Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

        //                    //            if (Module1.RegTel.IsMatch(Tel))
        //                    //            {
        //                    //                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
        //                    //                if (exist == false)
        //                    //                {
        //                    //                    Module1.PossibleWords.Add(CncTel);
        //                    //                    t = objBoundingWord.BottomWords.Length;
        //                    //                }
        //                    //            }

        //                    //        }
        //                    //    }
        //                    //}
        //                }
        //                //else
        //                //{
        //                //    //clsCnCWord BOL = new clsCnCWord();
        //                //    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
        //                //    if (objBoundingWord.BottomWords != null)
        //                //    {
        //                //        for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
        //                //        {
        //                //            clsCnCWord CncTel = objBoundingWord.BottomWords[t];
        //                //            Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
        //                //            Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

        //                //            if (Module1.RegTel.IsMatch(Tel))
        //                //            {
        //                //                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
        //                //                if (exist == false)
        //                //                {
        //                //                    Module1.PossibleWords.Add(CncTel);
        //                //                    t = objBoundingWord.BottomWords.Length;
        //                //                }
        //                //            }

        //                //        }
        //                //    }
        //                //}

        //                //if (TopWord.strWord.ToUpper() == "TO")
        //                //{
        //                //    if (objBoundingWord.RightWord != null)
        //                //    {
        //                //        if (objBoundingWord.RightWord.strWord.ToUpper() == "PHONE#:")
        //                //        {
        //                //            objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
        //                //            if (objBoundingWord.RightWord != null)
        //                //            {
        //                //                if (Module1.RegTel.IsMatch(objBoundingWord.RightWord.strWord))
        //                //                {
        //                //                    Module1.PossibleWords.Add(objBoundingWord.RightWord);
        //                //                }
        //                //            }
        //                //        }
        //                //    }
        //                //}

        //                if (Module1.PossibleWords.Count == 0)
        //                {
        //                    if (keyword.ToUpper() != "TO")
        //                    {
        //                        string strTel;

        //                        int KeywordLeft = FindX1(TopWord);

        //                        for (i = 1; i <= 8; i++)
        //                        {
        //                            if (i != 2)
        //                            {
        //                                if (lineno + i <= oMeta.Page[intCurrPageNumber].LineCount)
        //                                {
        //                                    clsCnCLine NewLine = oMeta.Page[intCurrPageNumber].Line[lineno + i];
        //                                    if (Module1.RegTelwithMob.IsMatch(NewLine.strLine))
        //                                    {
        //                                        foreach (Match match in Module1.RegTelwithMob.Matches(NewLine.strLine))
        //                                        {
        //                                            clsCnCWord[] mergeWords = new clsCnCWord[5];
        //                                            //Match tel = match.Value;
        //                                            strTel = match.Value;
        //                                            String[] splitTel = strTel.Split();
        //                                            for (int iWord = 1; iWord <= oMeta.Page[pgno].Line[lineno + i].WordCount; iWord++)
        //                                            {
        //                                                if (oMeta.Page[pgno].Line[lineno + i].Word[iWord].strWord.ToUpper() == splitTel[0].ToUpper())
        //                                                {
        //                                                    int TelLeft = FindX1(oMeta.Page[pgno].Line[lineno + i].Word[iWord]);

        //                                                    if (TelLeft >= KeywordLeft - 300 && TelLeft <= KeywordLeft + 300)
        //                                                    {
        //                                                        for (int j = 0; j < splitTel.Length; j++)
        //                                                        {
        //                                                            mergeWords[j] = oMeta.Page[pgno].Line[lineno + i].Word[iWord + j];
        //                                                        }
        //                                                        clsCnCWord cncTel = Module1.oClsCnc.MergeWords(mergeWords);
        //                                                        Boolean exist = CheckforExistWord(Module1.PossibleWords, cncTel);
        //                                                        if (exist == false)
        //                                                        {
        //                                                            Module1.PossibleWords.Add(cncTel);
        //                                                            i = 10;
        //                                                        }

        //                                                    }
        //                                                    iWord = oMeta.Page[pgno].Line[lineno + i].WordCount;
        //                                                }
        //                                            }
        //                                        }
        //                                    }


        //                                }
        //                            }
        //                        }

        //                    }
        //                }
        //                if (Module1.PossibleWords.Count == 0)
        //                {
        //                    //iRecursiveCallCount = iRecursiveCallCount + 1;
        //                    //if (iRecursiveCallCount == 1)
        //                    //{
        //                    //    goto Line1;
        //                    //}

        //                    Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
        //                    //Module1.objRetStructF3.Words = PossibleWords;
        //                    //Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
        //                    //conflevel.Add(90);
        //                    //Module1.objRetStructF3.ConfidenceLevelofSuspect = conflevel;
        //                    //Module1.objRetStructF3.Status = "S";
        //                    //Module1.objRetStructF3.ManualConfirmation = "N";
        //                    //Module1.objRetStructF3.Flag = "Y";
        //                }

        //                if (Module1.PossibleWords.Count == 1)
        //                {
        //                    Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                    Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                    Module1.conflevel.Add(90);
        //                    Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                    Module1.objRetStructF3.Flag = "Y";
        //                    if (iRecursiveCallCount == 1)
        //                    {
        //                        Module1.objRetStructF3.ManualConfirmation = "Y";
        //                    }
        //                    else
        //                    {
        //                        Module1.objRetStructF3.ManualConfirmation = "N";
        //                    }

        //                    // returnZones.ManualConfirmation = "N"
        //                    Module1.objRetStructF3.Status = "S";
        //                    return Module1.objRetStructF3;
        //                }
        //                if (Module1.PossibleWords.Count == 2)
        //                {
        //                    Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                    Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                    Module1.conflevel.Add(80);
        //                    Module1.conflevel.Add(80);
        //                    Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                    Module1.objRetStructF3.Flag = "Y";
        //                    Module1.objRetStructF3.ManualConfirmation = "Y";
        //                    Module1.objRetStructF3.Status = "S";
        //                    return Module1.objRetStructF3;
        //                }

        //            }
        //            else
        //            {
        //                foreach (DataRow drDataRow in foundRows)
        //                {
        //                    keyword = (drDataRow.ItemArray[0]).ToString();
        //                    pgno = Convert.ToInt32(drDataRow.ItemArray[1]);
        //                    lineno = Convert.ToInt32(drDataRow.ItemArray[2]);
        //                    wordno = Convert.ToInt32(drDataRow.ItemArray[3]);
        //                    LineWordNo = Convert.ToInt32(drDataRow.ItemArray[4]);
        //                    x1 = Convert.ToInt32(drDataRow.ItemArray[5]);
        //                    y1 = Convert.ToInt32(drDataRow.ItemArray[6]);
        //                    x2 = Convert.ToInt32(drDataRow.ItemArray[7]);
        //                    y2 = Convert.ToInt32(drDataRow.ItemArray[8]);

        //                    w = oMeta.Page[intCurrPageNumber].ImageWidth;

        //                    if (oMeta.Page[intCurrPageNumber].ImageHeight < 1000)
        //                    {
        //                        y2 = y2;
        //                    }
        //                    else
        //                    {
        //                        y2 = oMeta.Page[intCurrPageNumber].ImageHeight / 2;
        //                    }
        //                    clsCnCLine[] Lines = Module1.oClsCnc.GetLinesFromROI(oMeta.Page, pgno, 0, 0, w, y2);
        //                    if (Lines != null)
        //                    {
        //                        if (Lines.Length <= lineno)
        //                        {
        //                            continue;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        continue;
        //                    }

        //                    clsCnCWord[] TopKeywords = new clsCnCWord[50];
        //                    string[] SplitKeyword = keyword.Split();
        //                    int i;
        //                    int KeywordLength = keyword.Split().Length;
        //                    for (i = 0; i < KeywordLength; i++)
        //                    {
        //                        TopKeywords[i] = oMeta.Page[pgno].Line[lineno].Word[LineWordNo + i];
        //                    }

        //                    string strTel;
        //                    clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
        //                    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[KeywordLength - 1], oMeta, pgno);


        //                    //if (keyword.ToUpper() == "TO")
        //                    //{
        //                    //    if (objBoundingWord.RightWord != null)
        //                    //    {

        //                    //        if (objBoundingWord.RightWord.strWord.ToUpper() == "PHONE#" || objBoundingWord.RightWord.strWord.ToUpper() == "PHONE#:")
        //                    //        {
        //                    //            objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, pgno);
        //                    //            if (objBoundingWord.RightWord != null)
        //                    //            {
        //                    //                if (Module1.RegTel.IsMatch(objBoundingWord.RightWord.strWord))
        //                    //                {
        //                    //                    Module1.PossibleWords.Add(objBoundingWord.RightWord);
        //                    //                }
        //                    //            }
        //                    //        }
        //                    //    }
        //                    //}

        //                    string Tel, PossibleTel = string.Empty;
        //                    Label1: if (objBoundingWord.RightWord != null)
        //                    {
        //                        if (ExcludeWords.Contains(objBoundingWord.RightWord.strWord.ToUpper()))
        //                        {
        //                            objBoundingWord = Module1.oClsCnc.GetBoundingWords(objBoundingWord.RightWord, oMeta, intCurrPageNumber);
        //                            goto Label1;
        //                        }
        //                        Tel = objBoundingWord.RightWord.strWord.Replace(",", "");
        //                        Tel = objBoundingWord.RightWord.strWord.Replace("#", "");
        //                        if (Module1.RegTelwithMob.IsMatch(Tel))
        //                        {
        //                            clsCnCWord CncTel = objBoundingWord.RightWord;
        //                            Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
        //                            if (exist == false)
        //                            {
        //                                Module1.PossibleWords.Add(objBoundingWord.RightWord);
        //                                PossibleTel = objBoundingWord.RightWord.strWord;
        //                            }
        //                        }

        //                        //else
        //                        //{
        //                        //    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
        //                        //    if (objBoundingWord.BottomWords != null)
        //                        //    {
        //                        //        for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
        //                        //        {
        //                        //            clsCnCWord CncTel = objBoundingWord.BottomWords[t];
        //                        //            Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
        //                        //            Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

        //                        //            if (Module1.RegTel.IsMatch(Tel))
        //                        //            {
        //                        //                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
        //                        //                if (exist == false)
        //                        //                {
        //                        //                    Module1.PossibleWords.Add(CncTel);
        //                        //                    t = objBoundingWord.BottomWords.Length;
        //                        //                }
        //                        //            }

        //                        //        }
        //                        //    }
        //                        //}
        //                    }
        //                    //else
        //                    //{
        //                    //    //clsCnCWord BOL = new clsCnCWord();
        //                    //    objBoundingWord = Module1.oClsCnc.GetBoundingWords(TopWord, oMeta, pgno);
        //                    //    if (objBoundingWord.BottomWords != null)
        //                    //    {
        //                    //        for (int t = 0; t < objBoundingWord.BottomWords.Length; t++)
        //                    //        {
        //                    //            clsCnCWord CncTel = objBoundingWord.BottomWords[t];
        //                    //            Tel = objBoundingWord.BottomWords[t].strWord.Replace(",", "");
        //                    //            Tel = objBoundingWord.BottomWords[t].strWord.Replace("#", "");

        //                    //            if (Module1.RegTel.IsMatch(Tel))
        //                    //            {
        //                    //                Boolean exist = CheckforExistWord(Module1.PossibleWords, CncTel);
        //                    //                if (exist == false)
        //                    //                {
        //                    //                    Module1.PossibleWords.Add(CncTel);
        //                    //                    t = objBoundingWord.BottomWords.Length;
        //                    //                }
        //                    //            }

        //                    //        }
        //                    //    }
        //                    //}

        //                    //if (Module1.PossibleWords.Count==0)
        //                    if (string.IsNullOrEmpty(PossibleTel))
        //                    {
        //                        if (keyword.ToUpper() != "TO")
        //                        {
        //                            int KeywordLeft = FindX1(TopWord);

        //                            for (i = 1; i <= 8; i++)
        //                            {
        //                                if (i != 2)
        //                                {
        //                                    if (lineno + i <= oMeta.Page[intCurrPageNumber].LineCount)
        //                                    {
        //                                        clsCnCLine NewLine = oMeta.Page[intCurrPageNumber].Line[lineno + i];
        //                                        if (Module1.RegTelwithMob.IsMatch(NewLine.strLine))
        //                                        {
        //                                            foreach (Match match in Module1.RegTelwithMob.Matches(NewLine.strLine))
        //                                            {
        //                                                clsCnCWord[] mergeWords = new clsCnCWord[5];
        //                                                //Match tel = match.Value;
        //                                                strTel = match.Value;
        //                                                String[] splitTel = strTel.Split();
        //                                                for (int iWord = 1; iWord <= oMeta.Page[pgno].Line[lineno + i].WordCount; iWord++)
        //                                                {
        //                                                    if (oMeta.Page[pgno].Line[lineno + i].Word[iWord].strWord.ToUpper() == splitTel[0].ToUpper())
        //                                                    {
        //                                                        int TelLeft = FindX1(oMeta.Page[pgno].Line[lineno + i].Word[iWord]);

        //                                                        if (TelLeft >= KeywordLeft - 300 && TelLeft <= KeywordLeft + 300)
        //                                                        {
        //                                                            for (int j = 0; j < splitTel.Length; j++)
        //                                                            {
        //                                                                mergeWords[j] = oMeta.Page[pgno].Line[lineno + i].Word[iWord + j];
        //                                                            }
        //                                                            clsCnCWord cncTel = Module1.oClsCnc.MergeWords(mergeWords);
        //                                                            Boolean exist = CheckforExistWord(Module1.PossibleWords, cncTel);
        //                                                            if (exist == false)
        //                                                            {
        //                                                                Module1.PossibleWords.Add(cncTel);
        //                                                                i = 8;
        //                                                            }

        //                                                        }
        //                                                        iWord = oMeta.Page[pgno].Line[lineno + i].WordCount;
        //                                                    }
        //                                                }
        //                                            }
        //                                        }


        //                                    }

        //                                }
        //                            }

        //                        }
        //                    }

        //                } // end of for 

        //                if (Module1.PossibleWords.Count == 0)
        //                {
        //                    ////iRecursiveCallCount = iRecursiveCallCount + 1;
        //                    ////if (iRecursiveCallCount == 1)
        //                    ////{
        //                    ////    goto Line1;
        //                    ////}

        //                    Module1.PossibleWords.Add(StringToCncWord(1, "*****"));
        //                    //Module1.objRetStructF3.Words = PossibleWords;
        //                    //Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
        //                    //conflevel.Add(90);
        //                    //Module1.objRetStructF3.ConfidenceLevelofSuspect = conflevel;
        //                    //Module1.objRetStructF3.Status = "S";
        //                    //Module1.objRetStructF3.ManualConfirmation = "N";
        //                    //Module1.objRetStructF3.Flag = "Y";
        //                }

        //                if (Module1.PossibleWords.Count > 3)
        //                {
        //                    Module1.PossibleWords = GetTopThreeValue(Module1.PossibleWords);
        //                }

        //                if (Module1.PossibleWords.Count == 1)
        //                {
        //                    Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                    Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                    Module1.conflevel.Add(90);
        //                    Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                    Module1.objRetStructF3.Flag = "Y";
        //                    if (iRecursiveCallCount == 1)
        //                    {
        //                        Module1.objRetStructF3.ManualConfirmation = "Y";
        //                    }
        //                    else
        //                    {
        //                        Module1.objRetStructF3.ManualConfirmation = "N";
        //                    }

        //                    // returnZones.ManualConfirmation = "N"
        //                    Module1.objRetStructF3.Status = "S";
        //                    return Module1.objRetStructF3;
        //                }
        //                else
        //                {
        //                    //returnZones = GetTotalInvoices_InvoiceNumber(oMeta, intCurrPageNumber, PossibleWords, sKeyword);
        //                    Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                    Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                    switch (Module1.objRetStructF3.NoOfieldsSuspects)
        //                    {
        //                        case 0:
        //                            {
        //                                Module1.objRetStructF3.Flag = "N";
        //                                Module1.objRetStructF3.ManualConfirmation = "N";
        //                                Module1.objRetStructF3.Status = "S";
        //                                return Module1.objRetStructF3;
        //                            }

        //                        case 2:
        //                            {
        //                                Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                                Module1.conflevel.Add(80);
        //                                Module1.conflevel.Add(80);
        //                                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                                Module1.objRetStructF3.Flag = "Y";
        //                                Module1.objRetStructF3.ManualConfirmation = "Y";
        //                                Module1.objRetStructF3.Status = "S";
        //                                return Module1.objRetStructF3;
        //                            }

        //                        case 3:
        //                            {
        //                                Module1.objRetStructF3.Words = Module1.PossibleWords;
        //                                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
        //                                Module1.conflevel.Add(70);
        //                                Module1.conflevel.Add(70);
        //                                Module1.conflevel.Add(70);
        //                                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
        //                                Module1.objRetStructF3.Flag = "Y";
        //                                Module1.objRetStructF3.ManualConfirmation = "Y";
        //                                Module1.objRetStructF3.Status = "S";
        //                                return Module1.objRetStructF3;
        //                            }
        //                    }
        //                }
        //            }


        //            // end of the F3
        //        }
        //    }
        //    catch (Exception)
        //    {
        //       // throw;
        //    }

        //    return Module1.objRetStructF3;
        //}

        //private clsCnCWord ConsigneeTel(clsCncMetaData oMeta, int intCurrPageNumber, int lineno, clsCnCWord TopWord)
        //{
        //    clsCnCWord cncTel = new clsCnCWord();
        //    int KeywordLeft = FindX1(TopWord);

        //    for (int i = 3; i <= 10; i++)
        //    {
        //        if (lineno + i <= oMeta.Page[intCurrPageNumber].LineCount)
        //        {
        //            int currLineNo = lineno + i;
        //            clsCnCLine NewLine = oMeta.Page[intCurrPageNumber].Line[currLineNo];
        //            if (Module1.RegTel.IsMatch(NewLine.strLine))
        //            {
        //                foreach (Match match in Module1.RegTel.Matches(NewLine.strLine))
        //                {
        //                    clsCnCWord[] mergeWords = new clsCnCWord[3];
        //                    //Match tel = match.Value;
        //                    string strTel = match.Value;
        //                    String[] splitTel = strTel.Split();
        //                    for (int iWord = 1; iWord <= oMeta.Page[intCurrPageNumber].Line[currLineNo].WordCount; iWord++)
        //                    {
        //                        if (oMeta.Page[intCurrPageNumber].Line[currLineNo].Word[iWord].strWord.ToUpper() == splitTel[0].ToUpper())
        //                        {
        //                            int TelLeft = FindX1(oMeta.Page[intCurrPageNumber].Line[currLineNo].Word[iWord]);

        //                            if (TelLeft >= KeywordLeft - 300 && TelLeft <= KeywordLeft + 300)
        //                            {
        //                                for (int j = 0; j < splitTel.Length; j++)
        //                                {
        //                                    mergeWords[j] = oMeta.Page[intCurrPageNumber].Line[currLineNo].Word[iWord + j];
        //                                }
        //                                cncTel = Module1.oClsCnc.MergeWords(mergeWords);

        //                                i = 7;

        //                            }
        //                            //iWord = oMeta.Page[intCurrPageNumber].Line[lineno + i].WordCount;
        //                        }
        //                    }
        //                }
        //            }


        //        }
        //    }
        //    return cncTel;

        //}
        //#endregion       

        #region Common Functions
        private Boolean CheckforExistWord(List<clsCnCWord> possibleWords, clsCnCWord NewWord)
        {
            Boolean exist = false;
            try
            {
                foreach (clsCnCWord Word in possibleWords)
                {
                    if (Word.strWord != null)
                    {                        
                        // string StrWord = Word.strWord.Trim();
                        string StrWord = Word.strWord.Replace(" ", "");
                        if (Module1.SpecialCharacters.Contains(StrWord[0].ToString()))
                        {
                            StrWord = StrWord.Remove(0, 1);
                        }
                        if (Module1.SpecialCharacters.Any(StrWord.EndsWith))
                        {
                            StrWord = StrWord.Remove(StrWord.Length - 1, 1);
                        }
                        //string StrNewWord = NewWord.strWord.Trim();
                        string StrNewWord = NewWord.strWord.Replace(" ", "");
                        if (Module1.SpecialCharacters.Contains(StrNewWord[0].ToString()))
                        {
                            StrNewWord = StrNewWord.Remove(0, 1);
                        }
                        if (Module1.SpecialCharacters.Any(StrNewWord.EndsWith))
                        {
                            StrNewWord = StrNewWord.Remove(StrNewWord.Length - 1, 1);
                        }
                        if (Word.PageNo.Equals(NewWord.PageNo)&&Word.LineNo.Equals(NewWord.LineNo) && Word.WordNumber.Equals(NewWord.WordNumber))
                        {
                            exist = true;
                        }
                        else if (StrWord.ToUpper().Equals(StrNewWord.ToUpper()))
                        {
                            //if (Word.LineNo.Equals(NewWord.LineNo))
                            //{
                            exist = true;
                            //}
                            //else
                            //{
                            //    if (exist != true)
                            //    {
                            //        exist = false;
                            //    }
                            //}
                        }
                        else if (StrWord.ToUpper().Contains("~"))                 //if existing word of possiblewordlist contains ~
                        {
                            string[] strWordSplitTilda = StrWord.Split('~');
                            if (strWordSplitTilda.Contains(StrNewWord))
                            {
                                exist = true;
                            }
                        }
                        else if(StrNewWord.ToUpper().Contains("~"))                 //if new word contains ~
                        {
                            string[] NewWordSplitTilda = StrNewWord.Split('~');
                            if(NewWordSplitTilda.Contains(StrWord))
                            {
                                exist = true;
                            }
                        }
                        
                        
                        else
                        {
                            if (exist != true)
                            {
                                exist = false;
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return exist;

        }
        private clsCnCWord FindBottomfromCoordinates(clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCLine CurrentLine, clsCnCWord CurrentWord, int length, int FieldNo, clsCnCWord[] TopKeyword, int iRecursiveCallCount,string keyword)
        {
            //clsCNCSBR ObjSbr = new clsCNCSBR();
            int CurrentX1 = FindX1(CurrentWord);
            //int CurrentX2 = FindX2(CurrentWord);
            int Line = CurrentLine.LineNo;
            clsCnCWord PossibleNumber = new clsCnCWord();
            //Regex regex = new Regex("^([a-zA-Z]*[-,./:%]?[0-9]+[-,./:%]?[a-zA-Z]*[-,./:%]?)*$");
            //Regex regex1 = new Regex("^([a-zA-Z]*[-,./#%]?[a-zA-Z]*[-,./#%]?[0-9]+[-,./#%]?)*$");
            //Regex Time = new Regex(@"^(?:[01]?[0-9]|2[0-3]):[0-5][0-9]$");
            clsCnCBoundingWords ObjBoundingWord = new clsCnCBoundingWords();
            clsCnCWord topWord = Module1.oClsCnc.MergeWords(TopKeyword);
            int KeywordBottom = topWord.Bottom;
            int Keywordheight = topWord.Bottom - topWord.Top;
            List<string> Separators = new List<string>() { ",", "&", "AND", "/" };
            char ch = ',';
            Boolean FlagShipment = false;

            try
            {
                if (FieldNo == 2)
                {
                    for (int i = 1; i <= 2; i++)
                    {
                        if (Line + i <= ObjMetaData.Page[intCurrPageNumber].LineCount)
                        {
                            clsCnCLine NewLine = ObjMetaData.Page[intCurrPageNumber].Line[Line + i];
                            for (int iword = 1; iword <= NewLine.WordCount; iword++)
                            {
                                int NewX1 = FindX1(NewLine.Word[iword]);

                                if (NewX1 >= CurrentX1 - 100 && NewX1 <= CurrentX1 + 205)
                                {
                                    string word = NewLine.Word[iword].strWord;
                                    
                                    word = word.Replace("'", "");
                                    if (word.Length > 2)
                                    {
                                        // word = word.Replace("'", "");
                                        if (Module1.SpecialCharacters.Contains(word[0].ToString()) && (!word.StartsWith("(")))
                                        {
                                            word = word.Remove(0, 1);
                                        }
                                        if (Module1.SpecialCharacters.Any(word.EndsWith) && (!word.EndsWith(")")))
                                        {
                                            word = word.Remove(word.Length - 1, 1);
                                        }
                                        word = Module1.ReplaceBraces(word);
                                    }
                                    RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");
                                    if (Obj008.Status == "S")
                                    {
                                        if (word.All(char.IsLetterOrDigit))
                                        {
                                            Obj008.Status = "F";
                                        }
                                    }

                                    int NumberOfSpecialCharaters = word.Count(c => !char.IsLetterOrDigit(c));
                                    int NumberofCharandNum = word.Count(c => char.IsLetterOrDigit(c));
                                    string WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                                    if (!string.IsNullOrEmpty(word))
                                    {
                                        //if ((Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word != Module1.ProNo && word.Length >= 3 && NumberOfSpecialCharaters<NumberofCharandNum)
                                        if ((!(Module1.NewRegexForAlphanumeric.IsMatch(word))) && (!WordToChckLength.All(char.IsLetter)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word != Module1.ProNo && word.Length >= 3 && NumberOfSpecialCharaters < NumberofCharandNum)
                                        {
                                            //if (word.Count(char.IsLetter) <= word.Count(char.IsDigit))
                                            //{
                                            //if (word.Count(char.IsDigit) == 1)
                                            //{
                                                int wordTop = NewLine.Word[iword].Top;
                                                int height = wordTop - KeywordBottom;

                                            if (height <= Math.Round(Keywordheight * 3.1))
                                            {
                                                clsCnCBoundingWords NewObjBoundingWordforLeft = Module1.oClsCnc.GetBoundingWords(NewLine.Word[iword], ObjMetaData, intCurrPageNumber);
                                                if (NewObjBoundingWordforLeft != null && NewObjBoundingWordforLeft.LeftWord != null)
                                                {
                                                    //if (NewObjBoundingWordforLeft.LeftWord.strWord[NewObjBoundingWordforLeft.LeftWord.strWord.Length - 1] != ':')
                                                    if (!Module1.IgnoreWordforLeftSplChar.Any(NewObjBoundingWordforLeft.LeftWord.strWord.EndsWith))
                                                    {
                                                        //if (word.All(char.IsDigit))
                                                        if (NewLine.Word[iword].strWord.All(char.IsDigit))
                                                        {
                                                            clsCnCWord CncPoWord = NewLine.Word[iword];
                                                            int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                            int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                            if (word.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                            else if (word.Length >= 4 && CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                            else
                                                            {
                                                                clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                {
                                                                    CncPoWord = MergedCncPoWord;
                                                                }
                                                                string strCncBLWord = CncPoWord.strWord.Replace(" ", "");
                                                                if (strCncBLWord.Length >= 4)
                                                                {
                                                                    CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                                    PossibleNumber = CncPoWord;
                                                                    iword = NewLine.WordCount;
                                                                    i = 2;
                                                                }
                                                            }

                                                        }
                                                        else
                                                        {
                                                            word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                                                            if (word.Length >= 4)
                                                            {
                                                                clsCnCWord CncPoWord = NewLine.Word[iword];
                                                                clsCnCBoundingWords NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                                                int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                                int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                                if (CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                                {
                                                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                                }
                                                                else if (CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                                {
                                                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                                }
                                                                else if (NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                                                {
                                                                    foreach (string sepratorStr in Separators)
                                                                    {
                                                                        if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                                        {
                                                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, sepratorStr, iRecursiveCallCount,FieldNo);

                                                                            break;
                                                                        }
                                                                    }

                                                                }
                                                                else if (CncPoWord.strWord.Contains(",") || CncPoWord.strWord.Contains("/") || CncPoWord.strWord.Contains(";"))  // 15/12/2020
                                                                {
                                                                    if (CncPoWord.strWord.Contains(","))
                                                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ",", iRecursiveCallCount); }
                                                                    //else if (CncPoWord.strWord.Contains("-"))
                                                                    //{ CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "-", iRecursiveCallCount); }
                                                                    else if (CncPoWord.strWord.Contains("/"))
                                                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "/", iRecursiveCallCount); }
                                                                    else if (CncPoWord.strWord.Contains(";"))
                                                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ";", iRecursiveCallCount); }

                                                                    //if (Convert.ToInt32(CncPoWord.WordNumber) != 0)
                                                                    if (!CncPoWord.strWord.Contains("~"))
                                                                    {
                                                                        CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                    }
                                                                }

                                                                else
                                                                {
                                                                    clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                    //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                    if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                    {
                                                                        CncPoWord = MergedCncPoWord;
                                                                    }
                                                                    CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                                }
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                            else { break; }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (NewLine.Word[iword].strWord.All(char.IsDigit))
                                                    {
                                                        clsCnCWord CncPoWord = NewLine.Word[iword];
                                                        int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                        int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                        if (word.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                        {
                                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                            PossibleNumber = CncPoWord;
                                                            iword = NewLine.WordCount;
                                                            i = 2;
                                                        }
                                                        else if (word.Length >= 4 && CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                        {
                                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                            PossibleNumber = CncPoWord;
                                                            iword = NewLine.WordCount;
                                                            i = 2;
                                                        }
                                                        else
                                                        {
                                                            clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                            //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                            if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                            {
                                                                CncPoWord = MergedCncPoWord;
                                                            }
                                                            string strCncBLWord = CncPoWord.strWord.Replace(" ", "");
                                                            if (strCncBLWord.Length >= 4)
                                                            {
                                                                CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                        }

                                                    }
                                                    else
                                                    {
                                                        word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                                                        if (word.Length >= 4)
                                                        {
                                                            clsCnCWord CncPoWord = NewLine.Word[iword];
                                                            clsCnCBoundingWords NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                                            int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                            int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                            if (CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                            }
                                                            else if (CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                            }
                                                            else if (NewObjBoundingWord != null && NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                                            {
                                                                foreach (string sepratorStr in Separators)
                                                                {
                                                                    if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                                    {
                                                                        CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, sepratorStr, iRecursiveCallCount,FieldNo);

                                                                        break;
                                                                    }
                                                                }

                                                            }
                                                            else if (CncPoWord.strWord.Contains(",") || CncPoWord.strWord.Contains("/") || CncPoWord.strWord.Contains(";"))  // 15/12/2020
                                                            {
                                                                if (CncPoWord.strWord.Contains(","))
                                                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ",", iRecursiveCallCount); }
                                                                //else if (CncPoWord.strWord.Contains("-"))
                                                                //{ CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "-", iRecursiveCallCount); }
                                                                else if (CncPoWord.strWord.Contains("/"))
                                                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "/", iRecursiveCallCount); }
                                                                else if (CncPoWord.strWord.Contains(";"))
                                                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ";", iRecursiveCallCount); }

                                                                //if (Convert.ToInt32(CncPoWord.WordNumber) != 0)
                                                                if (!CncPoWord.strWord.Contains("~"))
                                                                {
                                                                    CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                }
                                                            }

                                                            else
                                                            {
                                                                clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                {
                                                                    CncPoWord = MergedCncPoWord;
                                                                }
                                                                CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                            }
                                                            PossibleNumber = CncPoWord;
                                                            iword = NewLine.WordCount;
                                                            i = 2;
                                                        }
                                                        else { break; }
                                                    }
                                                }
                                            }
                                            //}
                                        }
                                        else { break; }
                                    }
                                }
                            }
                        }
                    }
                }
                else
                {
                    for (int i = 1; i <= 2; i++)
                    {
                        if (Line + i <= ObjMetaData.Page[intCurrPageNumber].LineCount)
                        {
                            clsCnCLine NewLine = ObjMetaData.Page[intCurrPageNumber].Line[Line + i];
                            for (int iword = 1; iword <= NewLine.WordCount; iword++)
                            {
                                int NewX1 = FindX1(NewLine.Word[iword]);
                                //int NewX2 = FindX2(NewLine.Word[iword]);

                                // if ((NewX1 >= CurrentX1 && NewX1 <= CurrentX1 + 200) && (NewX2 >= CurrentX2 && NewX2 <= CurrentX2 + 200))
                                //if ((NewX1 >= CurrentX1 && NewX1 <= CurrentX1 + 200))
                                if (NewX1 >= CurrentX1 - 50 && NewX1 <= CurrentX1 + 200)
                                {
                                    string word = NewLine.Word[iword].strWord;
                                    word = word.Replace("'", "");
                                    if (word.Length > 2)
                                    {
                                        // word = word.Replace("'", "");
                                        if (Module1.SpecialCharacters.Contains(word[0].ToString()) && (!word.StartsWith("(")))
                                        {
                                            word = word.Remove(0, 1);
                                        }
                                        if (Module1.SpecialCharacters.Any(word.EndsWith) && (!word.EndsWith(")")))
                                        {
                                            word = word.Remove(word.Length - 1, 1);
                                        }
                                    }
                                    RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");
                                    if (Obj008.Status == "S")
                                    {
                                        if (word.All(char.IsLetterOrDigit))
                                        {
                                            Obj008.Status = "F";
                                        }
                                    }
                                    if (FieldNo == 3 && (keyword.ToUpper() == "SHIPMENT" || keyword.ToUpper() == "SHIPMENTS"))
                                    {
                                        if (word.Count(char.IsLetter) >= word.Count(char.IsDigit))
                                        {
                                            FlagShipment = true;
                                        }
                                    }
                                    string WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                                    if (!string.IsNullOrEmpty(word))
                                    {
                                        //if ((Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word.Length >= length && word != Module1.ProNo && FlagShipment == false)
                                        if ((!(Module1.NewRegexForAlphanumeric.IsMatch(word))) && (!WordToChckLength.All(char.IsLetter)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word.Length >= length && word != Module1.ProNo && FlagShipment == false)
                                        {
                                            //if (word.Count(char.IsLetter) <= word.Count(char.IsDigit))
                                            //{
                                            //if (word.Count(char.IsDigit) == 1)
                                            //{
                                                int wordTop = NewLine.Word[iword].Top;
                                                int height = wordTop - KeywordBottom;

                                            if (height <= Keywordheight * 3.1)
                                            {
                                                clsCnCBoundingWords NewObjBoundingWordforLeft = Module1.oClsCnc.GetBoundingWords(NewLine.Word[iword], ObjMetaData, intCurrPageNumber);
                                                if (NewObjBoundingWordforLeft != null && NewObjBoundingWordforLeft.LeftWord != null)
                                                {
                                                    //if (NewObjBoundingWordforLeft.LeftWord.strWord[NewObjBoundingWordforLeft.LeftWord.strWord.Length - 1] != ':')
                                                    if (!Module1.IgnoreWordforLeftSplChar.Any(NewObjBoundingWordforLeft.LeftWord.strWord.EndsWith))
                                                    {
                                                        //if (word.All(char.IsDigit))
                                                        if (NewLine.Word[iword].strWord.All(char.IsDigit))
                                                        {
                                                            clsCnCWord CncPoWord = NewLine.Word[iword];
                                                            int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                            int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                            clsCnCBoundingWords NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                                            if (word.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                            else if (word.Length >= 4 && CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                            else if (NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                                            {
                                                                foreach (string sepratorStr in Separators)
                                                                {
                                                                    if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                                    {
                                                                        CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, sepratorStr, iRecursiveCallCount,FieldNo);
                                                                        PossibleNumber = CncPoWord;
                                                                        iword = NewLine.WordCount;
                                                                        i = 2;
                                                                        break;
                                                                    }
                                                                }

                                                            }
                                                            else
                                                            {
                                                                clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                {
                                                                    CncPoWord = MergedCncPoWord;
                                                                }
                                                                string strCncBLWord = CncPoWord.strWord.Replace(" ", "");
                                                                if (strCncBLWord.Length >= 4)
                                                                {
                                                                    CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                                    PossibleNumber = CncPoWord;
                                                                    iword = NewLine.WordCount;
                                                                    i = 2;
                                                                }
                                                            }
                                                        }
                                                        else
                                                        {
                                                            word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                                                            if (word.Length >= 4)
                                                            {
                                                                clsCnCWord CncPoWord = NewLine.Word[iword];
                                                                clsCnCBoundingWords NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                                                int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                                int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                                if (CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                                {
                                                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                                }
                                                                else if (CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                                {
                                                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                                }
                                                                else if (NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                                                {
                                                                    foreach (string sepratorStr in Separators)
                                                                    {
                                                                        if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                                        {
                                                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, sepratorStr, iRecursiveCallCount,FieldNo);

                                                                            break;
                                                                        }
                                                                    }

                                                                }
                                                                else if (CncPoWord.strWord.Contains(",") || CncPoWord.strWord.Contains("/") || CncPoWord.strWord.Contains(";"))  // 15/12/2020
                                                                {
                                                                    if (CncPoWord.strWord.Contains(","))
                                                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ",", iRecursiveCallCount); }
                                                                    //else if (CncPoWord.strWord.Contains("-"))
                                                                    //{ CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "-", iRecursiveCallCount); }
                                                                    else if (CncPoWord.strWord.Contains("/"))
                                                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "/", iRecursiveCallCount); }
                                                                    else if (CncPoWord.strWord.Contains(";"))
                                                                    { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ";", iRecursiveCallCount); }

                                                                    // if (Convert.ToInt32(CncPoWord.WordNumber) != 0)
                                                                    if (!CncPoWord.strWord.Contains("~"))
                                                                    {
                                                                        CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                    }
                                                                }

                                                                else
                                                                {
                                                                    clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                    //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                    if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                    {
                                                                        CncPoWord = MergedCncPoWord;
                                                                    }
                                                                    CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                                }
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                            else { break; }
                                                        }
                                                    }
                                                }
                                                else
                                                {
                                                    if (NewLine.Word[iword].strWord.All(char.IsDigit))
                                                    {
                                                        clsCnCWord CncPoWord = NewLine.Word[iword];
                                                        int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                        int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                        clsCnCBoundingWords NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                                        if (word.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                        {
                                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                            PossibleNumber = CncPoWord;
                                                            iword = NewLine.WordCount;
                                                            i = 2;
                                                        }
                                                        else if (word.Length >= 4 && CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                        {
                                                            CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                            PossibleNumber = CncPoWord;
                                                            iword = NewLine.WordCount;
                                                            i = 2;
                                                        }
                                                        else if (NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                                        {
                                                            foreach (string sepratorStr in Separators)
                                                            {
                                                                if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                                {
                                                                    CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, sepratorStr, iRecursiveCallCount,FieldNo);
                                                                    PossibleNumber = CncPoWord;
                                                                    iword = NewLine.WordCount;
                                                                    i = 2;
                                                                    break;
                                                                }
                                                            }

                                                        }
                                                        else
                                                        {
                                                            clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                            //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                            if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                            {
                                                                CncPoWord = MergedCncPoWord;
                                                            }
                                                            string strCncBLWord = CncPoWord.strWord.Replace(" ", "");
                                                            if (strCncBLWord.Length >= 4)
                                                            {
                                                                CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                                PossibleNumber = CncPoWord;
                                                                iword = NewLine.WordCount;
                                                                i = 2;
                                                            }
                                                        }
                                                    }
                                                    else
                                                    {
                                                        word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                                                        if (word.Length >= 4)
                                                        {
                                                            clsCnCWord CncPoWord = NewLine.Word[iword];
                                                            clsCnCBoundingWords NewObjBoundingWord = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                                                            int CommaCount = CncPoWord.strWord.Count(c => (c == ch));
                                                            int SlashCount = CncPoWord.strWord.Count(c => (c == '/'));
                                                            if (CncPoWord.strWord.EndsWith(",") && CommaCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, ",", iRecursiveCallCount,FieldNo);
                                                            }
                                                            else if (CncPoWord.strWord.EndsWith("/") && SlashCount == 1)
                                                            {
                                                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, Line + i, CncPoWord, "/", iRecursiveCallCount,FieldNo);
                                                            }
                                                            else if (NewObjBoundingWord.RightWord != null && Separators.Contains(NewObjBoundingWord.RightWord.strWord.ToUpper()))
                                                            {
                                                                foreach (string sepratorStr in Separators)
                                                                {
                                                                    if (NewObjBoundingWord.RightWord.strWord.ToUpper().Contains(sepratorStr))
                                                                    {
                                                                        CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, sepratorStr, iRecursiveCallCount,FieldNo);

                                                                        break;
                                                                    }
                                                                }

                                                            }
                                                            else if (CncPoWord.strWord.Contains(",") || CncPoWord.strWord.Contains("/") || CncPoWord.strWord.Contains(";"))  // 15/12/2020
                                                            {
                                                                if (CncPoWord.strWord.Contains(","))
                                                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ",", iRecursiveCallCount); }
                                                                //else if (CncPoWord.strWord.Contains("-"))
                                                                //{ CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "-", iRecursiveCallCount); }
                                                                else if (CncPoWord.strWord.Contains("/"))
                                                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "/", iRecursiveCallCount); }
                                                                else if (CncPoWord.strWord.Contains(";"))
                                                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ";", iRecursiveCallCount); }

                                                                // if (Convert.ToInt32(CncPoWord.WordNumber) != 0)
                                                                if (!CncPoWord.strWord.Contains("~"))
                                                                {
                                                                    CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                }
                                                            }

                                                            else
                                                            {
                                                                clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, Line + i, CncPoWord, keyword);
                                                                //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                                                {
                                                                    CncPoWord = MergedCncPoWord;
                                                                }
                                                                CncPoWord = BottomMultiValue(ObjMetaData, intCurrPageNumber, Module1.oClsCnc.MergeWords(TopKeyword), CncPoWord, iRecursiveCallCount,FieldNo);
                                                            }
                                                            PossibleNumber = CncPoWord;
                                                            iword = NewLine.WordCount;
                                                            i = 2;
                                                        }
                                                        else { break; }
                                                    }
                                                }
                                            }
                                            
                                        }
                                        else { break; }
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
                ObjBoundingWord = null;
                Separators = null;
            }
            return PossibleNumber;

        }
        private int FindX1(clsCnCWord CurrentWord)
        {
            //string x1 = CurrentWord.X1Char;
            //string[] arrX1 = x1.Split(',');
            //int x1Coordinate = Convert.ToInt32(arrX1[0]);
            //return x1Coordinate;
            int left = CurrentWord.Left;
            return left;
        }
        private int FindX2(clsCnCWord CurrentWord)
        {
            string x2 = CurrentWord.X2Char;
            string[] arrX2 = x2.Split(',');
            int x2Coordinate = Convert.ToInt32(arrX2[0]);
            return x2Coordinate;
        }
        //private int CheckTopWord(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, List<clsCnCWord> PWords, clsCnCWord[] TopKeywords, int iLine, int iWord)
        //{
        //    ClsCNC objCNC = new ClsCNC();
        //    clsCNCSBR ObjSbr = new clsCNCSBR();
        //    Regex regex = new Regex("^([a-zA-Z]*[-,.#/]?[0-9]+[-,.#/]?[a-zA-Z]*[-,.#/]?)*$");
        //    Regex regex1 = new Regex("^([a-zA-Z]*[-,./#%]?[a-zA-Z]*[-,./#%]?[0-9]+[-,./#%]?)*$");
        //    // Regex InvalidDate = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.]*(Map|MAP|mAP|MAp)(\d{4}|(\d{3}[a-zA-Z])|(\d{1}[a-zA-Z]\d{2}))");

        //    for (int j = 0; j < ObjBoundingWord.TopWords.Length; j++)
        //    {
        //        clsCnCWord CncSNWord = ObjBoundingWord.TopWords[j];
        //        string word = CncSNWord.strWord;

        //        RetStructIQSBR008 Obj008 = ObjSbr.IQSBR008(word, "E");

        //        if ((regex.IsMatch(word) || regex1.IsMatch(word)) && (!Module1.InvalidDate.IsMatch(word)) && Obj008.Status == "F" && word.Length >= 4)
        //        {
        //            Boolean exist = CheckforExistWord(PWords, CncSNWord);
        //            if (exist == false)
        //            {
        //                PWords.Add(CncSNWord);
        //                j = ObjBoundingWord.TopWords.Length;
        //                iWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount;
        //            }
        //        }

        //    }
        //    return iWord;
        //}
        //private clsCnCWord FindTopfromCoordinates(clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCLine CurrentLine, clsCnCWord CurrentWord)
        //{
        //    clsCNCSBR ObjSbr = new clsCNCSBR();
        //    int CurrentX1 = FindX1(CurrentWord);
        //    int CurrentX2 = FindX2(CurrentWord);
        //    int Line = CurrentLine.LineNo;
        //    clsCnCWord PossibleNumber = new clsCnCWord();
        //    Regex regex = new Regex("^([a-zA-Z]*[-,.#/]?[0-9]+[-,.#/]?[a-zA-Z]*[-,.#/]?)*$");
        //    Regex regex1 = new Regex("^([a-zA-Z]*[-,./#%]?[a-zA-Z]*[-,./#%]?[0-9]+[-,./#%]?)*$");
        //    //Regex InvalidDate = new Regex(@"(3[01]|[12][0-9]|0?[1-9]|[1-9])[-/.]*(Map|MAP|mAP|MAp)(\d{4}|(\d{3}[a-zA-Z])|(\d{1}[a-zA-Z]\d{2}))");

        //    for (int i = 1; i <= 2; i++)
        //    {
        //        if (Line - i >= 1)
        //        {
        //            clsCnCLine NewLine = ObjMetaData.Page[intCurrPageNumber].Line[Line - i];
        //            for (int iword = 1; iword <= NewLine.WordCount; iword++)
        //            {
        //                int NewX1 = FindX1(NewLine.Word[iword]);
        //                int NewX2 = FindX2(NewLine.Word[iword]);

        //                if ((NewX1 >= CurrentX1 && NewX1 <= CurrentX1 + 300) && (NewX2 >= CurrentX2 && NewX2 <= CurrentX2 + 300))
        //                {
        //                    string word = NewLine.Word[iword].strWord;
        //                    RetStructIQSBR008 Obj008 = ObjSbr.IQSBR008(word, "E");
        //                    if ((regex.IsMatch(word) || regex1.IsMatch(word)) && (!Module1.InvalidDate.IsMatch(word)) && Obj008.Status == "F" && word.Length >= 5)
        //                    {
        //                        PossibleNumber = NewLine.Word[iword];
        //                        iword = NewLine.WordCount;
        //                        i = 2;
        //                    }
        //                }
        //            }
        //        }
        //    }
        //    return PossibleNumber;

        //}
        //private List<clsCnCWord> GetTopThreeValue(List<clsCnCWord> ArrayList)
        //{
        //    var lstTopThreeValue = new List<clsCnCWord>();
        //    int iCount = 0;
        //    try
        //    {
        //        foreach (clsCnCWord Word in ArrayList)
        //        {
        //            if (iCount > 2)
        //            {
        //                break;
        //            }

        //            lstTopThreeValue.Add(Word);
        //            iCount = iCount + 1;
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //MessageBox.Show(ex.Message);
        //    }

        //    return lstTopThreeValue;
        //}

        //To find Multiple Values
        private clsCnCWord RightMultiValue(clsCncMetaData ObjMetaData, int intCurrPageNumber, int iLine, clsCnCWord FirstWord, string Separator, int iRecursiveCallCount,int field)
        {
            Module1.lstConfFlag.Clear();
            Module1.lstRemarks.Clear();
            clsCnCWord MultiWord = new clsCnCWord();

            int WordNo = Convert.ToInt32(FirstWord.WordNumber);
            clsCnCWord[] Word = new clsCnCWord[20];
            List<string> LstStrWord = new List<string>();
            int counter = 0;
            string CurrentLine = ObjMetaData.Page[intCurrPageNumber].Line[iLine].strLine;
            //string[] splitWords = CurrentLine.Split(',');
            string[] splitWords = CurrentLine.Split(Separator.ToArray());
            string strFirstWord = FirstWord.strWord.Replace(",", "");
            if (strFirstWord.Length > 2)
            {
                strFirstWord = ReplaceStartEndSpecialchar(strFirstWord);
            }
            string pattern = GetPattern(strFirstWord);
            int index = 1, NxtWordCount = 1;

            Word[0] = FirstWord;
            if (FlagForMultiValues(strFirstWord) == false)
            {
                Module1.lstConfFlag.Add("0");
                Module1.lstRemarks.Add("NA");
            }
            else
            {
                Boolean ConfidenceFlag = HighestConfidence(FirstWord,field);

                if (ConfidenceFlag == true)
                //if (FirstWord.Confidence >= 85)
                {
                    if (iRecursiveCallCount == 1)
                    {
                        Module1.lstConfFlag.Add("0");
                        Module1.lstRemarks.Add("NA");
                    }
                    else
                    {
                        Module1.lstConfFlag.Add("1");
                        Module1.lstRemarks.Add("1F");
                    }
                }
                else
                {
                    Module1.lstConfFlag.Add("0");
                    Module1.lstRemarks.Add("NA");
                }
            }
            for (int i = 0; i < splitWords.Length; i++)
            {
                if (counter == 0)
                {
                    if (splitWords[i].Contains(strFirstWord))
                    {
                        if (i < splitWords.Length - 1)
                        {
                            clsCnCWord NewWord;
                            if (Word[0].strWord.EndsWith(Separator))
                            {
                                //NewWord = ModifyWord(Word[0], string.Concat(Word[0].strWord.Remove(Word[0].strWord.Length - 1), "~"));
                                NewWord = ModifyWord(Word[0], Word[0].strWord.Remove(Word[0].strWord.Length - 1));
                            }
                            else
                            {
                                // NewWord = ModifyWord(Word[0], string.Concat(Word[0].strWord, "~"));
                                NewWord = ModifyWord(Word[0], Word[0].strWord);
                            }
                            Word[0] = NewWord;
                            LstStrWord.Add(NewWord.strWord);
                        }
                        counter = 1;
                    }
                }
                else
                {

                    string[] splitNew = splitWords[i].TrimStart().Split();   //if word contains space
                    string NextWordPattern = string.Empty;

                    if (splitNew.Length > 1)
                    {
                        if (!string.IsNullOrEmpty(splitNew[0]))
                        {
                            splitNew[0] = ReplaceStartEndSpecialchar(splitNew[0]);
                            NextWordPattern = GetPattern(splitNew[0]);
                        }
                        else
                        {
                            splitNew[1] = ReplaceStartEndSpecialchar(splitNew[1]);
                            NextWordPattern = GetPattern(splitNew[1]);
                        }
                        i = splitWords.Length;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(splitWords[i]))
                        {
                            splitWords[i] = ReplaceStartEndSpecialchar(splitWords[i]);
                            NextWordPattern = GetPattern(splitWords[i]);
                        }
                    }
                    RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(pattern, NextWordPattern);

                    if (obj50.PercentageMatch >= 90)
                    {
                        if (WordNo + NxtWordCount <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount)
                        {
                            if (ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[WordNo + NxtWordCount].strWord == Separator)
                            {
                                NxtWordCount = NxtWordCount + 1;
                            }

                            clsCnCWord NewWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[WordNo + NxtWordCount];

                            int NxtWordLeft = NewWord.Left;
                            int PrevWordRight = Word[index - 1].Right;
                            if (NxtWordLeft - PrevWordRight <= 70)
                            {

                                if (i < splitWords.Length - 1)
                                {

                                    if (NewWord.strWord.EndsWith(Separator))
                                    {
                                        //NewWord = ModifyWord(NewWord, string.Concat(NewWord.strWord.Remove(NewWord.strWord.Length - 1), "~"));
                                        NewWord = ModifyWord(NewWord, NewWord.strWord.Remove(NewWord.strWord.Length - 1));
                                    }
                                    else
                                    {
                                        //NewWord = ModifyWord(NewWord, string.Concat(NewWord.strWord, "~"));
                                        NewWord = ModifyWord(NewWord, NewWord.strWord);
                                    }
                                    //clsCnCWord NewWord = ModifyWord(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[WordNo + NxtWordCount], string.Concat(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[WordNo + NxtWordCount].strWord, "~"));
                                    if (!CheckExistWordforBottomMulti(Word, NewWord))
                                    {
                                        Word[index] = NewWord;
                                        LstStrWord.Add(NewWord.strWord);
                                    }
                                    else { continue; }
                                }
                                else
                                {
                                    if (NewWord.strWord.EndsWith(Separator))
                                    {
                                        NewWord = ModifyWord(NewWord, NewWord.strWord.Remove(NewWord.strWord.Length - 1));
                                    }
                                    else
                                    {
                                        NewWord = NewWord;
                                    }
                                    if (!CheckExistWordforBottomMulti(Word, NewWord))
                                    {
                                        Word[index] = NewWord;
                                        LstStrWord.Add(NewWord.strWord);
                                    }
                                    else { continue; }
                                }
                                //if (NewWord.Confidence >= 85)
                                if (FlagForMultiValues(NewWord.strWord) == false)
                                {
                                    Module1.lstConfFlag.Add("0");
                                    Module1.lstRemarks.Add("NA");
                                }
                                else
                                {
                                   Boolean ConfidenceFlag = HighestConfidence(NewWord,field);
                                    //if (Module1.PossibleWords[0].Confidence >= 85)
                                    if (ConfidenceFlag == true)
                                    {
                                        if (iRecursiveCallCount == 1)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else
                                        {
                                            Module1.lstConfFlag.Add("1");
                                            Module1.lstRemarks.Add("1F");
                                        }
                                    }
                                    else
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                }
                                index++; NxtWordCount++;
                            }
                            else { break; }
                        }
                        else
                        {
                            Word[index - 1] = null;
                            LstStrWord[index - 1] = null;
                            break;
                        }
                    }
                    else
                    {
                        //If the pattern of any value doesnt match then make flag of all the values red
                        //if(Separator==",")
                        //{
                            Module1.lstConfFlag.Clear();
                            Module1.lstRemarks.Clear();

                            for(int j=index-1;j>=0;j--)
                            {
                                Module1.lstConfFlag.Add("0");
                                Module1.lstRemarks.Add("NA");
                            }
                            //Module1.lstConfFlag.Add("0");
                            //Module1.lstRemarks.Add("NA");
                        //}
                        break;
                    }
                }

            }

            MultiWord = Module1.oClsCnc.MergeWords(Word);
            if (Word[1] == null && Word[2] == null)
            {
                MultiWord = FirstWord;
            }
            else
            {
                MultiWord = Module1.oClsCnc.MergeWords(Word);
                MultiWord.strWord = string.Join("~",LstStrWord.ToArray());
            }
            if (MultiWord.strWord.EndsWith("~"))
            {
                MultiWord.strWord = MultiWord.strWord.Remove(MultiWord.strWord.Length - 1);
            }
            MultiWord.Flag = string.Join("~", Module1.lstConfFlag.ToArray());
            MultiWord.Remarks = string.Join("~", Module1.lstRemarks.ToArray());
            MultiWord.Confidence = 200;

            return MultiWord;
        }
        private clsCnCWord BottomMultiValue(clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCWord Keyword, clsCnCWord FirstWord, int iRecursiveCallCount,int field)
        {
            //if (Module1.regex.IsMatch(FirstWord.strWord) || Module1.regex1.IsMatch(FirstWord.strWord))
            if (!Module1.NewRegexForAlphanumeric.IsMatch(FirstWord.strWord))
            {
                Module1.lstConfFlag.Clear();
                Module1.lstRemarks.Clear();
                //ClsCNC oCnc = new ClsCNC();
                clsCnCWord MultiWord = new clsCnCWord();
                clsCnCWord[] ArrayMultiWord = new clsCnCWord[20];
                List<string> lstMultiWord = new List<string>();

                int ROI_Left;
                //ROI_Left = Keyword.Left - 100;          
                ROI_Left = FirstWord.Left - 50;
                int ROI_Right = FirstWord.Right + 50;
                int ROI_Top = Keyword.Bottom;
                int ROI_Bottom = Keyword.Bottom + 300;
                //List<string> lstConfFlag = new List<string>();
                string StrFirst = FirstWord.strWord;
                StrFirst = ReplaceStartEndSpecialchar(StrFirst);
                string pattern = GetPattern(StrFirst);

                int index = 0;
                //int NxtWordCount = 1;
                clsCnCWord[] ROIData = Module1.oClsCnc.GetDataForROI(ObjMetaData.Page, intCurrPageNumber, ROI_Left, ROI_Top, ROI_Right, ROI_Bottom);
                if (ROIData != null)
                {
                    if (ROIData[0].strWord == FirstWord.strWord)
                    {
                        //Word[0] = ModifyWord(ROIData[i], ROIData[i].strWord + '~'); ;
                        for (int i = 0; i < ROIData.Length; i++)
                        {
                            string nextStrWord = ROIData[i].strWord;
                            nextStrWord = ReplaceStartEndSpecialchar(nextStrWord);
                            //string NextWordPattern = GetPattern(ROIData[i].strWord);
                            string NextWordPattern = GetPattern(nextStrWord);
                            RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(pattern, NextWordPattern);
                            //if (pattern == NextWordPattern)
                            if (obj50.PercentageMatch >= 90)
                            {
                                clsCnCWord NewWord = new clsCnCWord();
                                if (i < ROIData.Length - 1)
                                {
                                    //NewWord = ModifyWord(ROIData[i], ROIData[i].strWord + '~');
                                    NewWord = ModifyWord(ROIData[i], ROIData[i].strWord);
                                }
                                else
                                {
                                    NewWord = ModifyWord(ROIData[i], ROIData[i].strWord);
                                }
                                if (!CheckExistWordforBottomMulti(ArrayMultiWord, NewWord))    // if loop commented on 6/1/12 
                                {
                                    NewWord.strWord = ReplaceStartEndSpecialchar(NewWord.strWord);
                                    ArrayMultiWord[index] = NewWord;
                                    lstMultiWord.Add(NewWord.strWord);

                                    // if (ROIData[i].Confidence >= 85)
                                    if (NewWord.Flag == null)
                                    {
                                        if (FlagForMultiValues(NewWord.strWord) == false)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else
                                        {
                                            Boolean ConfidenceFlag = HighestConfidence(ROIData[i],field);

                                            if (ConfidenceFlag == true)
                                            {
                                                //if (iRecursiveCallCount == 1)
                                                //{
                                                //    Module1.lstConfFlag.Add("0");
                                                //    Module1.lstRemarks.Add("NA");
                                                //}
                                                //else
                                                //{
                                                Module1.lstConfFlag.Add("1");
                                                Module1.lstRemarks.Add("1F");
                                                //}
                                            }
                                            else
                                            {
                                                Module1.lstConfFlag.Add("0");
                                                Module1.lstRemarks.Add("NA");
                                            }
                                        }
                                    }
                                    else
                                    {
                                        Module1.lstConfFlag.Add(NewWord.Flag);
                                        Module1.lstRemarks.Add(NewWord.Remarks);
                                    }
                                    index++;

                                }
                            }
                            else
                            {
                                if (ROIData[i].strWord == "/")
                                {
                                    continue;
                                }
                                else { break; }
                            }
                        }
                    }
                    else
                    {
                        ArrayMultiWord[0] = FirstWord;
                        lstMultiWord.Add(FirstWord.strWord);
                        //if (FirstWord.Confidence >= 85)
                        if (FirstWord.Flag == null)
                        {
                            if (FlagForMultiValues(FirstWord.strWord) == false)
                            {
                                Module1.lstConfFlag.Add("0");
                                Module1.lstRemarks.Add("NA");
                            }
                            else
                            {
                                Boolean ConfidenceFlag = HighestConfidence(FirstWord,field);

                                if (ConfidenceFlag == true)
                                {
                                    if (iRecursiveCallCount == 1)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else
                                    {
                                        Module1.lstConfFlag.Add("1");
                                        Module1.lstRemarks.Add("1F");
                                    }
                                }
                                else
                                {
                                    Module1.lstConfFlag.Add("0");
                                    Module1.lstRemarks.Add("NA");
                                }
                            }
                        }
                        else
                        {
                            Module1.lstConfFlag.Add(FirstWord.Flag);
                            Module1.lstRemarks.Add(FirstWord.Remarks);
                        }
                    }
                    if (ArrayMultiWord[1] == null && ArrayMultiWord[2] == null)
                    {
                        MultiWord = FirstWord;
                    }
                    else
                    {
                        MultiWord = Module1.oClsCnc.MergeWords(ArrayMultiWord);
                        MultiWord.strWord = string.Join("~", lstMultiWord.ToArray());
                    }
                    // MultiWord = oCnc.MergeWords(Word);
                    MultiWord.Flag = string.Join("~", Module1.lstConfFlag.ToArray());
                    MultiWord.Remarks = string.Join("~", Module1.lstRemarks.ToArray());
                    MultiWord.Confidence = 200;
                    MultiWord.WordNumber = FirstWord.WordNumber;
                    //string Conf = string.Join("~", lstConfFlag.ToArray());
                    string strWord = MultiWord.strWord;
                    if (strWord.EndsWith("~"))
                    {
                        strWord = MultiWord.strWord.Remove(MultiWord.strWord.Length - 1);
                        MultiWord = ModifyWord(MultiWord, strWord);
                    }

                    return MultiWord;
                }
                else
                {
                    return FirstWord;
                }
            }
            else
            {
                return FirstWord;
            }
        }
        private clsCnCWord ContainsSeparatorwithMultiValue(clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCWord FirstWord, string Separator, int iRecursiveCallCount)
        {
            Module1.lstConfFlag.Clear();
            Module1.lstRemarks.Clear();
            clsCnCWord MultiWord = new clsCnCWord();
            string Currentword = ReplaceStartEndSpecialchar(FirstWord.strWord);
            //string Currentword = FirstWord.strWord.Replace("'","");
            string[] splitWords = Currentword.Split(Separator.ToArray());
            string confstring = FirstWord.ConfString;
            List<string> lstStr = new List<string>();
            char[] confnew = confstring.ToCharArray();
            Boolean duplicateFlag = false;
            //List<string> lstConfFlag = new List<string>();

            string frstStringPattern = string.Empty;
            if (splitWords[0].Length >= 4)
            {
                for (int iSplit = 0; iSplit < splitWords.Length; iSplit++)
                {
                    string conf = "1";
                    string remark = "1F";
                    if (iSplit == 0)
                    {
                        frstStringPattern = GetPattern(splitWords[0]);
                        lstStr.Add(splitWords[0]);
                        if (iRecursiveCallCount == 1)
                        {
                            conf = "0";
                            remark = "NA";
                        }
                        else if (FlagForMultiValues(splitWords[0]) == false)
                        {
                            conf = "0";
                            remark = "NA";
                        }
                        else
                        {
                            int index = Currentword.IndexOf(splitWords[iSplit]);
                            for (int i = index; i < Currentword.Length; i++)
                            {
                                // if (Currentword[i] != ',')
                                if (Currentword[i].ToString() != Separator)
                                {
                                    if (confstring[i] < '8')
                                    {
                                        conf = "0";
                                        remark = "NA";
                                        break;
                                    }
                                }
                                else
                                { break; }
                            }
                        }
                        Module1.lstConfFlag.Add(conf);
                        Module1.lstRemarks.Add(remark);
                    }
                    else
                    {
                        if (!Module1.SpecialCharacters.Contains(splitWords[iSplit]))
                        {
                            string NextStringPattern = GetPattern(splitWords[iSplit]);
                            RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(frstStringPattern, NextStringPattern);

                            if (obj50.PercentageMatch >= 90)
                            {
                                if (!lstStr.Contains(splitWords[iSplit]))
                                {
                                    lstStr.Add(splitWords[iSplit]);
                                    if (iRecursiveCallCount == 1)
                                    {
                                        conf = "0";
                                        remark = "NA";
                                    }
                                    else if(FlagForMultiValues(splitWords[iSplit]) ==false)
                                    {
                                        conf = "0";
                                        remark = "NA";
                                    }
                                    else
                                    {
                                        int index = Currentword.IndexOf(splitWords[iSplit]);
                                        for (int i = index; i < Currentword.Length; i++)
                                        {
                                            if (Currentword[i].ToString() != Separator)
                                            {
                                                if (confstring[i] < '8')
                                                {
                                                    conf = "0";
                                                    remark = "NA";
                                                    break;
                                                }
                                            }
                                            else
                                            { break; }
                                        }
                                    }
                                    Module1.lstConfFlag.Add(conf);
                                    Module1.lstRemarks.Add(remark);
                                }
                                else
                                {
                                    duplicateFlag = true;
                                }
                            }
                            else
                            {
                                if (Separator == "-" || Separator == "/" || Separator == ";")
                                {
                                    lstStr.Clear();
                                    lstStr.Add(FirstWord.strWord);
                                    break;
                                }
                                else
                                {
                                    continue;
                                }
                            }
                        }
                        else
                        {
                            continue;
                        }
                    }
                }
                if (lstStr.Count > 1)
                {
                    MultiWord = ModifyWord(FirstWord, string.Join("~", lstStr.ToArray()));
                    MultiWord.LineNo = FirstWord.LineNo;
                    MultiWord.WordNumber = FirstWord.WordNumber;
                    MultiWord.Flag = string.Join("~", Module1.lstConfFlag.ToArray());
                    MultiWord.Remarks = string.Join("~", Module1.lstRemarks.ToArray());
                    MultiWord.Confidence = 200;
                }
                else if(lstStr.Count==1 && duplicateFlag==true)
                {
                    MultiWord = ModifyWord(FirstWord, string.Join("~", lstStr.ToArray()));
                    MultiWord.LineNo = FirstWord.LineNo;
                    MultiWord.WordNumber = FirstWord.WordNumber;
                    MultiWord.Flag = string.Join("~", Module1.lstConfFlag.ToArray());
                    MultiWord.Remarks = string.Join("~", Module1.lstRemarks.ToArray());
                    MultiWord.Confidence = 200;
                }
                else
                {
                    MultiWord = FirstWord;
                }
            }
            else
            {
                MultiWord = FirstWord;
            }
            return MultiWord;
        }
        private string GetPattern(String inputText)
        {
            if (inputText.Trim().Length == 0)
            {
                return "";
            }
            int intAsciiCode = 0;
            string strTemp = "";
            inputText = inputText.Trim().ToLower();

            foreach (char Ch in inputText)
            {
                intAsciiCode = Convert.ToInt32(Ch);
                if ((intAsciiCode >= 65 && intAsciiCode <= 90) || (intAsciiCode >= 97 && intAsciiCode <= 122))
                    strTemp = strTemp + "a";
                else if (intAsciiCode >= 48 && intAsciiCode <= 57)
                    strTemp = strTemp + "n";
                else if (intAsciiCode == 32)
                    strTemp = strTemp + "b";
                else
                    strTemp = strTemp + "s";
            }

            return strTemp;
        }

        //To find next value with space
        private clsCnCWord FindNextWord(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, int iLine, clsCnCWord FirstNumericWord, string keyword)
        {
            //ClsCNC oCnc = new ClsCNC();
            clsCnCWord MergedNumericWord = new clsCnCWord();
            clsCnCWord[] NumericWords = new clsCnCWord[30];
            List<string> filterEndswith = new List<string>() { ":",";",".","*"};
            List<string> FilterRightWord = new List<string>() { "CARRIERS", "CARRIER", "PLEASE", "PACKAGES","MUST","THE","HOUR", "PAGE","PIECES", "CARTONS","APPT","LBS", "FT","SHIPMENT","QUOTE","CAMERM","CAMER","TOTAL","REFERENCE","REF","DELIVERY","DO","DEPT","INVOICE","MABD","PICKUP", "PICKUP#","ON","PRO","STOP"};
            List<string> ignoreKeywords = new List<string>() { "MANIFEST","MANIFEST ID", "RA", "RA#", "RMA", "RMA#", "RGA", "RGA#", "RETURN AUTHORIZATION", "RETURNAUTHORIZATION" };
            try
            {
                int WordNo = Convert.ToInt32(FirstNumericWord.WordNumber);
                string StrFirst = FirstNumericWord.strWord;
                StrFirst = ReplaceStartEndSpecialchar(StrFirst);
                //string pattern = GetPattern(FirstNumericWord.strWord);
                string pattern = GetPattern(StrFirst);
                NumericWords[0] = FirstNumericWord;
                int index = 1;
                Boolean PatternMatchFlag = false;
                for (int i = WordNo + 1; i <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; i++)
                {
                    string strPrevWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i - 1].strWord;
                    string strNextWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord;
                    int PrevWordRight = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i - 1].Right;
                    int NextWordLeft = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].Left;

                    if (i == WordNo + 1)
                    {
                        string ReplaceStrNextWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord;
                        ReplaceStrNextWord = ReplaceStartEndSpecialchar(ReplaceStrNextWord);
                        //string NextwordPattern = GetPattern(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord);
                        string NextwordPattern = GetPattern(ReplaceStrNextWord);
                        RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(pattern, NextwordPattern);
                        if (pattern == NextwordPattern)
                        {
                            PatternMatchFlag = true;
                            //MergedNumericWord = Module1.oClsCnc.MergeWords(NumericWords);
                            //MergedNumericWord.Flag = "0";
                            //MergedNumericWord.Remarks = "NA";
                            //for (int j = index - 1; j >= 0; j--)
                            //{
                            //    Module1.lstConfFlag.Add("0");
                            //    Module1.lstRemarks.Add("NA");
                            //}
                            break;
                        }
                    }
                    //if (strNextWord.All(char.IsDigit) && (NextWordLeft - PrevWordRight <= 40))
                    if (NextWordLeft - PrevWordRight <= 40 && (!strNextWord.Contains(":")) && (!FilterRightWord.Contains(strNextWord.ToUpper()))&&((!filterEndswith.Any(strPrevWord.EndsWith)))&&(!ignoreKeywords.Contains(keyword.ToUpper()))&&(strPrevWord!="24"))
                    {
                        string strword = Module1.ReplaceBraces(strNextWord);
                        // strNextWord = ReplaceStartEndSpecialchar(strNextWord);
                        //if (strNextWord.All(char.IsLetter) || (Module1.regex.IsMatch(strNextWord) || Module1.regex1.IsMatch(strNextWord))||(Module1.SpecialCharacters.Contains(strNextWord)))
                        if (strNextWord.All(char.IsLetter) || (!Module1.NewRegexForAlphanumeric.IsMatch(strword)) || (Module1.SpecialCharacters.Contains(strNextWord)))
                        {
                            NumericWords[index] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i];
                            index++;
                        }
                        else
                        {
                            break;
                        }
                    }
                    else
                    { break; }
                }                
                if (NumericWords[1] == null && NumericWords[2] == null)
                {
                    MergedNumericWord = FirstNumericWord;
                    if(PatternMatchFlag==true)
                    {
                        MergedNumericWord.Flag = "0";
                        MergedNumericWord.Remarks = "NA";
                    }
                }
                else
                {
                    MergedNumericWord = Module1.oClsCnc.MergeWords(NumericWords);
                }
                //MergedNumericWord = Module1.oClsCnc.MergeWords(NumericWords);
            }
            catch (Exception)
            { }
            finally
            {
                NumericWords = null;
            }
            return MergedNumericWord;
        }

        //private clsCnCWord FindNextWord(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, int iLine, clsCnCWord FirstNumericWord, string keyword)
        //{
        //    //ClsCNC oCnc = new ClsCNC();
        //    clsCnCWord MergedNumericWord = new clsCnCWord();
        //    clsCnCWord[] NumericWords = new clsCnCWord[30];
        //    List<string> filterEndswith = new List<string>() { ":", ";",".","*" };
        //    List<string> FilterRightWord = new List<string>() { "CARRIERS", "CARRIER","PLEASE","PACKAGES","MUST","THE","HOUR","FOR" };
        //    List<string> ignoreKeywords = new List<string>() { "MANIFEST", "MANIFEST ID" };

        //    try
        //    {
        //        int WordNo = Convert.ToInt32(FirstNumericWord.WordNumber);
        //        NumericWords[0] = FirstNumericWord;
        //        string pattern = GetPattern(FirstNumericWord.strWord);
        //        int index = 1;
        //        for (int i = WordNo + 1; i <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; i++)
        //        {
        //            string strPrevWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i - 1].strWord;
        //            string strNextWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord;
        //            int PrevWordRight = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i - 1].Right;
        //            int NextWordLeft = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].Left;
        //            //string NextwordPattern = GetPattern(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord);
        //            if (i == WordNo + 1)
        //            {
        //                string NextwordPattern = GetPattern(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord);
        //                RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(pattern, NextwordPattern);
        //                if (pattern == NextwordPattern)
        //                {
        //                    break;
        //                }
        //            }
        //            //if ((NextWordLeft - PrevWordRight <= 40) && (!strNextWord.Contains(":"))&&(strNextWord.ToUpper()!= "CARRIER") && ((!strPrevWord.EndsWith(";")) || (!strPrevWord.EndsWith(":"))))
        //            if ((NextWordLeft - PrevWordRight <= 40) && (!strNextWord.Contains(":"))&&(!FilterRightWord.Contains(strNextWord.ToUpper())) && (!filterEndswith.Any(strPrevWord.EndsWith)) && (!ignoreKeywords.Contains(keyword.ToUpper())))
        //            {
        //                if (strNextWord.All(char.IsLetter) || (Module1.regex.IsMatch(strNextWord) || Module1.regex1.IsMatch(strNextWord))|| (Module1.SpecialCharacters.Contains(strNextWord)))
        //                {
        //                    NumericWords[index] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i];
        //                    index++;
        //                }
        //                else { break; }
        //            }
        //            else
        //            { break; }
        //        }
        //        if (NumericWords[1] == null && NumericWords[2] == null)
        //        {
        //            MergedNumericWord = FirstNumericWord;
        //        }
        //        else
        //        {
        //            MergedNumericWord = Module1.oClsCnc.MergeWords(NumericWords);
        //        }
        //      //  MergedNumericWord = Module1.oClsCnc.MergeWords(NumericWords);
        //    }
        //    catch (Exception)
        //    { }
        //    finally
        //    {
        //        NumericWords = null;
        //    }
        //    return MergedNumericWord;
        //}

        private clsCnCWord FindNextWordforAlpha(clsCncMetaData ObjMetaData, clsCnCBoundingWords ObjBoundingWord, int intCurrPageNumber, int iLine, clsCnCWord FirstNumericWord)
        {
            //ClsCNC oCnc = new ClsCNC();
            clsCnCWord MergedNumericWord = new clsCnCWord();
            clsCnCWord[] NumericWords = new clsCnCWord[30];
            List<string> filterEndswith = new List<string>() { ":", ";", "*" };
            try
            {
                int WordNo = Convert.ToInt32(FirstNumericWord.WordNumber);
                NumericWords[0] = FirstNumericWord;
                string pattern = GetPattern(FirstNumericWord.strWord);
                int index = 1;
                for (int i = WordNo + 1; i <= ObjMetaData.Page[intCurrPageNumber].Line[iLine].WordCount; i++)
                {
                    if (i <= WordNo + 5)
                    {
                        string strPrevWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i - 1].strWord;
                        string strNextWord = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord;
                        int PrevWordRight = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i - 1].Right;
                        int NextWordLeft = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].Left;
                        //string NextwordPattern = GetPattern(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord);
                        if (i == WordNo + 1)
                        {
                            string NextwordPattern = GetPattern(ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i].strWord);
                            RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(pattern, NextwordPattern);
                            if (pattern == NextwordPattern)
                            {
                                break;
                            }
                        }
                        if ((NextWordLeft - PrevWordRight <= 40) && (!strNextWord.Contains(":"))&& (!filterEndswith.Any(strPrevWord.EndsWith)))
                        {
                            Boolean A = strNextWord.All(char.IsLetter);
                            //Boolean B = Module1.regex.IsMatch(strNextWord);
                            Boolean B = (!Module1.NewRegexForAlphanumeric.IsMatch(strNextWord.Replace(" ","")));
                            // Boolean C= Module1.regex1.IsMatch(strNextWord);
                            Boolean D = Module1.SpecialCharacters.Contains(strNextWord);

                            strNextWord = ReplaceStartEndSpecialchar(strNextWord);
                            //if (strNextWord.All(char.IsLetter) || (Module1.regex.IsMatch(strNextWord.Replace(" ","")) || Module1.regex1.IsMatch(strNextWord.Replace(" ", "")))|| (Module1.SpecialCharacters.Contains(strNextWord)))
                            if (strNextWord.All(char.IsLetter) || (!Module1.NewRegexForAlphanumeric.IsMatch(strNextWord.Replace(" ", "")))|| (Module1.SpecialCharacters.Contains(strNextWord)))
                            {
                                NumericWords[index] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[i];
                                index++;
                            }
                            else
                            {
                                break;
                            }
                        }
                        else
                        { break; }
                    }
                    else { break; }
                }
                MergedNumericWord = Module1.oClsCnc.MergeWords(NumericWords);
            }
            catch (Exception)
            { }
            finally
            {
                NumericWords = null;
                filterEndswith.Clear();
            }
            return MergedNumericWord;
        }
        //To compare word with barcode
        private List<clsCnCWord> CompareWord(List<clsCnCWord> PossibleWords)
        {
            List<clsCnCWord> Pwords = new List<clsCnCWord>();
            foreach (clsCnCWord StrPossibleWord in PossibleWords)
            {
                foreach (string barcode in Module1.Barcodes)
                {
                    string[] SplitBarcode = barcode.Split('~');
                    if(SplitBarcode[1].Length==Module1.ProNo.Length+1)
                    {
                        RetStructIQSBR050 obj50PRO = Module1.ObjSbr.IQSBR050(SplitBarcode[1], Module1.ProNo);
                        if (obj50PRO.CharChanged==1)
                        {
                            continue;
                        }
                    }
                    else if(SplitBarcode[1]==Module1.ProNo)
                    {
                        continue;
                    }
                    RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(StrPossibleWord.strWord, SplitBarcode[1]);
                    //if(obj50.PercentageMatch>=90)
                    if (obj50.PercentageMatch == 100)
                    {
                        StrPossibleWord.strWord = SplitBarcode[1];
                        StrPossibleWord.Flag = "1";
                        StrPossibleWord.Remarks = "1F";
                        StrPossibleWord.Confidence = 200;
                    }
                    //if (obj50.PercentageMatch >= 90)
                    ////if(StrPossibleWord.strWord== SplitBarcode[1])
                    //{
                    //    clsCnCWord ObjWord = new clsCnCWord();
                    //    ObjWord.X1Char = StrPossibleWord.X1Char;
                    //    ObjWord.Y1Char = StrPossibleWord.Y1Char;
                    //    ObjWord.X2Char = StrPossibleWord.X2Char;
                    //    ObjWord.Y2Char = StrPossibleWord.Y2Char;
                    //    ObjWord.Confidence = 200;
                    //    ObjWord.PageNo = StrPossibleWord.PageNo;
                    //    ObjWord.Left = StrPossibleWord.Left;
                    //    ObjWord.Right = StrPossibleWord.Right;
                    //    ObjWord.Top = StrPossibleWord.Top;
                    //    ObjWord.Bottom = StrPossibleWord.Bottom;
                    //    //ObjWord.strWord =string.Concat(StrPossibleWord.strWord,"@@@","1");
                    //    ObjWord.strWord = SplitBarcode[1];
                    //    ObjWord.ConfString = "9".PadLeft(barcode.Length, '9');
                    //    ObjWord.Flag = "1";
                    //    ObjWord.Remarks = "1F";
                    //    Pwords.Add(ObjWord);
                    //}
                }
            }
            if (Pwords.Count == 0)
            {
                Pwords = PossibleWords;
            }
            return Pwords;

        }

        private clsCnCWord ModifyWord(clsCnCWord oWord, string strChange)
        {
            clsCnCWord ObjWord = new clsCnCWord();
            ObjWord.X1Char = oWord.X1Char;
            ObjWord.Y1Char = oWord.Y1Char;
            ObjWord.X2Char = oWord.X2Char;
            ObjWord.Y2Char = oWord.Y2Char;
            ObjWord.Confidence = oWord.Confidence;
            ObjWord.PageNo = oWord.PageNo;
            ObjWord.LineNo = oWord.LineNo;
            ObjWord.WordNumber = oWord.WordNumber;
            ObjWord.Left = oWord.Left;
            ObjWord.Right = oWord.Right;
            ObjWord.Top = oWord.Top;
            ObjWord.Bottom = oWord.Bottom;
            ObjWord.strWord = strChange;
            ObjWord.ConfString = oWord.ConfString + "9";
            ObjWord.Flag = oWord.Flag;
            ObjWord.Remarks = oWord.Remarks;
            return ObjWord;
        }

        //To get all values from possiblewordList
        private RetStructF3 GetValues(clsCncMetaData oMeta, int pgno,int field, int iRecursive, clsCnCWord[] TopKeywords)
        {
            //RetStructF3 objRetStructF3 = new RetStructF3();
            clsCnCWord RetWord = new clsCnCWord();
            Module1.lstConfFlag.Clear();
            Module1.lstRemarks.Clear();
            DataTable dt = new DataTable();
            List<clsCnCWord> TempPossibleWordswithRA = new List<clsCnCWord>();

            try
            {
                if (Module1.PossibleWords.Count > 1)
                {
                    if (field == 1)
                    {
                        if (Module1.Barcodes != null)
                        {
                            List<clsCnCWord> NewWord = CompareWord(Module1.PossibleWords);
                            if (NewWord.Count == 1)
                            {
                                clsCnCWord NWord = NewWord[0];
                                Module1.PossibleWords.Clear();
                                Module1.PossibleWords.Add(NWord);
                            }
                        }
                    }
                    // clsCnCWord[] MultiWords = new clsCnCWord[Module1.PossibleWords.Count];

                    if (Module1.PossibleWords.Count > 1)
                    {
                        if(field==1)
                        {                            
                            List<clsCnCWord> TempPossibleWordswithoutRA = new List<clsCnCWord>();
                            foreach (clsCnCWord oWord in Module1.PossibleWords)
                            {
                                if(Module1.RAKeywords.Any(oWord.strWord.StartsWith))
                                {
                                    TempPossibleWordswithRA.Add(oWord);
                                }
                                else
                                {
                                    TempPossibleWordswithoutRA.Add(oWord);
                                }
                            }
                            if(TempPossibleWordswithRA.Count!=0)
                            {
                                foreach(clsCnCWord oWord1 in TempPossibleWordswithoutRA)
                                {
                                    TempPossibleWordswithRA.Add(oWord1);
                                }
                                Module1.PossibleWords.Clear();
                                Module1.PossibleWords = TempPossibleWordswithRA;
                            }
                            
                        }


                        dt.Columns.Add("Word");
                        dt.Columns.Add("PageNo");
                        dt.Columns.Add("Lineno");
                        dt.Columns.Add("index");
                        dt.Columns["PageNo"].DataType = Type.GetType("System.Int32");
                        dt.Columns["Lineno"].DataType = Type.GetType("System.Int32");
                        int i = 0;
                        foreach (clsCnCWord oWord in Module1.PossibleWords)
                        {
                            DataRow dr;
                            dr = dt.NewRow();
                            dr["Word"] = oWord.strWord;
                            dr["PageNo"] = oWord.PageNo;
                            dr["Lineno"] = oWord.LineNo;
                            dr["index"] = i;
                            dt.Rows.Add(dr);
                            i++;

                        }
                        DataView dv = dt.DefaultView;
                        if (field == 3)
                        {
                            //dv.Sort = "Lineno Asc";
                            dv.Sort = "PageNo, Lineno Asc";
                        }
                        //if (TempPossibleWordswithRA.Count == 0 && field==1)
                        //{
                        //    dv.Sort = "Lineno Asc";
                        //}
                        
                        //dt = dt.DefaultView.ToTable();
                        //DataTable NewDt = dt.DefaultView.ToTable();
                        DataTable NewDt = dv.ToTable();
                        List<string> sb = new List<string>();
                        //StringBuilder conf = new StringBuilder();
                        //StringBuilder Remarks = new StringBuilder();
                        for (int j = 0; j < NewDt.Rows.Count; j++)
                        {
                            string str = string.Empty;
                            if (j < NewDt.Rows.Count - 1)
                            {
                                str = NewDt.Rows[j]["Word"].ToString();
                                int countTilda = str.Count(c => (c == '~'));

                                str = string.Concat(NewDt.Rows[j]["Word"].ToString(), "~");
                                str = ReplaceStartEndSpecialchar(str);
                                if(field==2)
                                {
                                    if (str.ToUpper().StartsWith("PO") || str.ToUpper().StartsWith("PO#") || str.ToUpper().StartsWith("PO-"))
                                    {
                                        str = str.ToUpper().Replace("PO#", "");
                                        str = str.ToUpper().Replace("PO-", "");
                                        str = str.ToUpper().Replace("PO", "");
                                    }
                                }
                                //if (str.ToUpper().StartsWith("NUMBER"))
                                //{
                                //    str = str.ToUpper().Replace("NUMBER", "");
                                //}
                                //if (str.ToUpper().StartsWith("NO"))
                                //{
                                //    str = str.ToUpper().Replace("NO", "");
                                //}
                                //if (str.ToUpper().StartsWith("ID"))
                                //{
                                //    str = str.ToUpper().Replace("ID", "");
                                //}
                                //sb = sb.Append(str);
                                sb.Add(str);
                                string strToSplit = str;
                                str = str.Replace(" ", "");
                                RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(strToSplit, "E");
                                if (Obj008.Status == "S")
                                {
                                    if (strToSplit.All(char.IsLetterOrDigit))
                                    {
                                        Obj008.Status = "F";
                                    }
                                }

                                if (Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag == null)
                                {                                    
                                    if (Module1.NewRegexForAlphanumeric.IsMatch(str))
                                    {
                                        //RetWord.Flag = "0";
                                        //RetWord.Remarks = "NA";
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (str.All(char.IsUpper))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if (str.Replace(" ", "").Length < 4)
                                    else if (str.Replace(" ", "").Length <= 5)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if (str.Replace(" ", "").All(char.IsLetter))
                                    //{
                                    //    Module1.lstConfFlag.Add("0");
                                    //    Module1.lstRemarks.Add("NA");
                                    //}
                                    else if (Regex.Replace(str.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (Module1.MDYwithChar.IsMatch(str))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (Module1.Timewithoutcolon.IsMatch(str)|| Module1.Time.IsMatch(str)|| Module1.RegTel.IsMatch(str))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if ((str.Count(char.IsLetter) == 1) && (!char.IsLetter(str[0])) && (!char.IsLetter(str[str.Length - 1])))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if(Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith))
                                    else if ((Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith))||(Module1.ShouldNotContains.Any(str.ToUpper().Contains)))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (Module1.SplCharsToMakeRed.Any(str.ToUpper().Contains))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (str.Count(char.IsNumber) == 1)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (str.Count(char.IsLower) == 1)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if(strToSplit.Split().Length>2)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if(Obj008.Status=="S")
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else
                                    {
                                        Boolean ConfidenceFlag = HighestConfidence(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])],field);
                                        //    if (Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Confidence >= 85)
                                        //{
                                        if (ConfidenceFlag == true)
                                        {                                        
                                            //RetWord.Flag = "1";
                                            //RetWord.Remarks = "1F";
                                            Module1.lstConfFlag.Add("1");
                                            Module1.lstRemarks.Add("1F");
                                        }
                                        else
                                        {
                                            //RetWord.Flag = "0";
                                            //RetWord.Remarks = "NA";
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                    }
                                    //Module1.lstConfFlag.Add("0");
                                    //Module1.lstRemarks.Add("NA");
                                }
                                else
                                {
                                    if (!str.Contains("~"))
                                    {
                                        //if (str.Replace(" ", "").Length < 4)
                                        if (Module1.NewRegexForAlphanumeric.IsMatch(str))
                                        {
                                            //RetWord.Flag = "0";
                                            //RetWord.Remarks = "NA";
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.Replace(" ", "").Length <=5)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.All(char.IsUpper))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        //else if (str.Replace(" ", "").All(char.IsLetter))
                                        //{
                                        //    Module1.lstConfFlag.Add("0");
                                        //    Module1.lstRemarks.Add("NA");
                                        //}
                                        else if (Regex.Replace(str.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (Module1.MDYwithChar.IsMatch(str))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        //else if (Module1.Timewithoutcolon.IsMatch(str))
                                        else if (Module1.Timewithoutcolon.IsMatch(str) || Module1.Time.IsMatch(str) || Module1.RegTel.IsMatch(str))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if ((str.Count(char.IsLetter) == 1) && (!char.IsLetter(str[0])) && (!char.IsLetter(str[str.Length - 1])))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        //else if (Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith))
                                        else if ((Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith)) || (Module1.ShouldNotContains.Any(str.ToUpper().Contains)))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (Module1.SplCharsToMakeRed.Any(str.ToUpper().Contains))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.Count(char.IsNumber) == 1)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.Count(char.IsLower) == 1)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (strToSplit.Split().Length > 2)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (Obj008.Status == "S")
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }

                                        else
                                        {
                                            Module1.lstConfFlag.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag);
                                            Module1.lstRemarks.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Remarks);
                                        }
                                    }
                                    else
                                    {
                                        Module1.lstConfFlag.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag);
                                        Module1.lstRemarks.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Remarks);
                                    }
                                }
                            }                               
                            else
                            {
                                str = NewDt.Rows[j]["Word"].ToString();
                                int countTilda = str.Count(c => (c == '~'));
                                str = ReplaceStartEndSpecialchar(str);
                                // sb = sb.Append(str);
                                sb.Add(str);
                                string strToSplit = str;
                                RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(strToSplit, "E");
                                if (Obj008.Status == "S")
                                {
                                    if (strToSplit.All(char.IsLetterOrDigit))
                                    {
                                        Obj008.Status = "F";
                                    }
                                }
                                str = str.Replace(" ", "");
                                if (Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag == null)
                                {
                                    //if (iRecursive == 1)
                                    //{
                                    //    //RetWord.Flag = "0";
                                    //    //RetWord.Remarks = "NA";
                                    //    Module1.lstConfFlag.Add("0");
                                    //    Module1.lstRemarks.Add("NA");

                                    //}
                                    //else if (!(Module1.regex.IsMatch(str) || Module1.regex1.IsMatch(str)))
                                    if (Module1.NewRegexForAlphanumeric.IsMatch(str))
                                    {
                                        //RetWord.Flag = "0";
                                        //RetWord.Remarks = "NA";
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if(str.All(char.IsUpper))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if (str.Replace(" ", "").All(char.IsLetter))
                                    //{
                                    //    Module1.lstConfFlag.Add("0");
                                    //    Module1.lstRemarks.Add("NA");
                                    //}
                                    else if (Regex.Replace(str.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if (str.Replace(" ", "").Length < 4)
                                    else if (str.Replace(" ", "").Length <= 5)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (Module1.MDYwithChar.IsMatch(str))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if (Module1.Timewithoutcolon.IsMatch(str))
                                    else if (Module1.Timewithoutcolon.IsMatch(str) || Module1.Time.IsMatch(str) || Module1.RegTel.IsMatch(str))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if ((str.Count(char.IsLetter) == 1) && (!char.IsLetter(str[0])) && (!char.IsLetter(str[str.Length - 1])))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    //else if (Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith))
                                    else if ((Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith)) || (Module1.ShouldNotContains.Any(str.ToUpper().Contains)))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (Module1.SplCharsToMakeRed.Any(str.ToUpper().Contains))
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (str.Count(char.IsNumber) == 1)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (str.Count(char.IsLower) == 1)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (strToSplit.Split().Length > 2)
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else if (Obj008.Status == "S")
                                    {
                                        Module1.lstConfFlag.Add("0");
                                        Module1.lstRemarks.Add("NA");
                                    }
                                    else
                                    {
                                        Boolean ConfidenceFlag = HighestConfidence(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])],field);
                                        //if (Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Confidence >= 85)
                                        if(ConfidenceFlag==true)
                                        {
                                            //RetWord.Flag = "1";
                                            //RetWord.Remarks = "1F";
                                            Module1.lstConfFlag.Add("1");
                                            Module1.lstRemarks.Add("1F");
                                        }
                                        else
                                        {
                                            //RetWord.Flag = "0";
                                            //RetWord.Remarks = "NA";
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                    }
                                    //Module1.lstConfFlag.Add("0");
                                    //Module1.lstRemarks.Add("NA");
                                }
                                else
                                {
                                    if (!str.Contains("~"))
                                    {
                                        //if (str.Replace(" ", "").Length < 4)
                                        if (Module1.NewRegexForAlphanumeric.IsMatch(str))
                                        {                                            
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.Replace(" ", "").Length <= 5)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.All(char.IsUpper))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        //else if (str.Replace(" ", "").All(char.IsLetter))
                                        //{
                                        //    Module1.lstConfFlag.Add("0");
                                        //    Module1.lstRemarks.Add("NA");
                                        //}                                       
                                        else if (Regex.Replace(str.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (Module1.MDYwithChar.IsMatch(str))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        // else if (Module1.Timewithoutcolon.IsMatch(str))
                                        else if (Module1.Timewithoutcolon.IsMatch(str) || Module1.Time.IsMatch(str) || Module1.RegTel.IsMatch(str))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if ((str.Count(char.IsLetter) == 1) && (!char.IsLetter(str[0])) && (!char.IsLetter(str[str.Length - 1])))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        //else if (Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith))
                                        else if ((Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith)) || (Module1.ShouldNotContains.Any(str.ToUpper().Contains)))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (Module1.SplCharsToMakeRed.Any(str.ToUpper().Contains))
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (str.Count(char.IsNumber) == 1)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if(str.Count(char.IsLower)==1)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (strToSplit.Split().Length > 2)
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else if (Obj008.Status == "S")
                                        {
                                            Module1.lstConfFlag.Add("0");
                                            Module1.lstRemarks.Add("NA");
                                        }
                                        else
                                        {
                                            Module1.lstConfFlag.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag);
                                            Module1.lstRemarks.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Remarks);
                                        }
                                    }
                                    else
                                    {
                                        Module1.lstConfFlag.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag);
                                        Module1.lstRemarks.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Remarks);
                                    }
                                    //Module1.lstConfFlag.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Flag);
                                    //Module1.lstRemarks.Add(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[j]["index"])].Remarks);
                                }

                            }

                        }

                        string strValue = string.Join("~", sb.ToArray());
                        RetWord = ModifyWord(Module1.PossibleWords[Convert.ToInt32(NewDt.Rows[0]["index"])], strValue);
                        RetWord.Flag = string.Join("~", Module1.lstConfFlag.ToArray());
                        RetWord.Remarks = string.Join("~", Module1.lstRemarks.ToArray());

                        RetWord.Confidence = 200;
                        Module1.PossibleWords.Clear();
                        Module1.PossibleWords.Add(RetWord);
                        
                    }


                }
                else if (Module1.PossibleWords.Count == 1)
                {                    
                    if (Module1.PossibleWords[0].strWord != "*****")
                    {                        
                        // if (!(Module1.PossibleWords[0].strWord.Contains("@@@")))
                        if (Module1.PossibleWords[0].Flag == null)
                        {
                            if (field == 1)
                            {
                                if (Module1.Barcodes != null)
                                {
                                    List<clsCnCWord> NewWord = CompareWord(Module1.PossibleWords);
                                    clsCnCWord NWord = NewWord[0];
                                    Module1.PossibleWords.Clear();
                                    Module1.PossibleWords.Add(NWord);
                                    // Module1.objRetStructF3.NoOfieldsSuspects = NewWord.Count;
                                }
                            }
                            if (Module1.PossibleWords[0].Confidence != 200)
                            {
                                string str = string.Empty;
                                RetWord = Module1.PossibleWords[0];
                                RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(RetWord.strWord, "E");
                                if (Obj008.Status == "S")
                                {
                                    if (RetWord.strWord.All(char.IsLetterOrDigit))
                                    {
                                        Obj008.Status = "F";
                                    }
                                }
                                string retStrWord = RetWord.strWord.Replace(" ", "");
                                
                                //if (!(Module1.regex.IsMatch(retStrWord) || Module1.regex1.IsMatch(retStrWord)))
                                if (Module1.NewRegexForAlphanumeric.IsMatch(retStrWord))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";

                                }
                                else if (RetWord.strWord.Replace(" ", "").All(char.IsUpper))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                //else if(retStrWord.All(char.IsLetter))
                                //{
                                //    RetWord.Flag = "0";
                                //    RetWord.Remarks = "NA";
                                //}
                                else if (Regex.Replace(RetWord.strWord.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                //else if (RetWord.strWord.Replace(" ", "").Length < 4)
                                else if (RetWord.strWord.Replace(" ", "").Length <= 5)
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if(Module1.MDYwithChar.IsMatch(RetWord.strWord))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                //else if (Module1.Timewithoutcolon.IsMatch(RetWord.strWord))
                                else if (Module1.Timewithoutcolon.IsMatch(RetWord.strWord) || Module1.Time.IsMatch(RetWord.strWord) || Module1.RegTel.IsMatch(RetWord.strWord))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if ((RetWord.strWord.Count(char.IsLetter) == 1) && (!char.IsLetter(RetWord.strWord[0])) && (!char.IsLetter(RetWord.strWord[RetWord.strWord.Length - 1])))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                //else if (Module1.ShouldNotStartWith.Any(RetWord.strWord.ToUpper().StartsWith))
                                else if ((Module1.ShouldNotStartWith.Any(RetWord.strWord.ToUpper().StartsWith)) || (Module1.ShouldNotContains.Any(RetWord.strWord.ToUpper().Contains)))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if (Module1.SplCharsToMakeRed.Any(RetWord.strWord.ToUpper().Contains))
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if (RetWord.strWord.Count(char.IsNumber) == 1)
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if (RetWord.strWord.Count(char.IsLower) == 1)
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if (RetWord.strWord.Split().Length > 2)
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else if (Obj008.Status == "S")
                                {
                                    RetWord.Flag = "0";
                                    RetWord.Remarks = "NA";
                                }
                                else
                                {
                                    Boolean ConfidenceFlag = HighestConfidence(Module1.PossibleWords[0],field);
                                    //if (Module1.PossibleWords[0].Confidence >= 85)
                                    if (ConfidenceFlag == true)
                                    {
                                        RetWord.Flag = "1";
                                        RetWord.Remarks = "1F";
                                    }
                                    else
                                    {
                                        RetWord.Flag = "0";
                                        RetWord.Remarks = "NA";
                                    }
                                }
                                // RetWord = ModifyWord(Module1.PossibleWords[0], str);
                                RetWord.Confidence = 200;
                                Module1.PossibleWords.Clear();
                                Module1.PossibleWords.Add(RetWord);
                            }
                        }
                        else
                        {
                            RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(Module1.PossibleWords[0].strWord, "E");
                            if (Obj008.Status == "S")
                            {
                                if (Module1.PossibleWords[0].strWord.All(char.IsLetterOrDigit))
                                {
                                    Obj008.Status = "F";
                                }
                            }
                            if (!Module1.PossibleWords[0].strWord.Contains("~"))
                            {
                                if (Module1.PossibleWords[0].strWord.Replace(" ", "").All(char.IsUpper))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                //else if (Module1.PossibleWords[0].strWord.All(char.IsLetter))
                                //{
                                //    Module1.PossibleWords[0].Flag = "0";
                                //    Module1.PossibleWords[0].Remarks = "NA";
                                //}
                                else if (Regex.Replace(Module1.PossibleWords[0].strWord.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                //else if (Module1.PossibleWords[0].strWord.Replace(" ", "").Length < 4)
                                else if (Module1.PossibleWords[0].strWord.Replace(" ", "").Length <= 5)
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                else if (Module1.MDYwithChar.IsMatch(Module1.PossibleWords[0].strWord))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                //else if (Module1.Timewithoutcolon.IsMatch(Module1.PossibleWords[0].strWord))
                                else if (Module1.Timewithoutcolon.IsMatch(Module1.PossibleWords[0].strWord) || Module1.Time.IsMatch(Module1.PossibleWords[0].strWord) || Module1.RegTel.IsMatch(Module1.PossibleWords[0].strWord))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }                                
                                else if ((Module1.PossibleWords[0].strWord.Count(char.IsLetter) == 1) && (!char.IsLetter(Module1.PossibleWords[0].strWord[0])) && (!char.IsLetter(Module1.PossibleWords[0].strWord[Module1.PossibleWords[0].strWord.Length - 1])))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                // else if (Module1.ShouldNotStartWith.Any(Module1.PossibleWords[0].strWord.ToUpper().StartsWith))
                                else if ((Module1.ShouldNotStartWith.Any(Module1.PossibleWords[0].strWord.ToUpper().StartsWith)) || (Module1.ShouldNotContains.Any(Module1.PossibleWords[0].strWord.ToUpper().Contains)))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                else if (Module1.SplCharsToMakeRed.Any(Module1.PossibleWords[0].strWord.ToUpper().Contains))
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                else if (Module1.PossibleWords[0].strWord.Count(char.IsNumber) == 1)
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                else if (Module1.PossibleWords[0].strWord.Count(char.IsLower) == 1)
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                else if (Module1.PossibleWords[0].strWord.Split().Length > 2)
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                                else if (Obj008.Status == "S")
                                {
                                    Module1.PossibleWords[0].Flag = "0";
                                    Module1.PossibleWords[0].Remarks = "NA";
                                }
                            }
                        }
                    }
                }
                Module1.objRetStructF3.Words = Module1.PossibleWords;
                Module1.objRetStructF3.NoOfieldsSuspects = Module1.PossibleWords.Count;
                Module1.conflevel.Add(90);
                Module1.objRetStructF3.ConfidenceLevelofSuspect = Module1.conflevel;
                Module1.objRetStructF3.Flag = "Y";
                Module1.objRetStructF3.ManualConfirmation = "N";
                Module1.objRetStructF3.Status = "S";
            }
            catch (Exception ex)
            { }
            finally
            {
                dt.Dispose();
                TempPossibleWordswithRA=null;
            }

            return Module1.objRetStructF3;
        }

        private string ReplaceStartEndSpecialchar(string strWord)
        {
            String WordAlphanumeric = Regex.Replace(strWord, @"[^0-9a-zA-Z]+", "");
            if ((!string.IsNullOrEmpty(WordAlphanumeric))&&(!Module1.SpecialCharacters.Contains(strWord)))
            {
                strWord = strWord.Replace("'", "");
            LabelStart: if (Module1.SpecialCharacters.Contains(strWord[0].ToString()))
                {
                    strWord = strWord.Remove(0, 1);
                    goto LabelStart;
                }
            LabelEnd: if (Module1.SpecialCharacters.Any(strWord.EndsWith))
                {
                    strWord = strWord.Remove(strWord.Length - 1, 1);
                    goto LabelEnd;
                }
            }
            return strWord;
        }

        //to check duplicate value for Bottom Multiple values
        private Boolean CheckExistWordforBottomMulti(clsCnCWord[] BottomMulti, clsCnCWord NewWord)
        {
            Boolean exist = false;
            try
            {
                foreach (clsCnCWord Word in BottomMulti)
                {
                    if (Word != null && Word.strWord != null)
                    {
                        string StrWord = Word.strWord.Trim();
                        if (Module1.SpecialCharacters.Contains(StrWord[0].ToString()))
                        {
                            StrWord = StrWord.Remove(0, 1);
                        }
                        if (Module1.SpecialCharacters.Any(StrWord.EndsWith))
                        {
                            StrWord = StrWord.Remove(StrWord.Length - 1, 1);
                        }
                        string StrNewWord = NewWord.strWord.Trim();
                        if (Module1.SpecialCharacters.Contains(StrNewWord[0].ToString()))
                        {
                            StrNewWord = StrNewWord.Remove(0, 1);
                        }
                        if (Module1.SpecialCharacters.Any(StrNewWord.EndsWith))
                        {
                            StrNewWord = StrNewWord.Remove(StrNewWord.Length - 1, 1);
                        }
                        if (StrWord.Equals(StrNewWord))
                        {
                            exist = true;
                        }
                        else
                        {
                            if (exist != true)
                            {
                                exist = false;
                            }
                        }
                    }
                    else { break; }
                }
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
            }
            return exist;

        }
        #endregion

        private void ContainsKeyword(clsCncMetaData ObjMetaData, int intCurrPageNumber, clsCnCWord[] TopKeywords, clsCnCWord Topword, int field, int iRecursiveCallCount, string keyword,int keywordLength)
        {
            Boolean FilterwordsPO = false;
            string[] FilterKeywords = { "BOX", "OFFICE", "TYPE", "BOX#" };
            
            clsCnCWord KeywordwithValue = new clsCnCWord();
            KeywordwithValue = Topword;
            string word;
            if (keyword.ToUpper()=="PO"&& Topword.strWord.ToUpper().StartsWith("PO"))
            {
                word = Topword.strWord.ToUpper().Remove(0, 2);
            }
            else if (TopKeywords[keywordLength - 1].strWord.Contains(":"))
            {
                string[] splitword = TopKeywords[keywordLength - 1].strWord.Split(':');
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
               // word = ReplaceStartEndSpecialchar(word);
                if (!Module1.SpecialCharacters.Contains(word))
                {
                    word = word.Replace("'", "");
                LabelStart: if (Module1.SpecialCharacters.Contains(word[0].ToString()))
                    {
                        word = word.Remove(0, 1);
                        goto LabelStart;
                    }
                LabelEnd: if (Module1.SpecialCharacters.Any(word.EndsWith)&&(!word.EndsWith(",")))
                    {
                        word = word.Remove(word.Length - 1, 1);
                        goto LabelEnd;
                    }
                }
                //if (Module1.SpecialCharacters.Contains(word[0].ToString()) && (!word.StartsWith("(")))
                //{
                //    word = word.Remove(0, 1);
                //}
                //if (Module1.SpecialCharacters.Any(word.EndsWith) && (!word.EndsWith(")")))
                //{
                //    word = word.Remove(word.Length - 1, 1);
                //}                
            }
            RetStructIQSBR008 Obj008 = Module1.ObjSbr.IQSBR008(word, "E");
            if (Obj008.Status == "S")
            {
                if (word.All(char.IsLetterOrDigit))
                {
                    Obj008.Status = "F";
                }
            }
            if ((!Module1.SpecialCharacters.Contains(word)))
            {
                if (!Module1.Time.IsMatch(word))
                {
                    WordToChckLength = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                }
            }
            if (field == 2)
            {
                if (FilterKeywords.Any(word.ToUpper().StartsWith))
                {
                    FilterwordsPO = true;
                }
            }
            if (FilterwordsPO == false)
            {
                //if (!string.IsNullOrEmpty(word) && (Module1.regex.IsMatch(word) || Module1.regex1.IsMatch(word)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word != Module1.ProNo && WordToChckLength.Length >= 2)
                if (!string.IsNullOrEmpty(word) && (!(Module1.NewRegexForAlphanumeric.IsMatch(word))) && (!word.All(char.IsLetter)) && (!Module1.Time.IsMatch(word)) && (!Module1.RegTel.IsMatch(word)) && Obj008.Status == "F" && word != Module1.ProNo && WordToChckLength.Length >= 2)
                {
                    clsCnCWord BOL = Topword;
                    if (Module1.SpecialCharacters.Contains(word[0].ToString()))
                    {
                        word = word.Remove(0, 1);
                        
                    }
                    Boolean confidence = HighConfidenceforMergedWords(KeywordwithValue, word,field);
                    BOL.strWord = word;
                    clsCnCWord CncPoWord = new clsCnCWord();
                    //if (word.Count(char.IsLetter) <word.Count(char.IsNumber))
                    if (word.Count(char.IsNumber)>1)
                    {
                        word = Regex.Replace(word, @"[^0-9a-zA-Z]+", "");
                        if (word.Length >= 2)
                        {
                            CncPoWord = BOL;
                            string strCncWord = BOL.strWord;
                            clsCnCBoundingWords Newbounding = Module1.oClsCnc.GetBoundingWords(CncPoWord, ObjMetaData, intCurrPageNumber);
                            int CommaCount = CncPoWord.strWord.Count(c => (c == ','));
                            if (WordToChckLength.Length >= 4 && CncPoWord.strWord.EndsWith(",") && CommaCount == 1 && Newbounding != null && Newbounding.RightWord != null)
                            {
                                CncPoWord.LineNo = TopKeywords[keywordLength - 1].LineNo;
                                CncPoWord.WordNumber = TopKeywords[keywordLength - 1].WordNumber;
                                CncPoWord = RightMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord.LineNo, CncPoWord, ",", iRecursiveCallCount,field);
                            }
                           else if ((strCncWord.Contains(",") || strCncWord.Contains("-") || strCncWord.Contains("/") || strCncWord.Contains(";")|| strCncWord.Contains("&")) &&(word.Length >= 4))  // 07/12/2020
                            {
                                if (CncPoWord.strWord.Contains(","))
                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ",", iRecursiveCallCount); }
                                else if (CncPoWord.strWord.Contains("&"))
                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "&", iRecursiveCallCount); }
                                else if (CncPoWord.strWord.Contains("/"))
                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "/", iRecursiveCallCount); }                                
                                else if (CncPoWord.strWord.Contains(";"))
                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, ";", iRecursiveCallCount); }                                
                                else if (CncPoWord.strWord.Contains("-"))
                                { CncPoWord = ContainsSeparatorwithMultiValue(ObjMetaData, intCurrPageNumber, CncPoWord, "-", iRecursiveCallCount); }
                                //if (Convert.ToInt32(CncPoWord.WordNumber) != 0)
                                //{
                                //    CncPoWord = FindNextWord(ObjMetaData, ObjBoundingWord, intCurrPageNumber, iLine, CncPoWord);
                                //}
                                CncPoWord.LineNo = TopKeywords[keywordLength - 1].LineNo;
                                CncPoWord.WordNumber = TopKeywords[keywordLength - 1].WordNumber;
                            }
                            else
                            {
                                CncPoWord.WordNumber = TopKeywords[keywordLength-1].WordNumber;
                                clsCnCWord MergedCncPoWord = FindNextWord(ObjMetaData, Newbounding, intCurrPageNumber, CncPoWord.LineNo, CncPoWord,keyword);
                                //if (Module1.regex.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")) || Module1.regex1.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                if (!Module1.NewRegexForAlphanumeric.IsMatch(MergedCncPoWord.strWord.Replace(" ", "")))
                                {
                                    CncPoWord = MergedCncPoWord;
                                }
                            }

                            if (Module1.SpecialCharacters.Contains(CncPoWord.strWord[0].ToString()) && (!CncPoWord.strWord.StartsWith("(")))
                            {
                                CncPoWord = ModifyWord(CncPoWord, CncPoWord.strWord.Remove(0, 1));
                            }

                            string StrKeyword = Regex.Replace(keyword, @"[^0-9a-zA-Z]+", "");
                            if (Module1.RAKeywords.Contains(keyword.ToUpper()))    //give prefix for RA,RGA,RMA
                            {
                                CncPoWord = AddPrefix(CncPoWord, keyword);
                            }
                            CncPoWord.strWord = ReplaceValue(CncPoWord.strWord, field);
                            Boolean exist = CheckforExistWord(Module1.PossibleWords, CncPoWord);
                            if (exist == false)
                            {
                                if (CncPoWord.Flag == null)
                                {
                                    if (confidence == true)
                                    {
                                        CncPoWord.Flag = "1";
                                        CncPoWord.Remarks = "1F";
                                    }
                                    else
                                    {
                                        CncPoWord.Flag = "0";
                                        CncPoWord.Remarks = "NA";
                                    }
                                }
                                Module1.PossibleWords.Add(CncPoWord);
                                //Boolean confidence = HighConfidence(KeywordwithValue, CncPoWord.strWord);
                               
                            }

                        }

                    }
                   
                }
            }
        }

        private clsCnCWord AddPrefix(clsCnCWord WordToModify,string Keyword)
        {
            // string strKeyword =Regex.Replace( Keyword, @"[^0-9a-zA-Z]+", "");
            string strKeyword = Keyword;
            if (!Module1.RAKeywords.Any(WordToModify.strWord.StartsWith))
            {
                if (strKeyword.ToUpper() == "RA#" || strKeyword.ToUpper() == "RA" || strKeyword.ToUpper() == "RETURN AUTHORIZATION" || strKeyword.ToUpper() == "RETURNAUTHORIZATION")
                {
                    if (strKeyword.ToUpper() == "RA#")
                    {
                        WordToModify.strWord = string.Concat("RA# ", WordToModify.strWord);
                    }
                    else
                    {
                        WordToModify.strWord = string.Concat("RA ", WordToModify.strWord);
                    }
                }
                else if (strKeyword.ToUpper() == "RMA" || strKeyword.ToUpper() == "RMA#")
                {
                    if (strKeyword.ToUpper() == "RMA#")
                    {
                        WordToModify.strWord = string.Concat("RMA# ", WordToModify.strWord);
                    }
                    else
                    {
                        WordToModify.strWord = string.Concat("RMA ", WordToModify.strWord);
                    }
                }
                else if (strKeyword.ToUpper() == "RGA" || strKeyword.ToUpper() == "RGA#")
                {
                    if (strKeyword.ToUpper() == "RGA#")
                    {
                        WordToModify.strWord = string.Concat("RGA# ", WordToModify.strWord);
                    }
                    else
                    {
                        WordToModify.strWord = string.Concat("RGA ", WordToModify.strWord);
                    }
                }
            }
            return WordToModify;
        }

        private string ReplaceValue(string str,int field)
        {
            str = str.Replace("'", "");
            str = str.Replace(":", "");
          Label1:  str = ReplaceStartEndSpecialchar(str);
            //if (str.ToUpper().StartsWith("NUMBER"))
            //{
            //    str = str.ToUpper().Replace("NUMBER", "");
            //    goto Label1;
            //}
            //if (str.ToUpper().StartsWith("NO"))
            //{
            //    str = str.ToUpper().Replace("NO", "");
            //    goto Label1;
            //}
            //if (str.ToUpper().StartsWith("ID"))
            //{
            //    str = str.ToUpper().Replace("ID", "");
            //}
            if (str.ToUpper().EndsWith("CARRIERS"))
            {
                str = str.ToUpper().Replace("CARRIERS", "");
                goto Label1;
            }
            if (str.ToUpper().EndsWith("CARRIER"))
            {
                str = str.ToUpper().Replace("CARRIER", "");
                goto Label1;
            }
            if(str.ToUpper().EndsWith("CUSTOMER"))
            {
                str = str.ToUpper().Replace("CUSTOMER", "");
                goto Label1;
            }
            if (str.ToUpper().EndsWith("PICK"))
            {
                str = str.ToUpper().Replace("PICK", "");
                goto Label1;
            }
            if (str.ToUpper().StartsWith("NURNBER"))
            {
                str = str.ToUpper().Replace("NURNBER", "");
                goto Label1;
            }
            if (field == 2)
            {
                if (str.ToUpper().StartsWith("PO") || str.ToUpper().StartsWith("PO#") || str.ToUpper().StartsWith("PO-"))
                {
                    str = str.ToUpper().Replace("PO#", "");
                    str = str.ToUpper().Replace("PO-", "");
                    str = str.ToUpper().Replace("PO", "");
                    goto Label1;
                }
            }

            return str;
        }

        private clsCnCWord RetWordforAlpha(List<clsCnCWord> AlphaList)
        {
            //clsCnCWord oWord = new clsCnCWord();
            string Alpha = AlphaList[0].strWord.Replace(" ", "");
            Alpha = Regex.Replace(Alpha, @"[^0-9a-zA-Z]+", "");
            if (Alpha.All(char.IsUpper))
            {
                if(Module1.regex.IsMatch(AlphaList[1].strWord.Replace(" ","")) || Module1.regex1.IsMatch(AlphaList[1].strWord.Replace(" ", "")))
                {
                    return AlphaList[1];
                }
                else if(AlphaList[1].strWord.Contains("~"))
                {
                    return AlphaList[1];
                }
            }
            else if(Alpha.Any(char.IsLower))
            {
                if (Module1.regex.IsMatch(AlphaList[1].strWord.Replace(" ","")) || Module1.regex1.IsMatch(AlphaList[1].strWord.Replace(" ", "")))
                {
                    return AlphaList[1];
                }
            }
            AlphaList[0].Flag = "0";
            AlphaList[0].Remarks = "NA";
            return AlphaList[0];
        }

        private clsCnCWord SecondaryKeywordFlag(clsCncMetaData ObjMetaData, int intCurrPageNumber,clsCnCWord[] TopKeywords, clsCnCWord CncPoWord,int field)
        {
            clsCnCWord TopWord = Module1.oClsCnc.MergeWords(TopKeywords);
            string[] splitKeyword = TopWord.strWord.Split();
            clsCnCBoundingWords BoundingwordforSecKeyWord = Module1.oClsCnc.GetBoundingWords(TopKeywords[splitKeyword.Length - 1], ObjMetaData, intCurrPageNumber);
            if ((TopWord.strWord.EndsWith(":")) || (BoundingwordforSecKeyWord.RightWord != null && BoundingwordforSecKeyWord.RightWord.strWord == ":"))
            {
                Boolean ConfidenceFlag = HighestConfidence(CncPoWord,field);
                //if (Module1.PossibleWords[0].Confidence >= 85)
                if (ConfidenceFlag == true)
                {
                //    if (CncPoWord.Confidence >= 85)
                //{
                    CncPoWord.Flag = "1";
                    CncPoWord.Remarks = "1F";
                }
                else
                {
                    CncPoWord.Flag = "0";
                    CncPoWord.Remarks = "NA";
                }
            }
            //else
            //{
            //    CncPoWord.Flag = "0";
            //    CncPoWord.Remarks = "NA";
            //}
            return CncPoWord;
        }

        private Boolean HighConfidenceforMergedWords(clsCnCWord Currentword, string str,int field)
        {
            Boolean flagHighestConfidence = true;
            string confstring = Currentword.ConfString;
            int index = Currentword.strWord.ToUpper().IndexOf(str.ToUpper());
            if (index == -1)  //If not found index then takes confidence of whole value with keyword..this will work only for single merged keywords with value, not for multivalues.
            {
                flagHighestConfidence = HighestConfidence(Currentword,field);
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


        //---------Confidence 7#80  ----------------------
        //private Boolean HighestConfidence(clsCnCWord CurrentWord, int field)
        //{
        //    if (field == 1)
        //    {
        //        var NineDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
        //        var EightDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
        //        var ZeroDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
        //        if (CurrentWord.Confidence >= 85 && (CurrentWord.ConfString.Length == NineDigitcount + EightDigitcount + ZeroDigitcount))
        //        {
        //            return true;
        //        }
        //        else
        //        {
        //            return false;                    
        //        }
        //    }
        //    else
        //    {
        //        var NineDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '9').ToArray().Length;
        //        var EightDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '8').ToArray().Length;
        //        var SevenDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '7').ToArray().Length;
        //        var ZeroDigitcount = CurrentWord.ConfString.Select(c => c).Where(c => c == '0').ToArray().Length;
        //        if (CurrentWord.Confidence >= 80 && (CurrentWord.ConfString.Length == NineDigitcount + EightDigitcount + SevenDigitcount + ZeroDigitcount))
        //        {
        //            return true;
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //    }
        //}


        //---------Confidence 8#85  ----------------------
        private Boolean HighestConfidence(clsCnCWord CurrentWord,int field)
        {
            //clsCnCWord objShipperNameOCR = new clsCnCWord();
            ////objAdd1 = (clsCnCWord)Module1.TableRecordssFinal.Rows[0]["cncName"];
            //objShipperNameOCR = (clsCnCWord)dataRowsShipper[0]["cncName"];
            //PossibleWords.Add(objShipperNameOCR);
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

        private Boolean HighestConfidenceforHazTel(clsCnCWord CurrentWord)
        {
            //clsCnCWord objShipperNameOCR = new clsCnCWord();
            ////objAdd1 = (clsCnCWord)Module1.TableRecordssFinal.Rows[0]["cncName"];
            //objShipperNameOCR = (clsCnCWord)dataRowsShipper[0]["cncName"];
            //PossibleWords.Add(objShipperNameOCR);
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

        private DataRow[] sortFoundRowsonKeywordPreference(string sKeyword,DataRow[] Foundrows,DataTable dt)
        {
            DataTable Newdt = new DataTable();
            try
            {                
                Newdt = dt.Clone();
                Newdt = Foundrows.CopyToDataTable();
                List<string> SplitString = sKeyword.ToUpper().Replace("'", "").Split(',').ToList();
                var query = from b in SplitString
                            join r in Newdt.AsEnumerable() on b.ToUpper().Trim() equals r.Field<String>("Keyword").Trim()
                            select r;

                var orderedTable = query.CopyToDataTable();
                Foundrows = orderedTable.Select();                
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                //throw;
            }
            finally
            {
                Newdt.Dispose();
                Newdt = null;                
            }
            return Foundrows;
        }

        private Boolean FlagForMultiValues(string str)
        {
            str = ReplaceStartEndSpecialchar(str);
            if (Module1.NewRegexForAlphanumeric.IsMatch(str))
            {
                return false; 
            }
            else if (str.Replace(" ", "").All(char.IsUpper))
            {
                return false;
            }            
            else if (Regex.Replace(str.Replace(" ", ""), @"[^0-9a-zA-Z]+", "").All(char.IsLetter))
            {
                return false;
            }            
            else if (str.Replace(" ", "").Length <= 5)
            {
                return false;
            }
            else if (Module1.MDYwithChar.IsMatch(str))
            {
                return false;
            }
            
            else if (Module1.Timewithoutcolon.IsMatch(str) || Module1.Time.IsMatch(str) || Module1.RegTel.IsMatch(str))
            {
                return false;
            }
            else if ((str.Count(char.IsLetter) == 1) && (!char.IsLetter(str[0])) && (!char.IsLetter(str[str.Length - 1])))
            {
                return false;
            }
            else if (Module1.ShouldNotStartWith.Any(str.ToUpper().StartsWith))
            {
                return false;
            }
            else if (Module1.SplCharsToMakeRed.Any(str.ToUpper().Contains))
            {
                return false;
            }
            else if(str.Count(char.IsLower)==1)
            {
                return false;
            }
            return true;
        }
    }

        //private clsCnCWord RightMultiValue(clsCncMetaData ObjMetaData, int intCurrPageNumber, int iLine, clsCnCWord FirstWord)
        //{
        //    ClsCNC oCnc = new ClsCNC();
        //    clsCnCWord MultiWord = new clsCnCWord();
        //    clsCnCWord RetWord = new clsCnCWord();
        //    int WordNo = Convert.ToInt32(FirstWord.WordNumber);
        //    clsCnCWord[] Word = new clsCnCWord[20];
        //    int counter = 0;
        //    string CurrentLine = ObjMetaData.Page[intCurrPageNumber].Line[iLine].strLine;
        //    string[] splitWords = CurrentLine.Split(',');
        //    string strFirstWord = FirstWord.strWord.Replace(",", "");
        //    string pattern = GetPattern(strFirstWord);
        //    int index = 1, NxtWordCount = 1;
        //    //string[] ConfFlag = new string[20];
        //    List<string> lstConfFlag = new List<string>();
        //    Word[0] = FirstWord;
        //    if (FirstWord.Confidence >= 90)
        //    {
        //        lstConfFlag.Add("1");
        //    }
        //    else
        //    {
        //        lstConfFlag.Add("0");
        //    }

        //    for (int i = 0; i < splitWords.Length; i++)
        //    {
        //        if (counter == 0)
        //        {
        //            if (splitWords[i].Contains(strFirstWord))
        //            {
        //                counter = 1;
        //            }
        //        }
        //        else
        //        {
        //            string NextWordPattern = GetPattern(splitWords[i]);
        //            RetStructIQSBR050 obj50 = Module1.ObjSbr.IQSBR050(pattern, NextWordPattern);
        //            //if (pattern == NextWordPattern)
        //            if (obj50.PercentageMatch >= 90)
        //            {
        //                Word[index] = ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[WordNo + NxtWordCount];
        //                if (ObjMetaData.Page[intCurrPageNumber].Line[iLine].Word[WordNo + NxtWordCount].Confidence >= 90)
        //                {
        //                    lstConfFlag.Add("1");
        //                }
        //                else
        //                {
        //                    lstConfFlag.Add("0");
        //                }
        //                index++; NxtWordCount++;
        //            }
        //            else { break; }
        //        }

        //    }
        //    MultiWord = oCnc.MergeWords(Word);
        //    string Conf = string.Join("~", lstConfFlag.ToArray());
        //    String RetString = string.Concat(MultiWord.strWord.Remove(MultiWord.strWord.Length - 1), "@@@", Conf);
        //    RetWord = ModifyWord(MultiWord, RetString);
        //    return RetWord;
        //    //return MultiWord;
        //}


        //#region Supporting Functions
        //public static bool IsDate(Object obj)
        //{
        //    string strDate = obj.ToString();
        //    DateTime dtout;
        //    try
        //    {
        //        if (DateTime.TryParse(strDate, out dtout))

        //        {
        //            DateTime dt = DateTime.Parse(strDate);
        //            if ((dt.Month != System.DateTime.Now.Month) || (dt.Day < 1 && dt.Day > 31) || dt.Year != System.DateTime.Now.Year)
        //                return false;
        //            else
        //                return true;
        //        }
        //        else
        //        {
        //            return false;
        //        }
        //    }
        //    catch
        //    {
        //        return false;
        //    }
        //}

        //private DataTable F3_SearchKeyword(clsCncMetaData oMeta, int intCurrPageNo)
        //{
        //    var dtKeywordInfo = new DataTable();
        //    dtKeywordInfo = MakeDt();
        //    try
        //    {
        //        // -----------------------------------------BILL OF LADING NUMBER/PURCHASE ORDER NUMBER/SHIPPER NUMBER KEYWORD------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice no.", "Invoice #", "invoice Nurnber", "Document No", "Invoice no", "INVOIC E NO:", "No.", "NUMBER", "HP Sales order:", "Sales order :", "Lieferung :", "Shipper No:", "Delivery Note No.", "Delvery Note No.", "Lieferung Nummer:", "Lieferschein", "Your Reference:", "Ihre Bestellnr.", "bestellung", "Kundenauftragsnummer", "Ref. :", "Your order no. :", "Ihr Ansprechpartner :", "Rechnungsempfänger:", "Your Contact", "Lieferschein:", "IBAN :", "IBAN No:", "IBAN NO TL :", "IBAN code Austria", "MOMSNR.", "IBAN-Code =", "IBAN :", "IBAN", "IBAN:", "N° FACT.", "N° commande d`achat", "Ref Commande :", "Werkplaatsordernr", "Uw-bestelnummer", "Uw bestelling:", "Bestelhon", "Bestelbon", "Ihre Bestellnummer :", "Auftragsnr. Kunde", "Orderref.", "Ihr Auftrag", "Your orderno", "Deres reference", "Uw bestelbon :", "Commande de client", "Contract number:", "Référence", "ihre Referenznummer:", "Uw Ref. :", "Uw Ref.:", "Your Order", "Ihre Referenz:", "Ordre nummer :", "Work Order No :", "Order No.", "Bestellnummer:", "Order Numb", "ORDER No", "Bilagsnr. / Side :", "Nr.documento", "N° FACTURE", "FACTURE :", "Nr fv:", "Doc. No./Date", "Fattura #", "FACTURE N° :", "Numero della fattura:", "Beleg Nr;", "Document No.", "Number/Date", "Fatture N.", "Account Number", "INVolCE", "lnvoice #", "InvoiceNumber:", "Invoice#", "N. DOCUMENTO/DOCUMENTNO", "INVOICE", "Documento nr.:", "NUMERO DOCUMENTO", "Fattura n:", "Customer invoice", "Beleg:", "N° Fattura", "Nr. Documento", "Fattura nr", "N°Documento", "invoice #", "Invoice #", "invoice No:", "INVOICE NR.", "PF. INV. NO.:", "Fattura nr .", "Dokument Nr.", "Rechnungsnummer", "Beleg Nr.", "Invoice :", "invoice No", "invoice Number", "Invoice number :", "N° de facture :", "N° FA.", "Belegnr.:", "Numéro de facture", "Invoice No :", "InvoiceNO.", "Order No:", "Belegnummer", "Invoice Number:", "N. doc.", "Num doc / Date", "Invoice no.", "Fakturanummer", "Faktura", "Factura N°:", "INVOICE NR :", "Rechnung - Nr", "Fakt nr / Kundnr", "Faktura Nr.", "invoice No./Date", "Rechnungsnummer :", "Nummer / Datum", "Rechnung", "Rechnungs-Nr.:", "BELEG-NR.", "Rechnung Nr.", "Rechnungsnr.", "Rechnung:", "Rechnungsnummer:", "Numéro de facture:", "N° de facture", "Beleg-Nr. :", "FACTURE GLOBALE N°", "N° de la facture:", "FACTURE N°", "FACTURE (C) NO", "FACTURE  N':", "No. de la fact.", "Facture:", "Faktura VAT Nr", "Faktura VAT", "Nr Faktury", "Faktura nr", "Faktura", "Numer faktury :", "Numer faktury:", "FAKTURA / INVOICE:", "Rechn.Nr", "INVOICE N°", "Factuur :", "FAKTUUR", "Factuur", "Factuur nr :", "RECHNUNGS-NR.", "Factuurnr.", "FACTUURNUMMER", "factuurnr.", "Factuur:", "RECHNUNG :", "FACTUURNUMMER:", "Factuur nr.", "Fattura nr.", "Fattura -", "FATTURA N°", "Fatt. n°", "Fattura N.", "N. FATTURA", "Fakturanummer :", "Fakturanr.", "Faktura nr.:", "Faktura nummer", "Faktura nr. :", "Fakturanr./Invoice no.", "Faktura nr.", "Your purchase order :", "Your order no.", "Ihre Bestell-Nr. :", "Your order no.:", "Inköpsordernummer", "Ert ordernr", "Er referens", "Best:", "Ert ordernr", "Ert bestar", "Er referens/bestnr", "Ihre  Bestellnummer :", "ihre Bestell-Nr.", "IHRE AUFTRAGSNR :", "ihre Bestellung", "Ihre Bestell-Nr.:", "Kdnbestellnr. /cust. order no/n°de comm. du cl ent:", "Ihre Auftraos-Nr", "Bestelldaten: Nr.", "Auftragsdaten:", "Bestellung Nr.", "lhre Bestellung per FAX", "Externe Belegnummer", "Bestellung Nr.:", "Bestell Nr.", "BESTELNR.", "Uw ordernummer", "Uw Internet bestelling", "Referentienr.*", "IHRE BESTELLUNG :", "Bestelbon", "Uw referentie :", "Uw Bestelnummer", "Uw Order en ref.:", "Uw order nummer", "Vostro ordine nr.", "Vostro ordine nr.", "Votre no. de cmde:", "Deres reference :", "Deres reference rekv.nr.", "lhr Zeichen", "Bestell-Ref.", "Ihre  Bestellung:", "lhre Bestelinr.", "Ihre Referenz :", "Telefon", "fon", "Phone :", "Contact Telephone", "Tel.:", "Telefon:", "Tel.", "Telefon :", "Tel:", "Tel :", "Telefoon", "Phone:", "Sales Phone nr:", "Phone n° :", "Direct telefoonnr.  :", "doorkiesnummer", "Telephone;", "Telefon", "Telephone:", "Telefoonnummer :", "Telefoonnummer", "tlf", "Telephone No", "Téléphone", "Num. de téléphone", "Tfn:", "Tél. :", "Tél.", "Téléphone :", "Tél :", "Phone", "Téléphone:", "Pbone n° :", "Tél", "NIP", "PART.IVA", "P.iva", "Partita IVA", "IVA", "FISCALE:", "Ust.-IdNr.:", "USt.-ID-Nr.", "T.V.A", "TVA", "CVR nr.", "CVR-nr.", "CVR/SE nr.:", "CVR:", "TVA:", "B.T.W. NUMMER", "BTW nr.:", "unsere", "NIP:", "MOMSNR.", "Ust-Id ATU:", "UST-ID Nr./St.Nr", "USt-IdNr.", "Ihre-UID:", "IVA", "CVR nr.:", "ST-Id-Nr.:", "MOMSNR.", "Vostra P.IVA", "TVA", "Votre N° de TVA", "CVR-nr.:", "Customer VAT N°:", "VAT nr.:", "Momsreg.nr/VATnr:", "Momsreg.nr.", "VAT Nr/VAT No", "Momsreg.nr/VAT-nr:", "VAT no.", "Momsreg nr", "Vertr.Nr.", "Btw-nr.", "BTW:", "UST-ID NrJSt.Nr", "UST-ID Nr./St.Nr", "BTW N°", "BTW-nummer", "BTW", "VAT REG. NO.:", "Credit Invoice", "credit nota", "credit note freight", "credit note", "kreditnota nr.:", "DeblKred.", "entgeltminderung", "NOTA DEBITO", "Credit Number", "Credit memo Number", "Gutschrift", "Avoir", "Kredit nota/faktura", "nota de credito/abono", "nota de credit", "Buchunggutschift", "Stornorechnung", "Retourengutschrift", "Gutschein", "Korrekturrechnung", "Rechnungsstorno", "Abono", "Recticativa", "Warengutchrift", "Storno", "GUTSCHRIFT", "No:", "FATTURA", "Numero"}
        //        var sKeyarray = new string[] { "Bill of Lading Number", "BOL #", "B/L #", "Manifest #", "Bill of Lading No", "Shipper's bill of loading no.", "BOL#", "BOL No","Bill of L ading No","Bill of Lading", "BI LL 0 F LADING",
        //                                       "Shipper's Bill of Ladias No","BOL Number","Bill of Lading #","Bill of LadingNumber","B/L NO.","Purchase Order No","Purchase Order Number","Po#","Po #", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No",
        //                                       "Customer's PO","cust po#","purchase order","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","P.O. #","Consisnee's Refer.noe/po No","P.O. No","Purchase Order #","Cust P.O. No",
        //                                       "CUSTOMERORDERNUMBER","Order", "Order #","Shipper Number",  "Ship ID#", "Shipper Ref #", "SH PPER NUMBER", "SHIPPER'S NUMBER", "Shippers No", "Shipper No", "Shipment #", "Shipper's No","Shipper #", "Shipping Order No", "SID#", "Shipper No/Customer",
        //                                       "Shipment ID","Shipment No","Shipment number","SID","SHIPPER","No/Customer PO#","PO No.(s)","BOL NBR","SHIPPER BILL OF LADING NUMBER:","B/L Document No."};

        //        // Dim sKeyarray() As String = {"Your Contact"}
        //        // -----------------------------------------------------------------------------------------------------------------------------------
        //        // Dim sKeyarray() As String = {"Invoice No:"}

        //        Array.Sort(sKeyarray);
        //        var CNC = new ClsCNC();
        //        // Dim Words As clsCnCWord() = CNC.GetDataForROI(oMeta.Page, iPage, 0, 0, 9999, 9999)

        //        var PossibleInvoices = new List<clsCnCWord>();
        //        var oRetF3 = new RetStructF3();
        //        var oCnc = new ClsCNC();
        //        var conflevel = new List<int>();
        //        conflevel.Clear();
        //        int iPage;
        //        var loopTo = oMeta.PageCount;
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
        //                                                            drDataRow["Word No"] = i + iIndex;
        //                                                            drDataRow["LineWordNo"] = CncLines[line].Word[i].WordNumber;
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
        //                                        //if (sKeyarray[jLoop].Trim().ToUpper() == "PO" && Words[i].strWord.ToString().ToUpper().Trim().Contains("PO") && Words[i].strWord.ToString().ToUpper() != "PO")
        //                                        //{
        //                                        //    continue;
        //                                        //}

        //                                        drDataRow = dtKeywordInfo.NewRow();
        //                                        drDataRow["Keyword"] = sKeyarray[jLoop].Trim().ToUpper().Replace("'", "");
        //                                        drDataRow["Page No"] = iPage;
        //                                        drDataRow["Line No"] = CncLines[line].Word[i].LineNo;
        //                                        drDataRow["Word No"] = i;
        //                                        drDataRow["LineWordNo"] = CncLines[line].Word[i].WordNumber;
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
        //        MessageBox.Show(ex.Message, "SearchKeywords");
        //    }

        //    return dtKeywordInfo;
        //}
        //public DataTable MakeDt()
        //{
        //    var dtKeyword = new DataTable();
        //    dtKeyword.Columns.Add("Keyword", typeof(string));
        //    dtKeyword.Columns.Add("Page No", typeof(string));
        //    dtKeyword.Columns.Add("Line No", typeof(string));
        //    dtKeyword.Columns.Add("Word No", typeof(string));
        //    dtKeyword.Columns.Add("LineWordNo", typeof(string));
        //    dtKeyword.Columns.Add("X1", typeof(string));
        //    dtKeyword.Columns.Add("Y1", typeof(string));
        //    dtKeyword.Columns.Add("X2", typeof(string));
        //    dtKeyword.Columns.Add("Y2", typeof(string));
        //    return dtKeyword;
        //}

        ////lavenstein method to check character changed value
        //private bool CheckValidWord(string sWord, string[] arrcheckkeywords)
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
        //        lengthofword = sWord.Length;
        //        if (lengthofword <= 5)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                if (arrcheckkeywords[iLoop].Length <= 5)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 6 & lengthofword <= 8)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                if (arrcheckkeywords[iLoop].Length >= 4 && arrcheckkeywords[iLoop].Length <= 8)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 9 & lengthofword <= 11)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                if (arrcheckkeywords[iLoop].Length >= 9 && arrcheckkeywords[iLoop].Length <= 11)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }
        //        else if (lengthofword >= 12)
        //        {
        //            for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            {
        //                if (arrcheckkeywords[iLoop].Length >= 8)
        //                {
        //                    arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
        //                }
        //            }
        //        }

        //        for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)

        //        {
        //            objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
        //            if (lengthofword <= 5)
        //            {
        //                if (objRestruct058.CharChanged < 1)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 6 & lengthofword <= 8)
        //            {
        //                if (objRestruct058.CharChanged < 2)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 9 & lengthofword <= 11)
        //            {
        //                if (objRestruct058.CharChanged < 3)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //            else if (lengthofword >= 12)
        //            {
        //                if (objRestruct058.CharChanged < 4)
        //                {
        //                    bFlag = true;
        //                }
        //            }
        //        }
        //    }

        //    return bFlag;
        //}

        //private clsCnCWord GetNextJoinWord(clsCncMetaData oMeta, clsCnCWord Word, int PageNo, int LineNo)
        //{
        //    clsCnCWord oWord = null;
        //    clsCnCWord sJoinWord = Word;
        //    string[] sStringArray = new string[oMeta.Page[PageNo].Line[LineNo].WordCount + 1];
        //    Regex regex = new Regex("^([a-zA-Z]*[-./]?[0-9]+[-./]?[a-zA-Z]*[-./]?)*$");
        //    RetStructIQSBR008 objRetStructIQSBR008 = new RetStructIQSBR008();
        //    clsCNCSBR objclsCNCSBR = new clsCNCSBR();
        //    string[] sSpecialCharacter = new string[] { "-" };
        //    int iCharacterIndex = 0;
        //    int iWordIndex;
        //    ClsCNC oClsCNC = new ClsCNC();
        //    clsCnCWord[] oWordsarr = new clsCnCWord[6];
        //    try
        //    {
        //        foreach (string Character in sSpecialCharacter)
        //        {
        //            if (oMeta.Page[PageNo].Line[LineNo].strLine.ToUpper().Contains(Character))
        //            {
        //                clsCnCWord[] words = oMeta.Page[PageNo].Line[LineNo].Word;
        //                for (int iIndex = 0, loopTo = words.Length - 1; iIndex <= loopTo; iIndex++)
        //                {
        //                    if ((words[iIndex]) != null)
        //                    {
        //                        sStringArray[iIndex] = words[iIndex].strWord;
        //                    }
        //                }

        //                iCharacterIndex = Array.IndexOf(sStringArray, Character);
        //                iWordIndex = Array.IndexOf(sStringArray, Word.strWord);
        //                if (iCharacterIndex == iWordIndex + 1)
        //                {
        //                    if (regex.IsMatch(sStringArray[iCharacterIndex + 1]))
        //                    {
        //                        if (IsDate(sStringArray[iCharacterIndex + 1]))
        //                            return sJoinWord;
        //                        objRetStructIQSBR008 = objclsCNCSBR.IQSBR008(sStringArray[iCharacterIndex + 1], "E");
        //                        if (objRetStructIQSBR008.Status == "S")
        //                            return sJoinWord;
        //                        oWordsarr[0] = sJoinWord;
        //                        oWordsarr[1] = words[iCharacterIndex];
        //                        oWordsarr[2] = words[iCharacterIndex + 1];
        //                        oWord = oClsCNC.MergeWords(oWordsarr);
        //                    }
        //                    else
        //                    {
        //                        oWord = sJoinWord;
        //                    }
        //                }
        //                else
        //                {
        //                    oWord = sJoinWord;
        //                }
        //            }
        //            else
        //            {
        //                oWord = sJoinWord;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        //MsgBox(ex.Message, MsgBoxStyle.Critical, "F3:GetNextJoinWord");

        //    }

        //    return oWord;
        //}

        //private bool FoundWord(List<clsCnCWord> CNCWord, clsCnCWord NewWord)
        //{
        //    try
        //    {
        //        foreach (clsCnCWord Word in CNCWord)
        //        {
        //            if ((Word != null) && Word.strWord.Equals(NewWord.strWord))
        //            {
        //                return true;
        //            }
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        MessageBox.Show(ex.Message);
        //    }

        //    return false;
        //}



        //private RetStructF3 GetTotalInvoices_InvoiceNumber(clsCncMetaData ObjMetaData, int intCurrPageNo, List<clsCnCWord> FiltredInvoices, string strArgs)
        //{
        //    var ObjRetStructF3 = new RetStructF3();
        //    var oFoundWord = new List<clsCnCWord>();
        //    var oClsCNC = new ClsCNC();
        //    string KeyWords = strArgs;
        //    string strFirst = string.Empty;
        //    string strSecond = string.Empty;
        //    clsCnCWord oTempword;
        //    for (int iLoop = 0, loopTo = FiltredInvoices.Count - 1; iLoop <= loopTo; iLoop++)
        //    {
        //        try
        //        {
        //            string sLeftWords = string.Empty;
        //            string sTopWords = string.Empty;
        //            int iCount = 0;

        //            // If Amount are Found More Than 3
        //            if (oFoundWord.Count >= 3)
        //            {
        //                break;
        //            }

        //            // Check left words
        //            string sLeft = string.Empty;
        //            sLeftWords = GetLeftWords(ObjMetaData, FiltredInvoices[iLoop], FiltredInvoices[iLoop].PageNo, iCount, ref sLeft);
        //            if (FindKeywordsInvoice_InvoiceNumber(KeyWords, sLeftWords))
        //            {
        //                iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(FiltredInvoices[iLoop], ObjMetaData, FiltredInvoices[iLoop].PageNo);
        //                if (oBoundingWords.LeftWord is object)
        //                {
        //                    //if (Information.IsNumeric(oBoundingWords.LeftWord.strWord) && oBoundingWords.LeftWord.strWord.Length <= 3)
        //                    if ((oBoundingWords.LeftWord.strWord.All(char.IsNumber)) && oBoundingWords.LeftWord.strWord.Length <= 3)
        //                    {
        //                        var oWord = new clsCnCWord();
        //                        var oWords = new clsCnCWord[3];
        //                        oWords[0] = oBoundingWords.LeftWord;
        //                        oWords[1] = FiltredInvoices[iLoop];
        //                        // To Check Space Separated words if they are at minimum space merge that words else tke right side word
        //                        if (oWords[1].Left - oWords[0].Right <= 40)
        //                        {
        //                            oWord = oClsCNC.MergeWords(oWords);
        //                            oFoundWord.Add(oWord);
        //                            continue;
        //                        }
        //                        else
        //                        {
        //                            oFoundWord.Add(oWords[1]);
        //                            continue;
        //                        }
        //                    }
        //                    else
        //                    {
        //                        oFoundWord.Add(FiltredInvoices[iLoop]);
        //                        continue;
        //                    }
        //                }
        //                else
        //                {
        //                    oFoundWord.Add(FiltredInvoices[iLoop]);
        //                    continue;
        //                }
        //            }

        //            // Check Top words
        //            sTopWords = GetTopWords(ObjMetaData, FiltredInvoices[iLoop], FiltredInvoices[iLoop].PageNo);
        //            if (FindKeywordsInvoice_InvoiceNumber(KeyWords, sTopWords))
        //            {
        //                iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(FiltredInvoices[iLoop], ObjMetaData, FiltredInvoices[iLoop].PageNo);
        //                if (oBoundingWords.LeftWord is object)
        //                {
        //                    //if (Information.IsNumeric(oBoundingWords.LeftWord.strWord) && oBoundingWords.LeftWord.strWord.Length <= 3)
        //                    if ((oBoundingWords.LeftWord.strWord.All(char.IsNumber) && oBoundingWords.LeftWord.strWord.Length <= 3))
        //                    {
        //                        var oWord = new clsCnCWord();
        //                        var oWords = new clsCnCWord[3];
        //                        oWords[0] = oBoundingWords.LeftWord;
        //                        oWords[1] = FiltredInvoices[iLoop];
        //                        // To Check Space Separated words if they are at minimum space merge that words else tke right side word
        //                        if (oWords[1].Left - oWords[0].Right <= 40)
        //                        {
        //                            oWord = oClsCNC.MergeWords(oWords);
        //                            oFoundWord.Add(oWord);
        //                        }
        //                        else
        //                        {
        //                            oFoundWord.Add(oWords[1]);
        //                        }
        //                    }
        //                    else
        //                    {
        //                        oFoundWord.Add(FiltredInvoices[iLoop]);
        //                    }
        //                }
        //                else
        //                {
        //                    oFoundWord.Add(FiltredInvoices[iLoop]);
        //                }
        //            }
        //        }
        //        catch (Exception ex)
        //        {
        //            // MsgBox(ex.Message, MsgBoxStyle.Critical, "iQGeneric-F3A17_IsValidDate")
        //            iQPUBLIC.Common.ErrorLog(ex, "IQGeneric-F3C90_IsValidAmount");
        //        }
        //    }

        //    var ConfLevel = new List<int>();

        //    // *********************************************************************************************


        //    // To Decide Confidence of Amounts on Occurance
        //    ObjRetStructF3.Words = oFoundWord;
        //    ObjRetStructF3.NoOfieldsSuspects = oFoundWord.Count;
        //    switch (oFoundWord.Count)
        //    {
        //        case 1:
        //            {
        //                ConfLevel.Add(90);
        //                ObjRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
        //                ObjRetStructF3.Flag = "Y";
        //                ObjRetStructF3.ManualConfirmation = "Y";
        //                ObjRetStructF3.Status = "S";
        //                break;
        //            }

        //        case 2:
        //            {
        //                int intConf = 80;
        //                for (int intI = 0, loopTo1 = ObjRetStructF3.Words.Count - 1; intI <= loopTo1; intI++)
        //                {
        //                    ObjRetStructF3.Words[intI].Confidence = intConf;
        //                    ConfLevel.Add(intConf);
        //                    intConf = intConf - 5;
        //                }

        //                ObjRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
        //                ObjRetStructF3.Flag = "Y";
        //                ObjRetStructF3.ManualConfirmation = "Y";
        //                ObjRetStructF3.Status = "S";
        //                break;
        //            }

        //        case 3:
        //            {
        //                int intConf = 85;
        //                for (int intI = 0, loopTo2 = ObjRetStructF3.Words.Count - 1; intI <= loopTo2; intI++)
        //                {
        //                    ObjRetStructF3.Words[intI].Confidence = intConf;
        //                    ConfLevel.Add(intConf);
        //                    intConf = intConf - 5;
        //                }

        //                ObjRetStructF3.ConfidenceLevelofSuspect = ConfLevel;
        //                ObjRetStructF3.Flag = "Y";
        //                ObjRetStructF3.ManualConfirmation = "Y";
        //                ObjRetStructF3.Status = "S";
        //                break;
        //            }

        //        default:
        //            {
        //                ObjRetStructF3.Status = "F";
        //                break;
        //            }
        //    }

        //    return ObjRetStructF3;
        //}

        //private string GetLeftWords(clsCncMetaData metaData, clsCnCWord targetWord, int pageNo, int iCount, ref string sLeftWords)
        //{
        //    try
        //    {
        //        var oClsCNC = new ClsCNC();
        //        // Static sLeftWords As String = String.Empty

        //        if (targetWord == null)
        //        {
        //            return sLeftWords;
        //        }

        //        iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(targetWord, metaData, targetWord.PageNo);
        //        if (oBoundingWords == null)
        //        {
        //            return sLeftWords;
        //        }

        //        if (oBoundingWords.LeftWord == null)
        //        {
        //            return sLeftWords;
        //        }

        //        // changes by manoj kad on 6 march 2015

        //        sLeftWords = oBoundingWords.LeftWord.strWord + " " + sLeftWords;

        //        // sLeftWords = sLeftWords & " " & oBoundingWords.LeftWord.strWord

        //        while (iCount < 8)
        //        {
        //            iCount += 1;
        //            GetLeftWords(metaData, oBoundingWords.LeftWord, oBoundingWords.LeftWord.PageNo, iCount, ref sLeftWords);
        //        }
        //    }
        //    catch (Exception ex)
        //    {
        //        // MsgBox(ex.Message, MsgBoxStyle.Critical, "iQGeneric-F3A17_GetLeftWords")
        //        iQPUBLIC.Common.ErrorLog(ex, "INVOICE : GetLeftWords");
        //    }

        //    return sLeftWords;
        //}

        //private bool FindKeywordsInvoice_InvoiceNumber(string Keywords, string ROIWords)
        //{
        //    if (string.IsNullOrEmpty(ROIWords.Trim()))
        //        return false;
        //    var sKeyarray = new string[] {"Bill of Lading Number", "BOL #", "B/L #", "Manifest #", "Bill of Lading No", "Shipper's bill of loading no.", "BOL#", "BOL No","Bill of L ading No","Bill of Lading", "BI LL 0 F LADING",
        //                                       "Shipper's Bill of Ladias No","BOL Number","Bill of Lading #","Bill of LadingNumber","Purchase Order No","Purchase Order Number","Po#","Po #", "Customer Order Number","Customer Order No","Customer PO Number", "Customer PO No",
        //                                       "Customer's PO","cust po#","purchase order","customer p.o. no","Customer P.O. No","customer p.o. number", "PO Number","Purchase Ader","Customer PO","P.O. #","Consisnee's Refer.noe/po No","P.O. No","Purchase Order #","Cust P.O. No",
        //                                       "CUSTOMERORDERNUMBER","Order", "Order #","Shipper Number",  "Ship ID#", "Shipper Ref #", "SH PPER NUMBER", "SHIPPER'S NUMBER", "Shippers No", "Shipper No", "Shipment #", "Shipper's No","Shipper #", "Shipping Order No", "SID#", "Shipper No/Customer",
        //                                       "Shipment ID","Shipment No","Shipment number","SID","SHIPPER","PO"};
        //    var arrKeyword = Keywords.Split(',');
        //    for (int iLoop = 0, loopTo = sKeyarray.Length - 1; iLoop <= loopTo; iLoop++)
        //    {
        //        if (ROIWords.Trim().ToUpper().Contains(sKeyarray[iLoop].Trim().ToUpper()))
        //        {
        //            return true;
        //        }
        //    }

        //    return false;
        //}

        //private string GetTopWords(clsCncMetaData metaData, clsCnCWord targetWord, int pageNo)
        //{
        //    string sTopWords = string.Empty;
        //    try
        //    {
        //        var oClsCNC = new ClsCNC();
        //        if (targetWord == null)
        //        {
        //            return sTopWords;
        //        }

        //        iQDataProvider.clsCnCBoundingWords oBoundingWords = oClsCNC.GetBoundingWords(targetWord, metaData, targetWord.PageNo);
        //        if (oBoundingWords == null)
        //        {
        //            return sTopWords;
        //        }

        //        if (oBoundingWords.TopWords == null)
        //        {
        //            return sTopWords;
        //        }

        //        for (int iLoop = 0, loopTo = oBoundingWords.TopWords.Length - 1; iLoop <= loopTo; iLoop++)
        //            sTopWords = sTopWords + " " + oBoundingWords.TopWords[iLoop].strWord;
        //    }
        //    catch (Exception ex)
        //    {
        //        // MsgBox(ex.Message, MsgBoxStyle.Critical, "iQGeneric-F3A17_GetBottomWords")
        //        iQPUBLIC.Common.ErrorLog(ex, "INVOICE : GetBottomWords");
        //    }

        //    return sTopWords;
        //}


        //#endregion	


    
}
