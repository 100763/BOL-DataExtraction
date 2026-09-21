using iQDataProvider;
using iQPUBLIC;
using System;
using System.Collections.Generic;
//using System.Windows.Forms;
using System.Text.RegularExpressions;
using System.Data;
using System.Windows;
using System.Linq;
using System.Data.Odbc;
using System.Collections;
using System.Configuration;
//using System.Windows.Forms;
using System.Diagnostics;
//using System.Windows.Forms;



namespace BOLSET4IQ
{

    public class clsF3
    {
        //Stopwatch watch;
        //System.Data.OleDb.OleDbConnection MyConnection;
        //System.Data.DataSet DtSet;
        //System.Data.OleDb.OleDbDataAdapter MyCommand;
        //string path = Application.StartupPath + "\\FRP001 PRO 29012021.xlsx";
        //string As400_ConnectionString = "Driver={iSeries Access ODBC Driver};System=SAIATEST;UID=BOTBOLOCR;PWD=saia789$;Naming=1;DBQ=*USRLIBL;Compression=1;";
        //string As400_ConnectionString = "";

        public RetStructF3 F3(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strRoutineNo, string strArg)
        {
            string isEDI = "0";
            if (strArg.Contains('#'))
            {
                isEDI = strArg.Split('#')[1].Trim();
                strArg = strArg.Split('#')[0].Trim();
            }

            Module1.GetPackageCode();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;



            if (isEDI == "0")
            {
                switch (strRoutineNo)
                {
                    case "CKM256":

                        //strArg = "Quantity~Weight~Cl~Item Description~Commodity Description~Class~NMFC~Qty~Pieces~Units~Qty~CTNS~Pieces/Quantity~Package~WT~Pkgs~Net Weight~Gross Weight";
                        Module1.objRetStructF3 = F3_AutoLocate_LineItem(ObjMetaData, intCurrPageNumber, strArg);

                        break;


                    case "CKM257":
                        //watch = System.Diagnostics.Stopwatch.StartNew();

                        if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_LineItemValue") == true)
                        {
                            PublicComponents.htMyVariable.Remove("BOLSET4_LineItemValue");
                        }

                        Module1.objRetStructF3 = F3_LineData(ObjMetaData, intCurrPageNumber, strArg);

                        break;
                }
            }
            else
            {
                if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_LineItemValue") == true)
                {
                    PublicComponents.htMyVariable.Remove("BOLSET4_LineItemValue");
                }

                List<clsCnCWord> PossibleWords = new List<clsCnCWord>();
                List<int> conflevel = new List<int>();
                PossibleWords.Add(StringToCncWord_EDI(1, "*****"));

                Module1.objRetStructF3.Words = PossibleWords;
                Module1.objRetStructF3.NoOfieldsSuspects = PossibleWords.Count;
                conflevel.Add(90);
                Module1.objRetStructF3.ConfidenceLevelofSuspect = conflevel;
                Module1.objRetStructF3.Flag = "Y";
                Module1.objRetStructF3.ManualConfirmation = "N";
                Module1.objRetStructF3.Status = "S";
            }


            return Module1.objRetStructF3;

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

        private RetStructF3 F3_LineData(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {

            //RetStructF3 objRetStructF3 = new RetStructF3();
            Module1.objRetStructF3.Status = "F";
            Module1.objRetStructF3.NoOfieldsSuspects = 0;

            DataTable dtHeader = null, dtDetail = null, dtMapTable = null, dtDetailFinal = null;
            string strNMFC = string.Empty, strSub = string.Empty, strClassCode = string.Empty;


            DataRow[] drdata = null;
            DataRow[] drdata1 = null;
            DataRow[] drdata3 = null;
            DataRow[] drdata4 = null;
            DataRow[] drdata8 = null;
            DataRow[] drdata9 = null;


            DataRow drmap = null;

            string[] arrNMFC;
            string Pkgcode, code;
            int PreviousLineno = 0;

            int hx1, hy1, hx2, hy2;
            int x1, y1, x2, y2;

            Rect myRectangle2, myRectangle1;

            char[] seperator = { '\\' };

            string[] batchname = PublicComponents.InputBatchPath.Split(seperator);

            string[] file = PublicComponents.PROPath.Split(seperator);


            string[] seperator1 = { "PRO" };
            //string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";

            string[] arrFilename = file[file.Count() - 1].Replace(".PRO", "").Split('_');

            string PROFilename;
            clsCnCWord word;

            try
            {
                dtHeader = GetHeaderLine(ObjMetaData, intCurrPageNumber, strArg); //to find headerline
                dtDetail = GetDataLineDeail(ObjMetaData, intCurrPageNumber, strArg);

                dtMapTable = new DataTable();

                if (dtDetail.Rows.Count == 0 && dtHeader.Rows.Count > 0)
                {
                    int LineNo = Convert.ToInt32(dtHeader.Rows[0]["Line No"]);
                    int maxCount = LineNo + 5;

                    dtDetail = new DataTable();
                    dtDetail.Columns.Add("Page No");
                    dtDetail.Columns.Add("Line No");
                    dtDetail.Columns.Add("Word No");
                    dtDetail.Columns.Add("Word");
                    dtDetail.Columns.Add("Left");
                    dtDetail.Columns.Add("Top");
                    dtDetail.Columns.Add("Right");
                    dtDetail.Columns.Add("Bottom");
                    dtDetail.Columns.Add("ObjWord", typeof(clsCnCWord));

                    #region To capture lines for Containing Pallet word if it DetailLines not found from numeric value logic
                    for (int k = LineNo + 1; k <= maxCount; k++)
                    {
                        clsCnCLine ObjLine = ObjMetaData.Page[intCurrPageNumber].Line[k];

                        if (ObjLine.strLine.ToUpper().Contains("TOTAL"))
                            break;

                        foreach (clsCnCWord item in ObjLine.Word)
                        {
                            if (item != null)
                            {
                                if ((item.strWord.ToUpper().Contains("PALLET") || item.strWord.ToUpper().Contains("PLT")))
                                {

                                    foreach (clsCnCWord dword in ObjMetaData.Page[intCurrPageNumber].Line[k].Word)
                                    {
                                        if (dword != null)
                                        {
                                            DataRow dr = dtDetail.NewRow();

                                            dr["Page No"] = intCurrPageNumber;
                                            dr["Line No"] = k;
                                            dr["Word No"] = dword.WordNumber;
                                            dr["Word"] = dword.strWord;
                                            dr["Left"] = dword.Left;
                                            dr["Top"] = dword.Top;
                                            dr["Right"] = dword.Right;
                                            dr["Bottom"] = dword.Bottom;
                                            dr["ObjWord"] = dword;
                                            dtDetail.Rows.Add(dr);

                                        }
                                    }
                                }
                            }
                        }

                    }
                }

                    #endregion

                #region logic to add Header and LineItems in dtMapTable
                int i = 0;
                if (dtHeader != null && dtDetail != null)
                {
                    if (dtHeader.Rows.Count > 0 && dtDetail.Rows.Count > 0)
                    {
                        foreach (DataRow dr in dtHeader.Rows)
                        {
                            if (dtMapTable.Columns.Contains(dr["word"].ToString()))
                            {
                                dtMapTable.Columns.Add(dr["word"].ToString() + i, typeof(clsCnCWord));
                                dr["word"] = dr["word"].ToString() + i;
                                dtHeader.AcceptChanges();
                                i++;
                            }
                            else
                                dtMapTable.Columns.Add(dr["word"].ToString(), typeof(clsCnCWord));
                        }


                        drmap = null;
                        PreviousLineno = 0;
                        int ROWID = 0;

                        foreach (DataRow item in dtDetail.Rows)
                        {

                            int Lineno = Convert.ToInt32(item["Line No"]);


                            if (PreviousLineno == 0 || (Lineno != PreviousLineno))
                            {
                                if (drmap != null && PreviousLineno != 0)
                                    dtMapTable.Rows.Add(drmap);
                                drmap = dtMapTable.NewRow();
                            }
                            ROWID++;



                            x1 = Convert.ToInt32(item["left"]);
                            y1 = 0;
                            x2 = Convert.ToInt32(item["right"]);
                            y2 = 1;

                            int headercount = 0;

                            foreach (DataRow dr in dtHeader.Rows)
                            {

                                x1 = Convert.ToInt32(item["left"]);
                                y1 = 0;
                                x2 = Convert.ToInt32(item["right"]);
                                y2 = 1;

                                hx1 = Convert.ToInt32(dr["left"]);
                                hy1 = 0;
                                hx2 = Convert.ToInt32(dr["right"]);
                                hy2 = 1;

                                headercount++;

                                if (hx1 == x1 && hx2 == x2)
                                    continue;
                                if (Convert.ToInt32(dr["Line No"]) > Convert.ToInt32(item["Line No"]))
                                    continue;


                                myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);
                                bool doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                //if (!doesIntersect)
                                //{
                                //    if (headercount == dtHeader.Rows.Count)
                                //    {
                                //        if (dtHeader.Rows[dtHeader.Rows.Count - 1]["word"].ToString().ToUpper().Contains("WEIGHT") || dtHeader.Rows[dtHeader.Rows.Count - 1]["word"].ToString().ToUpper().Contains("NMFC") ||
                                //            dtHeader.Rows[dtHeader.Rows.Count - 1]["word"].ToString().ToUpper().Contains("CLASS"))
                                //        {
                                //            hx1 = hx1 + 50;
                                //            hx2 = hx2 + 50;

                                //        }
                                //    }

                                //    //if ((hx1 - x1) < 50 && headercount == 1)
                                //    //{

                                //    //}
                                //    //else if (headercount == 1 && (hx1 - x1) < 100)
                                //    //{
                                //    //    x1 = x1 + 50;
                                //    //    x2 = x2 + 50;
                                //    //}
                                //    //else if (headercount == 1)
                                //    //{
                                //    //    x1 = x1 + 70;
                                //    //    x2 = x2 + 70;
                                //    //}

                                //    myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                //    myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);
                                //    doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                //    //if (!doesIntersect && ROWID == 1)
                                //    //{
                                //    //    x1 = hx1;

                                //    //    if ((x2 - x1) > 0)
                                //    //        myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                //    //    if (hx2 - hx1 > 0)
                                //    //        myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);

                                //    //    doesIntersect = myRectangle1.IntersectsWith(myRectangle2);
                                //    //}
                                //}


                                if (doesIntersect)
                                {
                                    //drmap[dr["word"].ToString()] = drmap[dr["word"].ToString()].ToString() + " " + item["Wword"];
                                    if (!Convert.IsDBNull(item["ObjWord"]))
                                    {
                                        Module1.ObjWord = (clsCnCWord)item["ObjWord"];

                                        if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                             (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                            Module1.ObjWord.Flag = "1";
                                        else
                                            Module1.ObjWord.Flag = "0";

                                        Module1.ObjWord.Remarks = "3F";
                                        item["ObjWord"] = Module1.ObjWord;
                                        Module1.ObjWord = null;
                                    }

                                    drmap[dr["word"].ToString()] = item["ObjWord"];
                                    break;
                                }
                            }


                            PreviousLineno = Convert.ToInt32(item["Line No"]);
                        }

                        if (drmap != null)
                            dtMapTable.Rows.Add(drmap);
                    }
                }

                #endregion

                dtDetailFinal = MakedtDetailFinal();
                DataRow drdetail;


                #region Logic To assign Values from dtmapTable to finalTable
                for (int k = 0; k < dtMapTable.Rows.Count; k++)
                {
                    drdetail = dtDetailFinal.NewRow();

                    drdetail["Line Number"] = k + 1;
                    drdetail["IsValid"] = 0;

                    Module1.ObjDescription = new clsCnCWord();
                    Module1.ObjNMFC = new clsCnCWord();
                    Module1.ObjClassCode = new clsCnCWord();
                    Module1.ObjWeight = new clsCnCWord();
                    Module1.ObjQuantity = new clsCnCWord();
                    Module1.ObjHazmatCode = new clsCnCWord();
                    Module1.ObjPalletCode = new clsCnCWord();

                    for (int j = 0; j < dtMapTable.Columns.Count; j++)
                    {
                        if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("DESCRIPTION"))
                        {
                            if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                            {

                                Module1.ObjDescription = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                drdetail["Description"] = Module1.ObjDescription.strWord;
                                drdetail["WDescription"] = Module1.ObjDescription;
                                drdetail["PROLineNo"] = Module1.ObjDescription.LineNo;
                            }
                        }
                        if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("PALLET"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Pallate Code"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Module1.ObjPalletCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];


                                    Pkgcode = Module1.ObjPalletCode.strWord.Replace("(s)", "").Replace("(", "").Replace(")", "");
                                    string quantity = Regex.Match(Pkgcode, @"\d+").Value;

                                    if (!string.IsNullOrEmpty(quantity))
                                    {
                                        double result;

                                        Module1.ObjQuantity = GetDefaultWord(intCurrPageNumber, quantity);

                                        if (double.TryParse(quantity, out  result))
                                        {
                                            drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                            drdetail["WQuantity"] = Module1.ObjQuantity;
                                            drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                        }
                                    }



                                    Module1.ObjPalletCode.strWord = "PT";
                                    Module1.ObjPalletCode.Flag = "1";

                                    drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                    drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                    drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;

                                }
                            }

                        }
                        if (!dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("TYPE0"))
                        {
                            if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("TYPE"))
                            {
                                if (string.IsNullOrEmpty(Convert.ToString(drdetail["Pallate Code"])))
                                {
                                    if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                    {
                                        Module1.ObjPalletCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];


                                        Pkgcode = Module1.ObjPalletCode.strWord.Replace("(s)", "").Replace("(", "").Replace(")", "");
                                        string quantity = Regex.Match(Pkgcode, @"\d+").Value;

                                        if (!string.IsNullOrEmpty(quantity))
                                        {
                                            double result;

                                            Module1.ObjQuantity = GetDefaultWord(intCurrPageNumber, quantity);


                                            if (Module1.ObjPalletCode.strWord.Length < 12)
                                            {

                                                if (double.TryParse(quantity, out  result))
                                                {
                                                    drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                                    drdetail["WQuantity"] = Module1.ObjQuantity;
                                                    drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                                }
                                            }
                                        }

                                        Pkgcode = Regex.Replace(Pkgcode, @"[\d-]", string.Empty);
                                        Pkgcode = Pkgcode.Trim();

                                        if (!string.IsNullOrEmpty(Pkgcode))
                                        {
                                            bool isfound = false;
                                            foreach (packagecode gg in Module1.PackageCodes)
                                            {

                                                if (Convert.ToString(gg.Value).ToUpper() == Pkgcode.ToUpper().Trim('S').Trim('(').Trim(')'))
                                                {

                                                    Module1.ObjPalletCode.strWord = gg.Key.ToString();
                                                    Module1.ObjPalletCode.Flag = "1";

                                                    drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                                    drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                                    drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                                    isfound = true;
                                                }
                                            }

                                            //if (!isfound)
                                            //{

                                            //    if (Module1.ObjPalletCode.strWord.Length <= 8)
                                            //    {
                                            //        Module1.ObjPalletCode.Flag = "1";

                                            //        drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                            //        drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                            //        drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                            //    }
                                            //}
                                        }
                                    }
                                }

                            }
                        }
                        if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("TYPE0"))
                        {

                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Pallate Code"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Module1.ObjPalletCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];


                                    Pkgcode = Module1.ObjPalletCode.strWord.Replace("(s)", "").Replace("(", "").Replace(")", "");
                                    Pkgcode = Regex.Replace(Pkgcode, @"[\d-]", string.Empty).Trim();
                                    if (!string.IsNullOrEmpty(Pkgcode))
                                    {
                                        bool isfound = false;
                                        foreach (packagecode gg in Module1.PackageCodes)
                                        {

                                            if (Convert.ToString(gg.Value).ToUpper() == Pkgcode.ToUpper().Trim('S').Trim('(').Trim(')'))
                                            {

                                                Module1.ObjPalletCode.strWord = gg.Key.ToString();
                                                Module1.ObjPalletCode.Flag = "1";

                                                drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                                drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                                drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                                isfound = true;
                                            }
                                        }

                                        //if (!isfound)
                                        //{

                                        //    if (Module1.ObjPalletCode.strWord.Length <= 8)
                                        //    {
                                        //        Module1.ObjPalletCode.Flag = "1";

                                        //        drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                        //        drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                        //        drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                        //    }
                                        //}
                                    }
                                }
                            }

                        }
                        if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("PACKAGE TYPE"))
                        {

                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Pallate Code"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Module1.ObjPalletCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];


                                    Pkgcode = Module1.ObjPalletCode.strWord.Replace("(s)", "").Replace("(", "").Replace(")", "");
                                    Pkgcode = Regex.Replace(Pkgcode, @"[\d-]", string.Empty).Trim();

                                    if (!string.IsNullOrEmpty(Pkgcode))
                                    {
                                        bool isfound = false;
                                        foreach (packagecode gg in Module1.PackageCodes)
                                        {

                                            if (Convert.ToString(gg.Value).ToUpper() == Pkgcode.ToUpper().Trim('S').Trim('(').Trim(')'))
                                            {

                                                Module1.ObjPalletCode.strWord = gg.Key.ToString();
                                                Module1.ObjPalletCode.Flag = "1";

                                                drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                                drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                                drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                                isfound = true;
                                            }
                                        }

                                        //if (!isfound)
                                        //{

                                        //    if (Module1.ObjPalletCode.strWord.Length <= 8)
                                        //    {
                                        //        Module1.ObjPalletCode.Flag = "1";

                                        //        drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                        //        drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                        //        drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                        //    }
                                        //}
                                    }
                                }
                            }
                        }

                        //else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("NMFC"))
                        //{
                        //    if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                        //    {
                        //        Module1.ObjNMFC = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                        //        drdetail["NMFC"] = Module1.ObjNMFC.strWord;
                        //        drdetail["WNMFC"] = Module1.ObjNMFC;
                        //    }
                        //}
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("CLASS"))
                        {
                            if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                            {
                                Module1.ObjClassCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                drdetail["Class Code"] = Module1.ObjClassCode.strWord.ToUpper().Replace("CLASS", " ").Replace("CL", "");
                                drdetail["WClass Code"] = Module1.ObjClassCode;
                                drdetail["PROLineNo"] = Module1.ObjClassCode.LineNo;
                            }
                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("WEIGHT") || dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("WEIGHT - LB"))
                        {
                            if (!dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("FREIGHT") && (!dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("HEIGHT")))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Module1.ObjWeight = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                    Module1.ObjWeight.strWord = Module1.ObjWeight.strWord.ToLower().Replace("1bs", "");
                                    double result;

                                    if (double.TryParse(Module1.ObjWeight.strWord.ToLower().Replace("lbs", "").Replace("lb", ""), out  result))
                                    {

                                        double weight = Math.Round(Convert.ToDouble(Module1.ObjWeight.strWord.ToLower().Replace("lbs", "").Replace("lb", "")));
                                        Module1.ObjWeight.strWord = Convert.ToString(weight);


                                        drdetail["Weight"] = Module1.ObjWeight.strWord;
                                        drdetail["WWeight"] = Module1.ObjWeight;
                                        drdetail["PROLineNo"] = Module1.ObjClassCode.LineNo;
                                    }
                                }
                            }
                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("WGT"))
                        {
                            if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                            {
                                Module1.ObjWeight = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                //Module1.ObjWeight.strWord = Module1.ObjWeight.strWord.ToLower().Replace("lbs", "");
                                //Module1.ObjWeight.strWord = Module1.ObjWeight.strWord.ToLower().Replace("lb", "");
                                Module1.ObjWeight.strWord = Module1.ObjWeight.strWord.ToLower().Replace("1bs", "");
                                double result;

                                if (double.TryParse(Module1.ObjWeight.strWord.ToLower().Replace("lbs", "").Replace("lb", ""), out  result))
                                {

                                    double weight = Math.Round(Convert.ToDouble(Module1.ObjWeight.strWord.ToLower().Replace("lbs", "").Replace("lb", "")));
                                    Module1.ObjWeight.strWord = Convert.ToString(weight);


                                    drdetail["Weight"] = Module1.ObjWeight.strWord;
                                    drdetail["WWeight"] = Module1.ObjWeight;
                                    drdetail["PROLineNo"] = Module1.ObjWeight.LineNo;
                                }
                            }

                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("PCS"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);

                                    if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                    {
                                        drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                        drdetail["WQuantity"] = Module1.ObjQuantity;
                                        drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                    }
                                }
                            }

                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("PKGS"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);

                                    if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                    {
                                        drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                        drdetail["WQuantity"] = Module1.ObjQuantity;
                                        drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                    }
                                }
                            }

                        }
                        //else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("PACKAGES"))
                        //{
                        //    if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                        //    {
                        //        if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                        //        {
                        //            Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);
                        //            drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                        //            drdetail["WQuantity"] = Module1.ObjQuantity;
                        //        }
                        //    }

                        //}
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("QUANTITY"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;

                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);
                                    if (Module1.ObjQuantity.strWord.Length < 12)
                                    {

                                        if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                        {
                                            drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                            drdetail["WQuantity"] = Module1.ObjQuantity;
                                            drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                        }
                                    }
                                }
                            }

                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("QTY") || dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("OTY"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);

                                    if (Module1.ObjQuantity.strWord.Length < 12)
                                    {
                                        Pkgcode = Module1.ObjQuantity.strWord.Replace("(s)", "").Replace("(", "").Replace(")", "");
                                        string quantity = Regex.Match(Pkgcode, @"\d+").Value;


                                        if (double.TryParse(quantity, out  result))
                                        {
                                            Module1.ObjQuantity.strWord = quantity;
                                            drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                            drdetail["WQuantity"] = Module1.ObjQuantity;
                                            drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                        }

                                        Pkgcode = Regex.Replace(Pkgcode, @"[\d-]", string.Empty).Trim();

                                        if (!string.IsNullOrEmpty(Pkgcode))
                                        {
                                            bool isfound = false;
                                            foreach (packagecode gg in Module1.PackageCodes)
                                            {

                                                if (Convert.ToString(gg.Value).ToUpper() == Pkgcode.ToUpper().Trim('S').Trim('(').Trim(')'))
                                                {

                                                    Module1.ObjPalletCode.strWord = gg.Key.ToString();
                                                    Module1.ObjPalletCode.Flag = "1";

                                                    drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                                    drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                                    drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                                    isfound = true;
                                                }
                                            }

                                            //if (!isfound)
                                            //{

                                            //    if (Module1.ObjPalletCode.strWord.Length <= 8)
                                            //    {
                                            //        Module1.ObjPalletCode.Flag = "1";

                                            //        drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                            //        drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                            //        drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                            //    }
                                            //}
                                        }
                                    }
                                }
                            }
                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("UNITS"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);
                                    if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                    {

                                        drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                        drdetail["WQuantity"] = Module1.ObjQuantity;
                                        drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                    }
                                }
                            }

                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("CTNS"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);
                                    if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                    {

                                        drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                        drdetail["WQuantity"] = Module1.ObjQuantity;
                                        drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                    }
                                }
                            }
                        }
                        else if (dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("PIECES"))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);
                                    if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                    {

                                        drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                        drdetail["WQuantity"] = Module1.ObjQuantity;
                                        drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                    }
                                }
                            }

                        }
                        else if (CheckValidWord(dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`"), new string[] { "DESCRIPTION", "COMMODITY DESCRIPTION", "ITEM DESCRIPTION" }))
                        {
                            if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                            {
                                Module1.ObjDescription = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                drdetail["Description"] = Module1.ObjDescription.strWord;
                                drdetail["WDescription"] = Module1.ObjDescription;
                                drdetail["PROLineNo"] = Module1.ObjDescription.LineNo;
                            }
                        }
                        else if (CheckValidWord(dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`"), new string[] { "NMFC", "NMFC Number", "NMFC Item", "NMFC No", "NMFC CODE" }))
                        {

                            if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                            {
                                Module1.ObjNMFC = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                if (!string.IsNullOrEmpty(Module1.ObjNMFC.strWord))
                                {
                                    drdetail["NMFC"] = Module1.ObjNMFC.strWord;
                                    drdetail["WNMFC"] = Module1.ObjNMFC;
                                    drdetail["PROLineNo"] = Module1.ObjNMFC.LineNo;
                                }
                            }

                        }
                        else if (CheckValidWord(dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`"), new string[] { "CLASS", "CL", "FRT", "FREIGHT", "CODE CLASS", "CLASS CODE" }))
                        {
                            if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                            {
                                Module1.ObjClassCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                drdetail["Class Code"] = Module1.ObjClassCode.strWord;
                                drdetail["WClass Code"] = Module1.ObjClassCode;
                                drdetail["PROLineNo"] = Module1.ObjClassCode.LineNo;
                            }
                        }
                        else if (CheckValidWord(dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`"), new string[] { "LBS", "WGHT", "WGT", "WElGHT", "WEIGHT", "GROSS WEIGHT", "NET WEIGHT", "WT", "ACT WT" }))
                        {

                            if (!dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("FREIGHT") && (!dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`").Contains("HEIGHT")))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Module1.ObjWeight = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                    Module1.ObjWeight.strWord = Module1.ObjWeight.strWord.ToLower().Replace("1bs", "");

                                    double result;

                                    if (double.TryParse(Module1.ObjWeight.strWord.ToLower().Replace("lbs", "").Replace("lb", ""), out  result))
                                    {
                                        double weight = Math.Round(Convert.ToDouble(Module1.ObjWeight.strWord.ToLower().Replace("lbs", "").Replace("lb", "")));
                                        Module1.ObjWeight.strWord = Convert.ToString(weight);

                                        drdetail["Weight"] = Module1.ObjWeight.strWord;
                                        drdetail["WWeight"] = Module1.ObjWeight;
                                        drdetail["PROLineNo"] = Module1.ObjWeight.LineNo;
                                    }
                                }
                            }
                        }
                        else if (CheckValidWord(dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`"), new string[] { "PC", "PCS", "PKGS", "PACKAGES", "UNITS", "QUANTITY", "QTY1", "CTNS", "PIECES" }))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Quantity"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    double result;
                                    Module1.ObjQuantity = (clsCnCWord)(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]);
                                    if (double.TryParse(Module1.ObjQuantity.strWord, out  result))
                                    {

                                        drdetail["Quantity"] = Module1.ObjQuantity.strWord;
                                        drdetail["WQuantity"] = Module1.ObjQuantity;
                                        drdetail["PROLineNo"] = Module1.ObjQuantity.LineNo;
                                    }
                                }
                            }
                        }
                        else if (CheckValidWord(dtMapTable.Columns[j].ColumnName.ToUpper().Trim().Replace("'", "`"), new string[] { "Packaging", "Package", "Type", "PACKAGE TYPE", "PKG", "UOM" }))
                        {
                            if (string.IsNullOrEmpty(Convert.ToString(drdetail["Pallate Code"])))
                            {
                                if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                {
                                    Module1.ObjPalletCode = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];


                                    Pkgcode = Module1.ObjPalletCode.strWord.Replace("(s)", "").Replace("(", "").Replace(")", "");
                                    Pkgcode = Regex.Replace(Pkgcode, @"[\d-]", string.Empty).Trim();

                                    if (!string.IsNullOrEmpty(Pkgcode))
                                    {
                                        bool isfound = false;
                                        foreach (packagecode gg in Module1.PackageCodes)
                                        {

                                            if (Convert.ToString(gg.Value).ToUpper() == Pkgcode.ToUpper().Trim('S').Trim('(').Trim(')'))
                                            {

                                                Module1.ObjPalletCode.strWord = gg.Key.ToString();
                                                Module1.ObjPalletCode.Flag = "1";

                                                drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                                drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                                drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                                isfound = true;
                                            }
                                        }

                                        //if (!isfound)
                                        //{

                                        //    if (Module1.ObjPalletCode.strWord.Length <= 8)
                                        //    {
                                        //        Module1.ObjPalletCode.Flag = "1";

                                        //        drdetail["Pallate Code"] = Module1.ObjPalletCode.strWord;
                                        //        drdetail["WPallate Code"] = Module1.ObjPalletCode;
                                        //        drdetail["PROLineNo"] = Module1.ObjPalletCode.LineNo;
                                        //    }
                                        //}
                                    }
                                }
                            }
                        }
                        else if (dtMapTable.Columns[j].ColumnName.Split(' ').Count() > 1)
                        {
                            string[] items = dtMapTable.Columns[j].ColumnName.Split(' ');

                            foreach (string item in items)
                            {
                                if (CheckValidWord(item.ToUpper().Trim().Replace("'", "`"), new string[] { "DESC", "DESCRIPTION", "COMMODITY DESCRIPTION", "ITEM DESCRIPTION" }))
                                {
                                    if (!Convert.IsDBNull(dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName]))
                                    {
                                        Module1.ObjDescription = (clsCnCWord)dtMapTable.Rows[k][dtMapTable.Columns[j].ColumnName];
                                        drdetail["Description"] = Module1.ObjDescription.strWord;
                                        drdetail["WDescription"] = Module1.ObjDescription;
                                        drdetail["PROLineNo"] = Module1.ObjDescription.LineNo;
                                    }
                                }
                            }
                        }
                    }

                    dtDetailFinal.Rows.Add(drdetail);


                    Module1.ObjDescription = null;
                    Module1.ObjNMFC = null;
                    Module1.ObjClassCode = null;
                    Module1.ObjWeight = null;
                    Module1.ObjQuantity = null;
                    Module1.ObjHazmatCode = null;
                    Module1.ObjPalletCode = null;

                }

                DataRow[] drdescriptions = dtDetailFinal.Select("Description is not null");

                if (drdescriptions.Count() == 0)
                {
                    int indexvalue = 0;

                    if (dtDetail.Rows.Count % dtHeader.Rows.Count == 0)
                    {
                        foreach (DataRow item in dtHeader.Rows)
                        {
                            if (Convert.ToString(item["word"]).ToUpper().Contains("DESCRIPTION"))
                            {
                                indexvalue = Convert.ToInt32(item["Word No"]);
                                break;
                            }
                        }

                        if (indexvalue != 0)
                        {
                            DataRow[] drdes = dtDetail.Select("[Word No]=" + indexvalue + "");
                            if (dtDetailFinal.Rows.Count == drdes.Length)
                            {
                                foreach (DataRow dritem in dtDetailFinal.Rows)
                                {
                                    DataRow[] dr1 = dtDetail.Select("[Word No]=" + indexvalue + " and [Line No]=" + dritem["PROLineNo"] + "");
                                    if (!Convert.IsDBNull(dr1[0]["ObjWord"]))
                                    {
                                        Module1.ObjWord = (clsCnCWord)dr1[0]["ObjWord"];

                                        if (Module1.ObjWord != null)
                                        {

                                            if (Module1.ObjWord.strWord.Length > 4)
                                            {
                                                if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                                    (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                                    Module1.ObjWord.Flag = "1";
                                                else
                                                    Module1.ObjWord.Flag = "0";

                                                Module1.ObjWord.Remarks = "3F";
                                                dritem["Description"] = dr1[0]["Word"];
                                                dritem["WDescription"] = Module1.ObjWord;

                                                Module1.ObjWord = null;


                                                dtDetailFinal.AcceptChanges();
                                            }
                                        }
                                    }
                                }
                            }

                        }
                    }
                }




                if (arrFilename.Count() == 1)
                {
                    PROFilename = arrFilename[0];
                }
                else
                    PROFilename = arrFilename[arrFilename.Count() - 1];


                code = "";


                if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_ShpCode"))
                {
                    code = Convert.ToString(PublicComponents.htMyVariable["BOLSET4_ShpCode"]);
                }
                else
                {
                    //MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                    //MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where [FHPRO]=" + PROFilename + "", MyConnection);


                    //DtSet = new System.Data.DataSet();
                    //MyCommand.Fill(DtSet);
                    //MyConnection.Close();

                    //MyCommand.Dispose();
                    //MyConnection.Dispose();


                    //if (DtSet != null)
                    //{
                    //    if (DtSet.Tables.Count > 0)
                    //    {
                    //        if (DtSet.Tables[0].Rows.Count > 0)
                    //            code = Convert.ToString(DtSet.Tables[0].Rows[0][1]);
                    //        //else
                    //        //    return objRetStructF3  ;
                    //    }
                    //    //if (code.StartsWith("99"))
                    //    //    return objRetStructF3;
                    //}
                }
                DataTable dtGetData = null;
                if (code.StartsWith("99") || string.IsNullOrEmpty(code) || code.StartsWith("**"))
                {

                }
                else
                {
                    //return;


                    string shipcode = string.Empty;
                    //code = "0739394";

                    shipcode = code.PadLeft(7, '0');
                    //if (code.Trim().Length < 7)
                    //{

                    //    for (int k = 1; k <= 7 - code.Trim().Length; k++)
                    //        shipcode = "0" + shipcode;
                    //}

                    //shipcode = "0107778";

                    //string Query_String = "SELECT F30.SASBLC, F31.SBDES,F31.SBNMFC,F31.SBSUB,F31.SBCMCL FROM FRP030 F30 Inner join FRP031 F31 on F30.SASBLC = F31.SBSBLC where F30.SASCD ='" + batchname[batchname.Count() - 1] + "'";
                    string Query_String = "SELECT F30.SASBLC, F31.SBDES,F31.SBNMFC,F31.SBSUB,F31.SBCMCL FROM FRP030 F30 Inner join FRP031 F31 on F30.SASBLC = F31.SBSBLC where F30.SASCD ='" + shipcode + "'";


                    dtGetData = GetDataTableByText("", Query_String);

                    if (dtGetData.Rows.Count == 0)
                        goto done;

                    int LineNo = 0;
                    try
                    {
                        foreach (DataRow dr in dtDetailFinal.Rows)
                        {
                            if (Convert.ToString(dr["NMFC"]).Contains("-"))
                                arrNMFC = Convert.ToString(dr["NMFC"]).Split('-');
                            else if (Convert.ToString(dr["NMFC"]).Contains("."))
                                arrNMFC = Convert.ToString(dr["NMFC"]).Split('.');
                            else if (Convert.ToString(dr["NMFC"]).Contains("S") || Convert.ToString(dr["NMFC"]).Contains("s"))
                                arrNMFC = Convert.ToString(dr["NMFC"]).Split(' ');
                            else
                                arrNMFC = Convert.ToString(dr["NMFC"]).Split('-');

                            strNMFC = string.Empty; strSub = string.Empty; strClassCode = string.Empty;

                            if (arrNMFC.Count() > 1)
                            {
                                strNMFC = arrNMFC[0].Trim().Replace("l", "");

                                if (arrNMFC[1] == "00" || arrNMFC[1] == "0" || arrNMFC[1] == string.Empty)
                                {
                                    strSub = "0";
                                }
                                else if (arrNMFC[1].Contains("S") || arrNMFC[1].Contains("s"))
                                {
                                    strSub = arrNMFC[1].Substring(1, arrNMFC[1].Length - 1);
                                    if (strSub.Length == 1 || strSub.Length == 2)
                                    {
                                        int  val;
                                        if (int.TryParse(strSub, out val))
                                        {
                                            strSub = strSub.TrimStart('0');
                                        }
                                        else
                                        {
                                            strSub ="0";
                                        }
                                    }
                                    else
                                    {
                                        strSub ="0";
                                    }
                                }
                                else
                                    strSub = arrNMFC[1].Replace('0', ' ').Trim();
                            }
                            else
                            {
                                strNMFC = Convert.ToString(dr["NMFC"]).Replace("l", "");
                                strSub = "0";
                            }



                            Module1.ObjDescription = new clsCnCWord();
                            drdata = null;
                            drdata1 = null;
                            drdata3 = null;
                            drdata4 = null;
                            drdata8 = null;
                            drdata9 = null;

                            strClassCode = Convert.ToString(dr["Class Code"]).Trim();



                            if (!Convert.IsDBNull(dr["WNMFC"]))
                            {
                                word = (clsCnCWord)dr["WNMFC"];
                                LineNo = word.LineNo;
                                Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];
                            }
                            else
                            {
                                Module1.ObjNMFC = null;
                            }

                            if (!Convert.IsDBNull(dr["WQuantity"]))
                            {

                                word = (clsCnCWord)dr["WQuantity"];
                                LineNo = word.LineNo;
                                Module1.ObjQuantity = (clsCnCWord)dr["WQuantity"];
                            }
                            else
                            {
                                Module1.ObjQuantity = null;
                            }

                            if (!Convert.IsDBNull(dr["WClass Code"]))
                            {
                                word = (clsCnCWord)dr["WClass Code"];
                                LineNo = word.LineNo;
                                Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                            }
                            else
                            {
                                Module1.ObjClassCode = null;
                            }

                            if (!Convert.IsDBNull(dr["WDescription"]))
                            {
                                word = (clsCnCWord)dr["WDescription"];
                                LineNo = word.LineNo;
                                Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                            }

                            if (!Convert.IsDBNull(dr["WWeight"]))
                            {
                                word = (clsCnCWord)dr["WWeight"];
                                LineNo = word.LineNo;
                                Module1.ObjWeight = (clsCnCWord)dr["WWeight"];
                            }
                            else
                            {
                                Module1.ObjWeight = null;
                            }

                            if (LineNo > 0 && (strSub == "0" || string.IsNullOrEmpty(strSub) && !string.IsNullOrEmpty(strNMFC)))
                            {
                                clsCnCLine line = PublicComponents.oMetaData1.Page[intCurrPageNumber].Line[LineNo];

                                if (line.strLine.ToUpper().Contains("SUB"))
                                {
                                    clsCnCWord[] words = line.Word;

                                    for (int m = 0; m < words.Count(); m++)
                                    {
                                        if (words[m] != null)
                                        {
                                            if (words[m].strWord.Contains("SUB"))
                                            {
                                                if (words[m + 1].strWord.Length == 1)
                                                    strSub = words[m + 1].strWord;
                                                else
                                                    strSub = words[m].strWord.Substring(words[m].strWord.IndexOf("SUB") + 3, 1);
                                            }
                                        }
                                    }
                                }
                            }
                            if ((strNMFC.Length >= 10 || (string.IsNullOrEmpty(strNMFC)) && LineNo > 0))
                            {
                                foreach (DataRow item in dtGetData.Rows)
                                {
                                    clsCnCLine line = PublicComponents.oMetaData1.Page[intCurrPageNumber].Line[LineNo];

                                    if (line.strLine.Contains(Convert.ToString(item["SBNMFC"]).Trim()))
                                    {
                                        foreach (clsCnCWord wnmfc in line.Word)
                                        {
                                            if (wnmfc != null)
                                            {
                                                if (wnmfc.strWord.Contains(Convert.ToString(item["SBNMFC"]).Trim()) && wnmfc.strWord.Length >= 5 && !wnmfc.strWord.Contains(".") && Convert.ToString(item["SBNMFC"]).Trim() != "0")
                                                {
                                                    if (wnmfc.strWord.Length > 12)
                                                        wnmfc.strWord = Convert.ToString(item["SBNMFC"]);

                                                    wnmfc.Confidence = 100;
                                                    wnmfc.Flag = "1";
                                                    wnmfc.Remarks = "2F";

                                                    dr["WNMFC"] = wnmfc;
                                                    dr["NMFC"] = wnmfc.strWord;
                                                    dtDetailFinal.AcceptChanges();

                                                    if (wnmfc.strWord.Contains('-'))
                                                        arrNMFC = wnmfc.strWord.Split('-');
                                                    else if (wnmfc.strWord.Contains('.'))
                                                        arrNMFC = wnmfc.strWord.Split('.');
                                                    else
                                                        arrNMFC = wnmfc.strWord.Split('-');

                                                    if (arrNMFC.Count() > 1)
                                                    {
                                                        strNMFC = arrNMFC[0].Trim().Replace("l", "");

                                                        if (arrNMFC[1] == "00" || arrNMFC[1] == "0" || arrNMFC[1] == string.Empty)
                                                        {
                                                            strSub = "0";
                                                        }
                                                        else
                                                            strSub = arrNMFC[1].Replace('0', ' ').Trim();
                                                    }
                                                    else
                                                    {
                                                        strNMFC = Convert.ToString(dr["NMFC"]).Replace("l", "");
                                                        //strSub = "0";
                                                    }

                                                    Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];
                                                    Module1.ObjNMFC.strWord = Convert.ToString(dr["NMFC"]);
                                                    Module1.ObjNMFC.Confidence = 100;
                                                    Module1.ObjNMFC.Flag = "1";
                                                    Module1.ObjNMFC.Remarks = "2F";

                                                }
                                            }
                                        }
                                    }

                                    if (line.strLine.Contains(Convert.ToString(item["SBCMCL"]).Trim()))
                                    {
                                        foreach (clsCnCWord wclass in line.Word)
                                        {
                                            if (wclass != null)
                                            {
                                                if (wclass.strWord == Convert.ToString(item["SBCMCL"]).Trim())
                                                {
                                                    wclass.Confidence = 100;
                                                    wclass.Flag = "1";
                                                    wclass.Remarks = "2F";
                                                    dr["WClass Code"] = wclass;
                                                    dr["Class Code"] = wclass.strWord;
                                                    dtDetailFinal.AcceptChanges();
                                                    strClassCode = wclass.strWord;



                                                    Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                                    Module1.ObjClassCode.strWord = Convert.ToString(dr["Class Code"]);
                                                    Module1.ObjClassCode.Confidence = 100;
                                                    Module1.ObjClassCode.Flag = "1";
                                                    Module1.ObjClassCode.Remarks = "2F";
                                                }
                                            }
                                        }
                                    }


                                    if (string.IsNullOrEmpty(Convert.ToString(dr["Description"]).Trim()))
                                    {

                                        foreach (clsCnCWord wDesc in line.Word)
                                        {
                                            if (wDesc != null)
                                            {
                                                clsCNCSBR ObjSbr = new clsCNCSBR();
                                                RetStructIQSBR050 objRet50 = new RetStructIQSBR050();
                                                objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(Convert.ToString(item["SBDES"])).Trim().ToUpper(), Module1.func_RemoveSpecialCharacter(wDesc.strWord.ToUpper()));
                                                if (objRet50.PercentageMatch >= 90)
                                                {
                                                    Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(item["SBDES"]));
                                                    Module1.ObjDescription.strWord = Convert.ToString(item["SBDES"]);
                                                    Module1.ObjDescription.Confidence = 100;
                                                    Module1.ObjDescription.Flag = "1";
                                                    Module1.ObjDescription.Remarks = "2F";



                                                    dr["WDescription"] = Module1.ObjDescription;
                                                }

                                            }
                                        }
                                    }
                                }


                            }


                            if (!strNMFC.All(x => char.IsDigit(x)))
                                goto X;

                            strNMFC = strNMFC.Replace(" ", "");
                            strSub = strSub.Replace(" ", "");
                            strClassCode = strClassCode.Replace(" ", "");

                            if (!string.IsNullOrEmpty(strNMFC) && !string.IsNullOrEmpty(strSub))
                                drdata = dtGetData.Select("SBNMFC='" + strNMFC + "' and SBSUB='" + strSub + "' and SBCMCL='" + strClassCode + "'");

                            if (!string.IsNullOrEmpty(strNMFC) && !string.IsNullOrEmpty(strSub))
                                drdata1 = dtGetData.Select("SBNMFC='" + strNMFC + "' and SBSUB='" + strSub + "'");

                            if (!string.IsNullOrEmpty(strNMFC) && !string.IsNullOrEmpty(strClassCode) && !string.IsNullOrEmpty(strSub))
                                drdata3 = dtGetData.Select("SBNMFC='" + strNMFC + "' and SBSUB='0' and SBCMCL='" + strClassCode + "'");

                            if (!string.IsNullOrEmpty(strNMFC))
                                drdata4 = dtGetData.Select("SBNMFC='" + strNMFC + "'");

                            if (string.IsNullOrEmpty(strNMFC))
                                drdata8 = dtGetData.Select("SBNMFC<>'0'");

                            if (string.IsNullOrEmpty(strNMFC))
                                drdata9 = dtGetData.Select("SBNMFC = '0'");


                            if (drdata != null)
                            {
                                if (drdata.Count() > 0)
                                {
                                    dr["Description"] = drdata[0]["SBDES"];

                                    if (!Convert.IsDBNull(dr["WDescription"]))
                                    {
                                        Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                        Module1.ObjDescription.strWord = Convert.ToString(drdata[0]["SBDES"]);
                                        Module1.ObjDescription.Confidence = 100;
                                        Module1.ObjDescription.Flag = "1";
                                        Module1.ObjDescription.Remarks = "2F";



                                        dr["WDescription"] = Module1.ObjDescription;

                                    }
                                    else
                                    {
                                        Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(drdata[0]["SBDES"]));

                                        Module1.ObjDescription.Confidence = 100;
                                        Module1.ObjDescription.Flag = "1";
                                        Module1.ObjDescription.Remarks = "2F";

                                        dr["WDescription"] = Module1.ObjDescription;
                                    }
                                    dr["IsValid"] = 1;

                                    if (!Convert.IsDBNull(dr["WClass Code"]))
                                    {
                                        Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                        Module1.ObjClassCode.Confidence = 100;
                                        Module1.ObjClassCode.Flag = "1";
                                        Module1.ObjClassCode.Remarks = "3F";
                                        dr["WClass Code"] = Module1.ObjClassCode;

                                    }

                                    if (!Convert.IsDBNull(dr["WNMFC"]))
                                    {
                                        Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];
                                        Module1.ObjNMFC.Confidence = 100;
                                        Module1.ObjNMFC.Flag = "1";
                                        Module1.ObjNMFC.Remarks = "3F";
                                        dr["WNMFC"] = Module1.ObjNMFC;

                                    }
                                    dtDetailFinal.AcceptChanges();
                                    continue;
                                }
                                else
                                {
                                    goto X;
                                }
                            }

                        X:
                            if (drdata1 != null)
                            {
                                if (drdata1.Count() > 0)
                                {
                                    DataRow[] drdata2 = drdata1.CopyToDataTable().Select("SBCMCL<>'" + Convert.ToString(dr["Class Code"]).Trim() + "'");

                                    if (drdata2.Count() > 0)
                                    {
                                        dr["IsValid"] = 1;
                                        dr["Description"] = drdata2[0]["SBDES"];


                                        if (!Convert.IsDBNull(dr["WDescription"]))
                                        {
                                            Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                            Module1.ObjDescription.strWord = Convert.ToString(drdata2[0]["SBDES"]);
                                            Module1.ObjDescription.Flag = "1";
                                            Module1.ObjDescription.Remarks = "2F";

                                            Module1.ObjDescription.Confidence = 100;
                                            dr["WDescription"] = Module1.ObjDescription;

                                        }
                                        else
                                        {
                                            Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(drdata2[0]["SBDES"]));
                                            Module1.ObjDescription.Confidence = 100;
                                            Module1.ObjDescription.Flag = "1";
                                            Module1.ObjDescription.Remarks = "2F";
                                            dr["WDescription"] = Module1.ObjDescription;
                                        }

                                        if (!Convert.IsDBNull(dr["WNMFC"]))
                                        {
                                            Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];
                                            Module1.ObjNMFC.Confidence = 100;
                                            Module1.ObjNMFC.Flag = "1";
                                            Module1.ObjNMFC.Remarks = "3F";
                                            dr["WNMFC"] = Module1.ObjNMFC;

                                        }

                                        dtDetailFinal.AcceptChanges();
                                        continue;
                                    }
                                    else
                                    {
                                        dr["IsValid"] = 0;
                                        dtDetailFinal.AcceptChanges();

                                    }
                                }
                                else
                                {

                                    dr["IsValid"] = 0;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }
                            else
                            {
                                dr["IsValid"] = 0;
                                dtDetailFinal.AcceptChanges();

                            }

                            if (drdata3 != null && strSub != "0")
                            {
                                if (drdata3.Count() > 0)
                                {
                                    dr["IsValid"] = 1;
                                    dr["Description"] = drdata3[0]["SBDES"];

                                    if (!Convert.IsDBNull(dr["WDescription"]))
                                    {
                                        Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                        Module1.ObjDescription.strWord = Convert.ToString(drdata3[0]["SBDES"]);
                                        Module1.ObjDescription.Confidence = 100;
                                        Module1.ObjDescription.Flag = "1";
                                        Module1.ObjDescription.Remarks = "2F";
                                        dr["WDescription"] = Module1.ObjDescription;

                                    }
                                    else
                                    {
                                        Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(drdata3[0]["SBDES"]));
                                        Module1.ObjDescription.Confidence = 100;
                                        Module1.ObjDescription.Flag = "1";
                                        Module1.ObjDescription.Flag = "2F";
                                        dr["WDescription"] = Module1.ObjDescription;
                                    }
                                    if (!Convert.IsDBNull(dr["WClass Code"]))
                                    {
                                        Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                        Module1.ObjClassCode.Confidence = 100;
                                        Module1.ObjClassCode.Flag = "1";
                                        Module1.ObjClassCode.Remarks = "3F";
                                        dr["WClass Code"] = Module1.ObjClassCode;

                                    }

                                    if (!Convert.IsDBNull(dr["WNMFC"]))
                                    {
                                        Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];
                                        Module1.ObjNMFC.Confidence = 100;
                                        Module1.ObjNMFC.Flag = "1";
                                        Module1.ObjNMFC.Remarks = "3F";
                                        dr["WNMFC"] = Module1.ObjNMFC;

                                    }

                                    dtDetailFinal.AcceptChanges();
                                    continue;
                                }
                                else
                                {
                                    dr["IsValid"] = 0;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }
                            else
                            {
                                dr["IsValid"] = 0;
                                dtDetailFinal.AcceptChanges();
                            }

                            if (drdata4 == null && strNMFC != "0")
                            {
                                DataRow[] drdata5 = dtGetData.Select("SBCMCL='" + Convert.ToString(dr["Class Code"]).Trim() + "'");
                                string Description = Convert.ToString(dr["Description"]);

                                if (drdata5 != null && !string.IsNullOrEmpty(Description))
                                {

                                    if (drdata5.Count() > 0)
                                    {
                                        if (CheckValidWord(Convert.ToString(drdata5[0]["SBDES"]).ToUpper(), new string[] { Description.ToUpper() }))
                                        {

                                            dr["IsValid"] = 1;
                                            dr["Description"] = drdata5[0]["SBDES"];


                                            if (!Convert.IsDBNull(dr["WDescription"]))
                                            {
                                                Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                                Module1.ObjDescription.strWord = Convert.ToString(drdata5[0]["SBDES"]);
                                                Module1.ObjDescription.Confidence = 100;
                                                Module1.ObjDescription.Flag = "1";
                                                Module1.ObjDescription.Remarks = "2F";
                                                dr["WDescription"] = Module1.ObjDescription;

                                            }
                                            else
                                            {
                                                Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(drdata5[0]["SBDES"]));
                                                Module1.ObjDescription.Confidence = 100;
                                                Module1.ObjDescription.Flag = "1";
                                                Module1.ObjDescription.Remarks = "2F";
                                                dr["WDescription"] = Module1.ObjDescription;
                                            }

                                            if (!Convert.IsDBNull(dr["WClass Code"]))
                                            {
                                                Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                                Module1.ObjClassCode.Confidence = 100;
                                                Module1.ObjClassCode.Flag = "1";
                                                Module1.ObjClassCode.Remarks = "3F";
                                                dr["WClass Code"] = Module1.ObjClassCode;

                                            }


                                            dtDetailFinal.AcceptChanges();
                                            continue;
                                        }
                                    }
                                    else
                                    {

                                        dr["IsValid"] = 0;
                                        dtDetailFinal.AcceptChanges();
                                    }
                                }
                                else
                                {
                                    dr["IsValid"] = 0;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }

                            //The skeleton has an NMFC#, but the OBL does not 
                            if (drdata8 != null && (string.IsNullOrEmpty(strNMFC) || strNMFC == "0"))
                            {
                                DataRow[] drdata5 = null;

                                if (drdata8.Count() > 0)
                                    drdata5 = drdata8.CopyToDataTable().Select("SBCMCL='" + Convert.ToString(dr["Class Code"]).Trim() + "'");

                                string Description = Convert.ToString(dr["Description"]);

                                if (drdata5 != null && !string.IsNullOrEmpty(Description))
                                {

                                    if (drdata5.Count() > 0)
                                    {
                                        if (CheckValidWord(Convert.ToString(drdata5[0]["SBDES"]).ToUpper(), new string[] { Description.ToUpper() }))
                                        {

                                            dr["IsValid"] = 1;
                                            dr["Description"] = drdata5[0]["SBDES"];


                                            if (!Convert.IsDBNull(dr["WDescription"]))
                                            {
                                                Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                                Module1.ObjDescription.strWord = Convert.ToString(drdata5[0]["SBDES"]);
                                                Module1.ObjDescription.Confidence = 100;
                                                Module1.ObjDescription.Flag = "1";
                                                Module1.ObjDescription.Remarks = "2F";
                                                dr["WDescription"] = Module1.ObjDescription;

                                            }
                                            else
                                            {
                                                Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(drdata5[0]["SBDES"]));
                                                Module1.ObjDescription.Confidence = 100;
                                                Module1.ObjDescription.Flag = "1";
                                                Module1.ObjDescription.Remarks = "2F";
                                                dr["WDescription"] = Module1.ObjDescription;
                                            }

                                            if (!Convert.IsDBNull(dr["WClass Code"]))
                                            {
                                                Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                                Module1.ObjClassCode.Confidence = 100;
                                                Module1.ObjClassCode.Flag = "1";
                                                Module1.ObjClassCode.Remarks = "3F";
                                                dr["WClass Code"] = Module1.ObjClassCode;

                                            }

                                            dtDetailFinal.AcceptChanges();
                                            continue;

                                        }
                                    }
                                    else
                                    {

                                        dr["IsValid"] = 0;
                                        dtDetailFinal.AcceptChanges();
                                    }
                                }
                                else
                                {
                                    dr["IsValid"] = 0;
                                    dtDetailFinal.AcceptChanges();
                                }

                            }

                            //Neither the OBL nor the skeleton has an NMFC#, but the class matches
                            if (drdata9 != null && (string.IsNullOrEmpty(strNMFC) || strNMFC == "0"))
                            {
                                DataRow[] drdata5 = null;
                                if (drdata9.Count() > 0)
                                    drdata5 = drdata9.CopyToDataTable().Select("SBCMCL='" + Convert.ToString(dr["Class Code"]).Trim() + "'");

                                string Description = Convert.ToString(dr["Description"]);

                                if (drdata5 != null && !string.IsNullOrEmpty(Description))
                                {

                                    if (drdata5.Count() > 0)
                                    {
                                        if (CheckValidWord(Convert.ToString(drdata5[0]["SBDES"]).ToUpper(), new string[] { Description.ToUpper() }))
                                        {

                                            dr["IsValid"] = 1;
                                            dr["Description"] = drdata5[0]["SBDES"];


                                            if (!Convert.IsDBNull(dr["WDescription"]))
                                            {
                                                Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                                Module1.ObjDescription.strWord = Convert.ToString(drdata5[0]["SBDES"]);
                                                Module1.ObjDescription.Confidence = 100;
                                                Module1.ObjDescription.Flag = "1";
                                                Module1.ObjDescription.Remarks = "2F";
                                                dr["WDescription"] = Module1.ObjDescription;

                                            }
                                            else
                                            {
                                                Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(drdata5[0]["SBDES"]));
                                                Module1.ObjDescription.Confidence = 100;
                                                Module1.ObjDescription.Flag = "1";
                                                Module1.ObjDescription.Remarks = "2F";
                                                dr["WDescription"] = Module1.ObjDescription;
                                            }

                                            if (!Convert.IsDBNull(dr["WClass Code"]))
                                            {
                                                Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                                Module1.ObjClassCode.Confidence = 100;
                                                Module1.ObjClassCode.Flag = "1";
                                                Module1.ObjClassCode.Remarks = "3F";
                                                dr["WClass Code"] = Module1.ObjClassCode;

                                            }


                                            dtDetailFinal.AcceptChanges();
                                            continue;

                                        }
                                    }
                                    else
                                    {
                                        //The OBL does not have an NMFC#, but has a word description and class code or no class code at all 
                                        //and the skeleton only shows a MEMO description line

                                        //

                                        drdata5 = drdata9.CopyToDataTable().Select("SBCMCL='MEMO'");

                                        if (drdata5 != null && !string.IsNullOrEmpty(Description) && (!string.IsNullOrEmpty(strClassCode) || strClassCode != "0"))
                                        {

                                            if (drdata5.Count() > 0)
                                            {
                                                foreach (var item in drdata5)
                                                {

                                                    if (CheckValidWord(Convert.ToString(item["SBDES"]).ToUpper(), new string[] { Description.Replace("AND", "").ToUpper() }))
                                                    {

                                                        dr["IsValid"] = 1;
                                                        dr["Description"] = item["SBDES"];


                                                        if (!Convert.IsDBNull(dr["WDescription"]))
                                                        {
                                                            Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                                                            Module1.ObjDescription.strWord = Convert.ToString(item["SBDES"]);
                                                            Module1.ObjDescription.Confidence = 100;
                                                            Module1.ObjDescription.Flag = "1";
                                                            Module1.ObjDescription.Remarks = "2F";
                                                            dr["WDescription"] = Module1.ObjDescription;

                                                        }
                                                        else
                                                        {
                                                            Module1.ObjDescription = GetDefaultWord(1, Convert.ToString(item["SBDES"]));
                                                            Module1.ObjDescription.Confidence = 100;
                                                            Module1.ObjDescription.Flag = "1";
                                                            Module1.ObjDescription.Remarks = "2F";
                                                            dr["WDescription"] = Module1.ObjDescription;
                                                        }

                                                        if (!Convert.IsDBNull(dr["WClass Code"]))
                                                        {
                                                            Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                                                            Module1.ObjClassCode.Confidence = 100;
                                                            Module1.ObjClassCode.Flag = "1";
                                                            Module1.ObjClassCode.Remarks = "3F";
                                                            dr["WClass Code"] = Module1.ObjClassCode;

                                                        }


                                                        dtDetailFinal.AcceptChanges();
                                                        continue;

                                                    }
                                                }
                                            }
                                            else
                                            {

                                                dr["IsValid"] = 0;
                                                dtDetailFinal.AcceptChanges();
                                            }
                                        }
                                        else
                                        {
                                            dr["IsValid"] = 0;
                                            dtDetailFinal.AcceptChanges();
                                        }
                                    }
                                }
                                else
                                {
                                    dr["IsValid"] = 0;
                                    dtDetailFinal.AcceptChanges();
                                }

                            }

                            //if ((string.IsNullOrEmpty(Module1.ObjDescription.Flag) || Module1.ObjDescription.Flag == "0") && !string.IsNullOrEmpty(Module1.ObjDescription.strWord))
                            //{
                            //    Module1.ObjDescription.strWord = "";
                            //    Module1.ObjDescription.Confidence = 0;
                            //    dr["WDescription"] = Module1.ObjDescription;
                            //    dr["Description"] = Module1.ObjDescription.strWord;
                            //    dtDetailFinal.AcceptChanges();
                            //}

                            if (Module1.ObjClassCode != null)
                            {
                                if (string.IsNullOrEmpty(Module1.ObjClassCode.Flag) || Module1.ObjClassCode.Flag == "0")
                                {

                                    if (!string.IsNullOrEmpty(Module1.ObjClassCode.strWord))
                                    {
                                        double result;
                                        if (!double.TryParse(Module1.ObjClassCode.strWord, out result))
                                        {

                                            Module1.ObjClassCode.strWord = "";
                                            Module1.ObjClassCode.Confidence = 0;
                                            dr["WClass Code"] = Module1.ObjClassCode;
                                            dr["Class Code"] = Module1.ObjClassCode.strWord;
                                            dtDetailFinal.AcceptChanges();
                                        }

                                        if (Module1.ObjClassCode.strWord.Length > 5)
                                        {

                                            Module1.ObjClassCode.strWord = "";
                                            Module1.ObjClassCode.Confidence = 0;
                                            dr["WClass Code"] = Module1.ObjClassCode;
                                            dr["Class Code"] = Module1.ObjClassCode.strWord;
                                            dtDetailFinal.AcceptChanges();
                                        }
                                    }
                                }
                            }

                            if (Module1.ObjWeight != null)
                            {

                                if (string.IsNullOrEmpty(Module1.ObjWeight.Flag) || Module1.ObjWeight.Flag == "0")
                                {


                                    if (!string.IsNullOrEmpty(Module1.ObjWeight.strWord))
                                    {
                                        if (Module1.ObjWeight.strWord.Length <= 1)
                                        {

                                            Module1.ObjWeight.strWord = "";
                                            Module1.ObjWeight.Confidence = 0;
                                            dr["WWeight"] = Module1.ObjWeight;
                                            dr["Weight"] = Module1.ObjWeight.strWord;
                                            dtDetailFinal.AcceptChanges();
                                        }
                                        string[] arrWeight = Module1.ObjWeight.strWord.Split(' ');

                                        if (arrWeight.Count() == 2)
                                        {
                                            if (!(arrWeight[1].ToLower() == "lbs" || arrWeight[1].ToLower() == "lb"))
                                            {
                                                Module1.ObjWeight.strWord = "";
                                                Module1.ObjWeight.Confidence = 0;
                                                dr["WWeight"] = Module1.ObjWeight;
                                                dr["Weight"] = Module1.ObjWeight.strWord;
                                                dtDetailFinal.AcceptChanges();
                                            }
                                            else
                                            {

                                                Module1.ObjWeight.strWord = "";
                                                Module1.ObjWeight.Confidence = 0;
                                                dr["WWeight"] = Module1.ObjWeight;
                                                dr["Weight"] = Module1.ObjWeight.strWord;
                                                dtDetailFinal.AcceptChanges();
                                            }
                                        }
                                        else if (arrWeight.Count() > 1)
                                        {

                                            Module1.ObjWeight.strWord = "";
                                            Module1.ObjWeight.Confidence = 0;
                                            dr["WWeight"] = Module1.ObjWeight;
                                            dr["Weight"] = Module1.ObjWeight.strWord;
                                            dtDetailFinal.AcceptChanges();
                                        }


                                    }
                                }

                                if (Module1.ObjWeight.strWord == "0" || Convert.ToDouble(Module1.ObjWeight.strWord) == 0.00 || Module1.ObjWeight.strWord == "0.00")
                                {
                                    Module1.ObjWeight.strWord = "";
                                    Module1.ObjWeight.Confidence = 0;
                                    dr["WWeight"] = Module1.ObjWeight;
                                    dr["Weight"] = Module1.ObjWeight.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }
                            if (Module1.ObjQuantity != null)
                            {
                                if (Module1.ObjQuantity.strWord == "0")
                                {
                                    Module1.ObjQuantity.strWord = "";
                                    Module1.ObjQuantity.Confidence = 0;
                                    dr["WQuantity"] = Module1.ObjQuantity;
                                    dr["Quantity"] = Module1.ObjQuantity.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }

                            if (Module1.ObjNMFC != null)
                            {
                                if (string.IsNullOrEmpty(Module1.ObjNMFC.Flag) || Module1.ObjNMFC.Flag == "0")
                                {
                                    if (!string.IsNullOrEmpty(Module1.ObjNMFC.strWord))
                                    {
                                        if (Module1.ObjNMFC.strWord.Length > 12 || Module1.ObjNMFC.strWord.Length < 5)
                                        {

                                            Module1.ObjNMFC.strWord = "";
                                            Module1.ObjNMFC.Confidence = 0;
                                            dr["WNMFC"] = Module1.ObjNMFC;
                                            dr["NMFC"] = Module1.ObjNMFC.strWord;
                                            dtDetailFinal.AcceptChanges();
                                        }

                                        string[] arrNmfc = Module1.ObjNMFC.strWord.Split('-');

                                        if (!arrNmfc[0].All(x => char.IsDigit(x)))
                                        {

                                            Module1.ObjNMFC.strWord = "";
                                            Module1.ObjNMFC.Confidence = 0;
                                            dr["WNMFC"] = Module1.ObjNMFC;
                                            dr["NMFC"] = Module1.ObjNMFC.strWord;
                                            dtDetailFinal.AcceptChanges();

                                        }


                                        if (arrNmfc.Count() == 2)
                                            if (!arrNmfc[1].All(x => char.IsDigit(x)))
                                            {
                                                Module1.ObjNMFC.strWord = "";
                                                Module1.ObjNMFC.Confidence = 0;
                                                dr["WNMFC"] = Module1.ObjNMFC;
                                                dr["NMFC"] = Module1.ObjNMFC.strWord;
                                                dtDetailFinal.AcceptChanges();
                                            }

                                    }
                                }
                            }

                        }
                    }
                    catch
                    {

                    }
                }
            done:

                foreach (DataRow dr in dtDetailFinal.Rows)
                {
                    if (string.IsNullOrEmpty(Convert.ToString(dr["Quantity"])))
                    {
                        DataRow[] dr1 = dtDetailFinal.Select("Quantity is not null");

                        if (dr1.Count() > 0)
                        {
                            Module1.ObjQuantity = (clsCnCWord)dr1[0]["WQuantity"];

                            hx1 = Convert.ToInt32(Module1.ObjQuantity.Left);
                            hy1 = 0;
                            hx2 = Convert.ToInt32(Module1.ObjQuantity.Right);
                            hy2 = 0;
                            myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);

                            foreach (DataRow item in dtDetail.Rows)
                            {

                                if (Convert.ToInt32(item["Line No"]) != Convert.ToInt32(dr["PROLineNo"]))
                                    continue;

                                x1 = Convert.ToInt32(item["left"]);
                                y1 = 0;
                                x2 = Convert.ToInt32(item["right"]);
                                y2 = 1;
                                myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                bool doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                if (hx1 == x1 && hx2 == x2)
                                    continue;

                                if (doesIntersect)
                                {
                                    if (!Convert.IsDBNull(item["ObjWord"]))
                                    {
                                        Module1.ObjWord = (clsCnCWord)item["ObjWord"];

                                        if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                             (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                            Module1.ObjWord.Flag = "1";
                                        else
                                            Module1.ObjWord.Flag = "0";

                                        Module1.ObjWord.Remarks = "3F";
                                        item["ObjWord"] = Module1.ObjWord;
                                        //Module1.ObjWord = null;


                                        dr["Quantity"] = Module1.ObjWord.strWord;
                                        dr["WQuantity"] = item["ObjWord"];
                                        dtDetailFinal.AcceptChanges();
                                        break;
                                    }

                                }

                            }
                        }
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dr["Description"])))
                    {
                        DataRow[] dr1 = dtDetailFinal.Select("Description is not null");

                        if (dr1.Count() > 0)
                        {
                            Module1.ObjDescription = (clsCnCWord)dr1[0]["WDescription"];

                            hx1 = Convert.ToInt32(Module1.ObjDescription.Left);
                            hy1 = 0;
                            hx2 = Convert.ToInt32(Module1.ObjDescription.Right);
                            hy2 = 0;
                            myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);

                            foreach (DataRow item in dtDetail.Rows)
                            {

                                if (Convert.IsDBNull(dr["PROLineNo"]))
                                    continue;

                                if (Convert.ToInt32(item["Line No"]) != Convert.ToInt32(dr["PROLineNo"]))
                                    continue;

                                x1 = Convert.ToInt32(item["left"]);
                                y1 = 0;
                                x2 = Convert.ToInt32(item["right"]);
                                y2 = 1;
                                myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                bool doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                if (hx1 == x1 && hx2 == x2)
                                    continue;

                                if (doesIntersect)
                                {
                                    if (!Convert.IsDBNull(item["ObjWord"]))
                                    {
                                        Module1.ObjWord = (clsCnCWord)item["ObjWord"];

                                        if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                             (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                            Module1.ObjWord.Flag = "1";
                                        else
                                            Module1.ObjWord.Flag = "0";

                                        Module1.ObjWord.Remarks = "3F";
                                        item["ObjWord"] = Module1.ObjWord;
                                        //Module1.ObjWord = null;


                                        dr["Description"] = Module1.ObjWord.strWord;
                                        dr["WDescription"] = item["ObjWord"];
                                        dtDetailFinal.AcceptChanges();
                                        break;
                                    }

                                }
                            }
                        }
                    }
                    if (string.IsNullOrEmpty(Convert.ToString(dr["NMFC"])))
                    {
                        DataRow[] dr1 = dtDetailFinal.Select("NMFC is not null");

                        if (dr1.Count() > 0)
                        {
                            Module1.ObjNMFC = (clsCnCWord)dr1[0]["WNMFC"];

                            hx1 = Convert.ToInt32(Module1.ObjNMFC.Left);
                            hy1 = 0;
                            hx2 = Convert.ToInt32(Module1.ObjNMFC.Right);
                            hy2 = 0;
                            myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);

                            foreach (DataRow item in dtDetail.Rows)
                            {

                                if (Convert.ToInt32(item["Line No"]) != Convert.ToInt32(dr["PROLineNo"]))
                                    continue;

                                x1 = Convert.ToInt32(item["left"]);
                                y1 = 0;
                                x2 = Convert.ToInt32(item["right"]);
                                y2 = 1;
                                myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                bool doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                if (hx1 == x1 && hx2 == x2)
                                    continue;

                                if (doesIntersect)
                                {
                                    if (!Convert.IsDBNull(item["ObjWord"]))
                                    {
                                        Module1.ObjWord = (clsCnCWord)item["ObjWord"];

                                        if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                             (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                            Module1.ObjWord.Flag = "1";
                                        else
                                            Module1.ObjWord.Flag = "0";

                                        Module1.ObjWord.Remarks = "3F";
                                        item["ObjWord"] = Module1.ObjWord;
                                        //Module1.ObjWord = null;


                                        dr["NMFC"] = Module1.ObjWord.strWord;
                                        dr["WNMFC"] = item["ObjWord"];
                                        dtDetailFinal.AcceptChanges();
                                        break;
                                    }

                                }

                            }
                        }
                    }


                    if (string.IsNullOrEmpty(Convert.ToString(dr["Weight"])))
                    {
                        DataRow[] dr1 = dtDetailFinal.Select("Weight is not null");

                        if (dr1.Count() > 0)
                        {
                            Module1.ObjWeight = (clsCnCWord)dr1[0]["WWeight"];

                            hx1 = Convert.ToInt32(Module1.ObjWeight.Left);
                            hy1 = 0;
                            hx2 = Convert.ToInt32(Module1.ObjWeight.Right);
                            hy2 = 0;
                            myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);

                            foreach (DataRow item in dtDetail.Rows)
                            {

                                if (Convert.ToInt32(item["Line No"]) != Convert.ToInt32(dr["PROLineNo"]))
                                    continue;

                                x1 = Convert.ToInt32(item["left"]);
                                y1 = 0;
                                x2 = Convert.ToInt32(item["right"]);
                                y2 = 1;
                                myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                bool doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                if (hx1 == x1 && hx2 == x2)
                                    continue;

                                if (doesIntersect)
                                {
                                    if (!Convert.IsDBNull(item["ObjWord"]))
                                    {
                                        double result;

                                        Module1.ObjWord = (clsCnCWord)item["ObjWord"];

                                        if (double.TryParse(Module1.ObjWord.strWord.ToLower().Replace(".lbs", "").Replace(".lb", "").Replace("lb", "").Replace("lbs", ""), out result))
                                        {
                                            if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                                (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                                Module1.ObjWord.Flag = "1";
                                            else
                                                Module1.ObjWord.Flag = "0";

                                            Module1.ObjWord.Remarks = "3F";
                                            item["ObjWord"] = Module1.ObjWord;
                                            //Module1.ObjWord = null;


                                            dr["Weight"] = Module1.ObjWord.strWord;
                                            dr["WWeight"] = item["ObjWord"];
                                            dtDetailFinal.AcceptChanges();
                                            break;
                                        }
                                    }

                                }
                            }
                        }
                    }

                    if (string.IsNullOrEmpty(Convert.ToString(dr["Class Code"])))
                    {
                        DataRow[] dr1 = dtDetailFinal.Select("[Class Code] is not null");

                        if (dr1.Count() > 0)
                        {
                            Module1.ObjClassCode = (clsCnCWord)dr1[0]["WClass Code"];

                            hx1 = Convert.ToInt32(Module1.ObjClassCode.Left);
                            hy1 = 0;
                            hx2 = Convert.ToInt32(Module1.ObjClassCode.Right);
                            hy2 = 0;
                            myRectangle1 = new Rect(hx1, hy1, hx2 - hx1, hy2);

                            foreach (DataRow item in dtDetail.Rows)
                            {

                                if (Convert.ToInt32(item["Line No"]) != Convert.ToInt32(dr["PROLineNo"]))
                                    continue;

                                x1 = Convert.ToInt32(item["left"]);
                                y1 = 0;
                                x2 = Convert.ToInt32(item["right"]);
                                y2 = 1;
                                myRectangle2 = new Rect(x1, y1, x2 - x1, y2);

                                bool doesIntersect = myRectangle1.IntersectsWith(myRectangle2);

                                if (hx1 == x1 && hx2 == x2)
                                    continue;

                                if (doesIntersect)
                                {
                                    if (!Convert.IsDBNull(item["ObjWord"]))
                                    {
                                        Module1.ObjWord = (clsCnCWord)item["ObjWord"];

                                        if (Module1.ObjWord.Confidence >= 80 && !(Module1.ObjWord.ConfString.Contains("5") || Module1.ObjWord.ConfString.Contains("4") || Module1.ObjWord.ConfString.Contains("6") ||
                                            (Module1.ObjWord.ConfString.Contains("3") || Module1.ObjWord.ConfString.Contains("2") || Module1.ObjWord.ConfString.Contains("1"))))
                                            Module1.ObjWord.Flag = "1";
                                        else
                                            Module1.ObjWord.Flag = "0";

                                        Module1.ObjWord.Remarks = "3F";
                                        item["ObjWord"] = Module1.ObjWord;
                                        //Module1.ObjWord = null;


                                        dr["Class Code"] = Module1.ObjWord.strWord;
                                        dr["WClass Code"] = item["ObjWord"];
                                        dtDetailFinal.AcceptChanges();
                                        break;
                                    }

                                }
                            }
                        }
                    }

                }

                /****For Hazmatcode*****/

                foreach (DataRow dr in dtDetailFinal.Rows)
                {
                    string[] strDescription = Convert.ToString(dr["Description"]).Split(' ');

                    foreach (string item in strDescription)
                    {
                        if (Regex.IsMatch(item, "^UN"))
                        {
                            string str = item.Substring(2);
                            str = RemoveSpecialChars(str);
                            int result;
                            if (int.TryParse(str, out result))
                            {
                                if (str.Length == 4 || str.Length == 2)
                                {
                                    dr["HazMat code"] = "X";
                                    Module1.ObjHazmatCode = GetDefaultWord(1, "X");
                                    Module1.ObjHazmatCode.Confidence = 100;
                                    Module1.ObjHazmatCode.Flag = "1";
                                    Module1.ObjHazmatCode.Remarks = "3F";
                                    dr["WHazMat code"] = Module1.ObjHazmatCode;
                                    dtDetailFinal.AcceptChanges();
                                    break;
                                }
                            }
                        }

                    }



                }



                foreach (DataRow drclass in dtDetailFinal.Rows)
                {
                    string sClassCode = Convert.ToString(drclass["Class Code"]);
                    sClassCode = sClassCode.TrimStart('0');

                    if (sClassCode.Trim() == "92.5")
                    {
                        drclass["Class Code"] = "92";
                        Module1.ObjClassCode = (clsCnCWord)drclass["WClass Code"];
                        Module1.ObjClassCode.strWord = "92";
                        drclass["WClass Code"] = Module1.ObjClassCode;
                    }
                    else if (sClassCode.Trim() == "77.5" || sClassCode.Trim() == "77,5")
                    {
                        drclass["Class Code"] = "77";
                        Module1.ObjClassCode = (clsCnCWord)drclass["WClass Code"];
                        Module1.ObjClassCode.strWord = "77";
                        drclass["WClass Code"] = Module1.ObjClassCode;


                    }
                    else if (sClassCode.Trim() == "70.0")
                    {
                        drclass["Class Code"] = "70";
                        Module1.ObjClassCode = (clsCnCWord)drclass["WClass Code"];
                        Module1.ObjClassCode.strWord = "70";
                        drclass["WClass Code"] = Module1.ObjClassCode;


                    }
                    else if (sClassCode.Trim() == "93" || sClassCode.Trim() == "78")
                    {
                        drclass["Class Code"] = "";
                        Module1.ObjClassCode = (clsCnCWord)drclass["WClass Code"];
                        Module1.ObjClassCode.strWord = "";
                        drclass["WClass Code"] = Module1.ObjClassCode;
                    }
                    else
                    {
                        if (!string.IsNullOrEmpty(sClassCode))
                        {

                            Module1.ObjClassCode = (clsCnCWord)drclass["WClass Code"];
                            Module1.ObjClassCode.strWord = sClassCode;
                            drclass["WClass Code"] = Module1.ObjClassCode;
                        }
                        else
                        {
                            drclass["WClass Code"] = null;
                        }
                    }

                    dtDetailFinal.AcceptChanges();
                }

                //reformat words in table if going blanks

                foreach (DataRow dr in dtDetailFinal.Rows)
                {
                    if (Convert.ToString(dr["NMFC"]).Contains("-"))
                        arrNMFC = Convert.ToString(dr["NMFC"]).Split('-');
                    else if (Convert.ToString(dr["NMFC"]).Contains("."))
                        arrNMFC = Convert.ToString(dr["NMFC"]).Split('.');
                    else if (Convert.ToString(dr["NMFC"]).Contains("S") || Convert.ToString(dr["NMFC"]).Contains("s"))
                        arrNMFC = Convert.ToString(dr["NMFC"]).Split(' ');
                    else
                        arrNMFC = Convert.ToString(dr["NMFC"]).Split('-');

                    strNMFC = string.Empty; strSub = string.Empty; strClassCode = string.Empty;

                    if (arrNMFC.Count() > 1)
                    {
                        strNMFC = arrNMFC[0].Trim().Replace("l", "");

                        if (arrNMFC[1] == "00" || arrNMFC[1] == "0" || arrNMFC[1] == string.Empty)
                        {
                            strSub = "0";
                        }
                        else if (arrNMFC[1].Contains("S") || arrNMFC[1].Contains("s"))
                        {
                            strSub = arrNMFC[1].Substring(1, arrNMFC[1].Length - 1);
                            if (strSub.Length == 1 || strSub.Length == 2)
                            {
                                int val;
                                if (int.TryParse(strSub, out val))
                                {
                                    strSub = strSub.TrimStart('0');
                                }
                                else
                                {
                                    strSub = "0";
                                }
                            }
                            else
                            {
                                strSub = "0";
                            }
                        }
                        else
                            strSub = arrNMFC[1].Replace('0', ' ').Trim();
                    }
                    else
                    {
                        strNMFC = Convert.ToString(dr["NMFC"]).Replace("l", "");
                        strSub = "0";
                    }

                    if (!Convert.IsDBNull(dr["WNMFC"]))
                    {
                        word = (clsCnCWord)dr["WNMFC"];
                        //LineNo = word.LineNo;
                        Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];
                    }
                    else
                    {
                        Module1.ObjNMFC = null;
                    }

                    if (!Convert.IsDBNull(dr["WQuantity"]))
                    {

                        word = (clsCnCWord)dr["WQuantity"];
                        //LineNo = word.LineNo;
                        Module1.ObjQuantity = (clsCnCWord)dr["WQuantity"];
                    }
                    else
                    {
                        Module1.ObjQuantity = null;
                    }

                    if (!Convert.IsDBNull(dr["WClass Code"]))
                    {
                        word = (clsCnCWord)dr["WClass Code"];
                        //LineNo = word.LineNo;
                        Module1.ObjClassCode = (clsCnCWord)dr["WClass Code"];
                    }
                    else
                    {
                        Module1.ObjClassCode = null;
                    }

                    if (!Convert.IsDBNull(dr["WDescription"]))
                    {
                        word = (clsCnCWord)dr["WDescription"];
                        //LineNo = word.LineNo;
                        Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                    }

                    if (!Convert.IsDBNull(dr["WWeight"]))
                    {
                        word = (clsCnCWord)dr["WWeight"];
                        //LineNo = word.LineNo;
                        Module1.ObjWeight = (clsCnCWord)dr["WWeight"];
                    }
                    else
                    {
                        Module1.ObjWeight = null;
                    }

                    if (Module1.ObjDescription != null)
                    {

                        double result;
                        if (double.TryParse(Module1.ObjDescription.strWord, out result) || Module1.ObjDescription.strWord.Contains("WEIGHT") || Module1.ObjDescription.strWord.Contains("TOTA"))
                        {

                            //Module1.ObjDescription.strWord = "";
                            Module1.ObjDescription.Confidence = 0;
                            Module1.ObjDescription.Flag = "0";
                            //Module1.ObjDescription.Remarks = "";
                            dr["WDescription"] = Module1.ObjDescription;
                            dr["Description"] = Module1.ObjDescription.strWord;
                            dtDetailFinal.AcceptChanges();
                        }
                    }


                    if (Module1.ObjClassCode != null)
                    {
                        //    if (string.IsNullOrEmpty(Module1.ObjClassCode.Flag) || Module1.ObjClassCode.Flag == "0")
                        //    {

                        if (!string.IsNullOrEmpty(Module1.ObjClassCode.strWord))
                        {
                            double result;

                            if (Module1.ObjClassCode.strWord.Contains(',') || Module1.ObjClassCode.strWord.Contains('/') || Module1.ObjClassCode.strWord.Contains(':') || Module1.ObjClassCode.strWord.Contains(';') || Module1.ObjClassCode.strWord.Contains('(')
                        || Module1.ObjClassCode.strWord.Contains(')') || Module1.ObjClassCode.strWord.Contains("\"") || Module1.ObjClassCode.strWord.Contains('!'))
                            {
                                Module1.ObjClassCode.Flag = "0";
                            }

                            if (!double.TryParse(Module1.ObjClassCode.strWord, out result))
                            {

                                Module1.ObjClassCode.strWord = "";
                                Module1.ObjClassCode.Confidence = 0;
                                Module1.ObjClassCode.Flag = "0";
                                Module1.ObjClassCode.Remarks = "";
                                dr["WClass Code"] = Module1.ObjClassCode;
                                dr["Class Code"] = Module1.ObjClassCode.strWord;
                                dtDetailFinal.AcceptChanges();
                            }

                            if (Module1.ObjClassCode.strWord.Length > 5)
                            {
                                Module1.ObjClassCode.strWord = "";
                                Module1.ObjClassCode.Confidence = 0;
                                Module1.ObjClassCode.Flag = "0";
                                Module1.ObjClassCode.Remarks = "";
                                dr["WClass Code"] = Module1.ObjClassCode;
                                dr["Class Code"] = Module1.ObjClassCode.strWord;
                                dtDetailFinal.AcceptChanges();
                            }
                        }
                        //}
                    }

                    if (Module1.ObjWeight != null)
                    {

                        //if (string.IsNullOrEmpty(Module1.ObjWeight.Flag) || Module1.ObjWeight.Flag == "0")
                        //{


                        if (!string.IsNullOrEmpty(Module1.ObjWeight.strWord))
                        {
                            //if (Module1.ObjWeight.strWord.Length <= 1)
                            //{

                            //    Module1.ObjWeight.strWord = "";
                            //    Module1.ObjWeight.Confidence = 0;
                            //    Module1.ObjWeight.Flag = "0";
                            //    Module1.ObjWeight.Remarks = "";

                            //    dr["WWeight"] = Module1.ObjWeight;
                            //    dr["Weight"] = Module1.ObjWeight.strWord;
                            //    dtDetailFinal.AcceptChanges();
                            //}
                            string[] arrWeight = Module1.ObjWeight.strWord.Split(' ');

                            if (arrWeight.Count() == 2)
                            {
                                if (!(arrWeight[1].ToLower() == "lbs" || arrWeight[1].ToLower() == "lb"))
                                {
                                    Module1.ObjWeight.strWord = "";
                                    Module1.ObjWeight.Confidence = 0;
                                    Module1.ObjWeight.Flag = "0";
                                    Module1.ObjWeight.Remarks = "";
                                    dr["WWeight"] = Module1.ObjWeight;
                                    dr["Weight"] = Module1.ObjWeight.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }
                                else
                                {

                                    Module1.ObjWeight.strWord = "";
                                    Module1.ObjWeight.Confidence = 0;
                                    Module1.ObjWeight.Flag = "0";
                                    Module1.ObjWeight.Remarks = "";

                                    dr["WWeight"] = Module1.ObjWeight;
                                    dr["Weight"] = Module1.ObjWeight.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }
                            else if (arrWeight.Count() > 1)
                            {

                                Module1.ObjWeight.strWord = "";
                                Module1.ObjWeight.Confidence = 0;
                                Module1.ObjWeight.Flag = "0";
                                Module1.ObjWeight.Remarks = "";

                                dr["WWeight"] = Module1.ObjWeight;
                                dr["Weight"] = Module1.ObjWeight.strWord;
                                dtDetailFinal.AcceptChanges();
                            }



                            //}

                            if (Module1.ObjWeight.strWord == "0" || Module1.ObjWeight.strWord == "0.00")
                            {
                                Module1.ObjWeight.strWord = "";
                                Module1.ObjWeight.Confidence = 0;
                                Module1.ObjWeight.Flag = "0";
                                Module1.ObjWeight.Remarks = "";

                                dr["WWeight"] = Module1.ObjWeight;
                                dr["Weight"] = Module1.ObjWeight.strWord;
                                dtDetailFinal.AcceptChanges();
                            }

                            double result;
                            if (!double.TryParse(Module1.ObjWeight.strWord, out result))
                            {
                                Module1.ObjWeight.strWord = "";
                                Module1.ObjWeight.Confidence = 0;
                                Module1.ObjWeight.Flag = "0";
                                Module1.ObjWeight.Remarks = "";

                                dr["WWeight"] = Module1.ObjWeight;
                                dr["Weight"] = Module1.ObjWeight.strWord;
                                dtDetailFinal.AcceptChanges();
                            }
                        }
                    }
                    if (Module1.ObjQuantity != null)
                    {
                        if (!string.IsNullOrEmpty(Module1.ObjQuantity.strWord))
                        {
                            if (Module1.ObjQuantity.strWord == "0" || Module1.ObjQuantity.strWord.Length >= 4)
                            {
                                Module1.ObjQuantity.strWord = "";
                                Module1.ObjQuantity.Confidence = 0;
                                Module1.ObjQuantity.Flag = "0";
                                Module1.ObjQuantity.Remarks = "";

                                dr["WQuantity"] = Module1.ObjQuantity;
                                dr["Quantity"] = Module1.ObjQuantity.strWord;
                                dtDetailFinal.AcceptChanges();
                            }


                            double result;
                            if (!double.TryParse(Module1.ObjQuantity.strWord, out result))
                            {
                                Module1.ObjQuantity.strWord = "";
                                Module1.ObjQuantity.Confidence = 0;
                                Module1.ObjQuantity.Flag = "0";
                                Module1.ObjQuantity.Remarks = "";

                                dr["WQuantity"] = Module1.ObjQuantity;
                                dr["Quantity"] = Module1.ObjQuantity.strWord;
                                dtDetailFinal.AcceptChanges();
                            }
                            else
                            {
                                if (result > 200)
                                {
                                    Module1.ObjQuantity.strWord = "";
                                    Module1.ObjQuantity.Confidence = 0;
                                    Module1.ObjQuantity.Flag = "0";
                                    Module1.ObjQuantity.Remarks = "";

                                    dr["WQuantity"] = Module1.ObjQuantity;
                                    dr["Quantity"] = Module1.ObjQuantity.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }


                        }
                    }

                    if (Module1.ObjNMFC != null)
                    {
                        //if (string.IsNullOrEmpty(Module1.ObjNMFC.Flag) || Module1.ObjNMFC.Flag == "0")
                        //{

                        if (!string.IsNullOrEmpty(Module1.ObjNMFC.strWord))
                        {
                            string[] arrNmfc = Module1.ObjNMFC.strWord.Split('-');

                            if (arrNmfc.Count() == 1)
                            {
                                arrNmfc = Module1.ObjNMFC.strWord.Split(' ');
                            }

                            if (arrNmfc.Count() == 1)
                            {
                                arrNmfc = Module1.ObjNMFC.strWord.Split('.');
                            }

                            if (!string.IsNullOrEmpty(Module1.ObjNMFC.strWord))
                            {
                                if (arrNMFC[0].Length > 12 || arrNMFC[0].Length < 5)
                                {

                                    Module1.ObjNMFC.strWord = "";
                                    Module1.ObjNMFC.Confidence = 0;
                                    Module1.ObjNMFC.Flag = "0";
                                    Module1.ObjNMFC.Remarks = "";

                                    dr["WNMFC"] = Module1.ObjNMFC;
                                    dr["NMFC"] = Module1.ObjNMFC.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }


                                if (!arrNmfc[0].All(x => char.IsDigit(x)))
                                {

                                    Module1.ObjNMFC.strWord = "";
                                    Module1.ObjNMFC.Confidence = 0;
                                    Module1.ObjNMFC.Flag = "0";
                                    Module1.ObjNMFC.Remarks = "";
                                    dr["WNMFC"] = Module1.ObjNMFC;
                                    dr["NMFC"] = Module1.ObjNMFC.strWord;
                                    dtDetailFinal.AcceptChanges();

                                }


                                if (arrNmfc.Count() == 2)
                                {
                                    if (!arrNmfc[1].ToUpper().Replace("SUB", "").Replace("S", "").Trim().All(x => char.IsDigit(x)))
                                    {
                                        Module1.ObjNMFC.strWord = "";
                                        Module1.ObjNMFC.Confidence = 0;
                                        Module1.ObjNMFC.Flag = "0";
                                        Module1.ObjNMFC.Remarks = "";
                                        dr["WNMFC"] = Module1.ObjNMFC;
                                        dr["NMFC"] = Module1.ObjNMFC.strWord;
                                        dtDetailFinal.AcceptChanges();
                                    }
                                }
                                else if (arrNmfc.Count() > 2)
                                {

                                    Module1.ObjNMFC.strWord = "";
                                    Module1.ObjNMFC.Confidence = 0;
                                    Module1.ObjNMFC.Flag = "0";
                                    Module1.ObjNMFC.Remarks = "";
                                    dr["WNMFC"] = Module1.ObjNMFC;
                                    dr["NMFC"] = Module1.ObjNMFC.strWord;
                                    dtDetailFinal.AcceptChanges();
                                }
                            }
                        }
                    }
                }
                foreach (DataRow dr in dtDetailFinal.Rows)
                {
                    string sWeight = Convert.ToString(dr["Weight"]);

                    if (!string.IsNullOrEmpty(sWeight))
                    {
                        Module1.ObjWeight = (clsCnCWord)dr["WWeight"];

                        if (sWeight.Contains(',') || sWeight.Contains('/') || sWeight.Contains(':') || sWeight.Contains(';') || sWeight.Contains('(')
                    || sWeight.Contains(')') || sWeight.Contains("\"") || sWeight.Contains('!'))
                        {

                            Module1.ObjWeight.Flag = "0";

                        }

                        Module1.ObjWeight.strWord = sWeight;


                        dr["WWeight"] = Module1.ObjWeight;
                    }
                    else
                    {
                        dr["WWeight"] = null;
                    }

                    string sDescription = Convert.ToString(dr["Description"]);

                    if (!string.IsNullOrEmpty(sDescription))
                    {
                        Module1.ObjDescription = (clsCnCWord)dr["WDescription"];
                        Module1.ObjDescription.strWord = sDescription;
                        dr["WDescription"] = Module1.ObjDescription;
                    }
                    else
                    {
                        dr["WDescription"] = null;
                    }

                    string sNMFC = Convert.ToString(dr["NMFC"]);

                    if (!string.IsNullOrEmpty(sNMFC))
                    {
                        Module1.ObjNMFC = (clsCnCWord)dr["WNMFC"];

                        //if (sNMFC.Contains('.') || sNMFC.Contains(',') || sNMFC.Contains('/') || sNMFC.Contains(':') || sNMFC.Contains(';') || sNMFC.Contains('(')
                        if (sNMFC.Contains(',') || sNMFC.Contains('/') || sNMFC.Contains(':') || sNMFC.Contains(';') || sNMFC.Contains('(')
                    || sNMFC.Contains(')') || sNMFC.Contains("\"") || sNMFC.Contains('!'))
                        {

                            Module1.ObjNMFC.Flag = "0";

                        }
                        Module1.ObjNMFC.strWord = sNMFC;
                        dr["WNMFC"] = Module1.ObjNMFC;
                    }
                    else
                    {
                        dr["WNMFC"] = null;
                    }


                    string sQuantity = Convert.ToString(dr["Quantity"]);

                    if (!string.IsNullOrEmpty(sQuantity))
                    {
                        Module1.ObjQuantity = (clsCnCWord)dr["WQuantity"];

                        if (sQuantity.Contains('.') || sQuantity.Contains(',') || sQuantity.Contains('/') || sQuantity.Contains(':') || sQuantity.Contains(';') || sQuantity.Contains('(')
                    || sQuantity.Contains(')') || sQuantity.Contains("\"") || sQuantity.Contains('!'))
                        {

                            Module1.ObjQuantity.Flag = "0";

                        }
                        Module1.ObjQuantity.strWord = sQuantity;
                        dr["WQuantity"] = Module1.ObjQuantity;
                    }
                    else
                    {
                        dr["WQuantity"] = null;
                    }


                    string sPalletCode = Convert.ToString(dr["Pallate Code"]);

                    if (!string.IsNullOrEmpty(sPalletCode))
                    {
                        Module1.ObjPalletCode = (clsCnCWord)dr["WPallate Code"];
                        Module1.ObjPalletCode.strWord = sPalletCode;
                        dr["WPallate Code"] = Module1.ObjPalletCode;

                    }
                    else
                    {
                        dr["WPallate Code"] = null;
                    }

                    string sPROLineNo = Convert.ToString(dr["PROLineNo"]);
                    if (!string.IsNullOrEmpty(sPROLineNo))
                    {
                        dr["PROLineNo"] = sPROLineNo.PadLeft(2, '0');
                    }

                    dtDetailFinal.AcceptChanges();
                }

                if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_LineItemValue") == true)
                {
                    PublicComponents.htMyVariable.Remove("BOLSET4_LineItemValue");
                }

                PublicComponents.htMyVariable.Add("BOLSET4_LineItemValue", dtDetailFinal);

                //watch.Stop();

                //MessageBox.Show(watch.ElapsedMilliseconds.ToString());

                //Form1 f = new Form1(dtDetailFinal, dtGetData);
                //f.Size = new System.Drawing.Size(1000, 500);




                //f.StartPosition = FormStartPosition.WindowsDefaultBounds;
                //f.SetDesktopLocation(1000, 1000);
                //f.ShowDialog();
            }
            catch (Exception ex)
            {
                //  MessageBox.Show("SET4:F3:Line1: " + ex.StackTrace);

            }
            finally
            {


                if (dtHeader != null)
                {
                    dtHeader.Dispose();
                }

                if (dtDetail != null)
                {
                    dtDetail.Dispose();
                }

                if (dtMapTable != null)
                {
                    dtMapTable.Dispose();
                }

                if (dtDetailFinal != null)
                {
                    dtDetailFinal.Dispose();
                }

                strNMFC = null; strSub = null; strClassCode = null;


                drdata = null;
                drdata1 = null;
                drdata3 = null;
                drdata4 = null;

                Module1.ObjDescription = null;
                Module1.ObjNMFC = null;
                Module1.ObjClassCode = null;
                Module1.ObjWeight = null;
                Module1.ObjQuantity = null;
                Module1.ObjHazmatCode = null;
                Module1.ObjPalletCode = null;
                Module1.ObjWord = null;
                drmap = null;
                word = null;
                arrNMFC = null;
                Pkgcode = null;
                code = null;
                //int PreviousLineno = 0;

                //int hx1, hy1, hx2, hy2;
                //int x1, y1, x2, y2;



                //seperator = null;

                //batchname = null;

                //file = null;




                //seperator1 =  null;

                //connStr = null;

                //arrFilename = null;

                // PROFilename=null;



                GC.Collect();

            }

            return Module1.objRetStructF3;
        }

        public string RemoveSpecialChars(string str)
        {
            // Create  a string array and add the special characters you want to remove

            string[] chars = new string[] { ",", ".", "/", "!", "@", "#", "$", "%", "^", "&", "*", "'", "\"", ";", "_", "(", ")", ":", "|", "[", "]" };
            //Iterate the number of times based on the String array length.
            for (int i = 0; i < chars.Length; i++)
            {
                if (str.Contains(chars[i]))
                {
                    str = str.Replace(chars[i], "");
                }
            }

            return str;
        }

        public static DataTable GetDataTableByText_1(string connectionString, string queryString, OdbcParameter[] parameters = null)
        {
            DataTable dataTable = new DataTable();
            try
            {
                using (OdbcConnection connection = new OdbcConnection(connectionString))
                {
                    using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                    {
                        if (connection != null && !string.IsNullOrEmpty(queryString))
                        {
                            cmd.CommandType = CommandType.Text;
                            if (parameters != null)
                            {
                                cmd.Parameters.AddRange(parameters);
                            }
                            using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                            {
                                da.Fill(dataTable);
                            }
                        }
                    }
                }
                return dataTable;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return null;
            }
        }

        public static DataTable GetDataTableByText(string connectionString, string queryString, OdbcParameter[] parameters = null)
        {
            DataTable dataTable = new DataTable();

            try
            {
                //using (OdbcConnection connection = new OdbcConnection(connectionString))
                //{​​​​​
                //connection.Open();
                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();

                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == true)
                    {
                        connectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                    }
                    else
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();
                    }

                    db2Connection.ConnectionString = connectionString;

                    if (db2Connection.State != ConnectionState.Open)
                    {
                        db2Connection.Open();
                        iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", db2Connection);
                    }
                }
                OdbcConnection connection = (OdbcConnection)iQPUBLIC.PublicComponents.htMyVariable["DB2Conn"];
                if (connection.State != ConnectionState.Open)
                {
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == true)
                    {
                        connectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                    }
                    else
                    {
                        connectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                    }

                    connection.ConnectionString = connectionString;
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }


                using (OdbcCommand cmd = new OdbcCommand(queryString, connection))
                {
                    if (connection != null && !string.IsNullOrEmpty(queryString))
                    {
                        cmd.CommandType = CommandType.Text;
                        if (parameters != null)
                        {
                            cmd.Parameters.AddRange(parameters);
                        }
                        using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                        {
                            da.Fill(dataTable);
                        }
                    }
                }
                //}​​​​​
                return dataTable;
            }
            catch (Exception ex)
            {
                //MessageBox.Show(ex.Message);
                return null;
            }
        }

        private RetStructF3 F3_AutoLocate_LineItem(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strArg)
        {
            RetStructF3 objRetStructF3 = new RetStructF3();
            objRetStructF3.Status = "F";
            objRetStructF3.NoOfieldsSuspects = 0;


            try
            {

                objRetStructF3 = Header(ObjMetaData, intCurrPageNumber, strArg);

            }


            catch (Exception ex)
            {
                //MessageBox.Show("F3: F3_AutoLocate_BillOfLading " + ex.Message);
            }


            return objRetStructF3;
        }


        #region BOLSET4
        //DataTable dtKeyword = new DataTable();
        DataTable dtDetailLine = new DataTable();
        DataTable dtDetailLineData = new DataTable();
        DataTable dtDetail = new DataTable();

        string LineNo = "0";


        public RetStructF3 Header(clsCncMetaData ObjMetaData, int intCurrPageNumber, string stringArgs)
        {
            RetStructF3 objRetStructF3 = new RetStructF3();
            objRetStructF3.Status = "F";
            objRetStructF3.NoOfieldsSuspects = 0;

            List<int> ConfLevel = new List<int>();

            //int h = ObjMetaData.Page[intCurrPageNumber].ImageHeight;
            //int w = ObjMetaData.Page[intCurrPageNumber].ImageWidth;
            //int Wid = w;
            //int Height = h / 2;

            List<clsCnCWord> possibleword = new List<clsCnCWord>();
            List<clsCnCWord> Finalpossibleword = new List<clsCnCWord>();
            //MakedtDetailFinal();

            try
            {
                string LineNo = "0";
                //for (intCurrPageNumber = 1; intCurrPageNumber < ObjMetaData.PageCount; intCurrPageNumber++)
                //possibleword = F3_GetValue(ObjMetaData, intCurrPageNumber, Synonyms, 0, 0, Wid, h);
                LineNo = F3_GetValue(ObjMetaData, intCurrPageNumber, stringArgs);
                //F3_GetDetailValue(ObjMetaData, intCurrPageNumber, Synonyms, LineNo);
                clsCnCLine line = new clsCnCLine();

                if (!string.IsNullOrEmpty(LineNo))
                {
                    ClsCNC Objcnc = new ClsCNC();
                    clsCnCWord word = Objcnc.MergeWords(ObjMetaData.Page[intCurrPageNumber].Line[Convert.ToInt32(LineNo)].Word);
                    Finalpossibleword.Add(word);
                }

                List<int> conflevel = new List<int>();
                conflevel.Clear();

                if (!(Finalpossibleword == null))
                {

                    if (Finalpossibleword.Count == 0)
                    {
                        conflevel.Clear();
                        Finalpossibleword.Add(GetDefaultWord(1, " "));
                        conflevel.Add(90);

                        objRetStructF3.NoOfieldsSuspects = Finalpossibleword.Count;
                        objRetStructF3.Words = Finalpossibleword;
                        objRetStructF3.ConfidenceLevelofSuspect = conflevel;
                        objRetStructF3.Status = "F";
                        //objRetStructF3.Flag = "Y";
                        //objRetStructF3.ManualConfirmation = "N";
                        return objRetStructF3;

                    }
                    for (int i = 0; i <= Finalpossibleword.Count - 1; i++)
                    {
                        conflevel.Add(90);

                    }
                    objRetStructF3.NoOfieldsSuspects = Finalpossibleword.Count;
                    objRetStructF3.Words = Finalpossibleword;
                    objRetStructF3.ConfidenceLevelofSuspect = conflevel;
                    objRetStructF3.Status = "S";
                    objRetStructF3.Flag = "Y";
                    objRetStructF3.ManualConfirmation = "N";


                    //if (Finalpossibleword.Count > 1)
                    //    objRetStructF3.ManualConfirmation = "Y";
                }


            }
            catch (Exception ex)
            {
                //Error.Log File
            }
            finally
            {
                GC.Collect();
            }

            return objRetStructF3;
        }

        //public DataTable GetHeaderLine(clsCncMetaData ObjMetaData, int intCurrPageNumber, string StrArgs)
        //{

        //    int h = ObjMetaData.Page[intCurrPageNumber].ImageHeight;
        //    int w = ObjMetaData.Page[intCurrPageNumber].ImageWidth;
        //    int Wid = w;
        //    int Height = h / 2;
        //    DataTable H;
        //    //Synonyms = "Quantity~Weight~Cl~Item Description~Commodity Description~Class~NMFC~Qty~Pieces~Units~Qty~CTNS~Pieces/Quantity~Package~WT~Pkgs~Net Weight~Gross Weight";
        //    //intCurrPageNumber = 1;
        //    MakeDtDeatilLineData();

        //    LineNo = F3_GetValue(ObjMetaData, intCurrPageNumber, StrArgs);

        //    int lineno = Convert.ToInt32(LineNo);

        //    if (lineno != 0)
        //    {

        //        //foreach (clsCnCWord word in ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word)
        //        //{
        //        //    if (word != null)
        //        //    {
        //        //        DataRow dr = dtDetailLineData.NewRow();

        //        //        dr["Page No"] = intCurrPageNumber;
        //        //        dr["Line No"] = lineno;
        //        //        dr["Word No"] = word.WordNumber;
        //        //        dr["Word"] = word.strWord;
        //        //        dr["Left"] = word.Left;
        //        //        dr["Top"] = word.Top;
        //        //        dr["Right"] = word.Right;
        //        //        dr["Bottom"] = word.Bottom;
        //        //        dtDetailLineData.Rows.Add(dr);

        //        //    }
        //        //}

        //        ClsCNC oClsCnc = new ClsCNC();

        //        List<clsCnCLine> Lines = new List<clsCnCLine>();

        //        List<int> Linenos = new List<int>();
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno + 1]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno + 2]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno + 3]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno + 4]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno - 1]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno - 2]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno - 3]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno - 4]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno - 5]);
        //        Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[lineno - 6]);

        //        string[] args = Synonyms.Split(new char[] { '~' });

        //        for (int i = 0; i <= Lines.Count - 1; i++)
        //        {
        //            int matchword = 0;
        //            if (Lines[i] != null)
        //            {
        //                //Possiblewords = new List<clsCnCWord>();

        //                for (int j = 0; j <= args.Length - 1; j++)
        //                {
        //                    for (int k = 0; k <= Lines[i].WordCount; k++)
        //                    {
        //                        //    if (Lines[i].WordCount <=3)
        //                        //        break;

        //                        if (Lines[i].Word[k] != null)
        //                        {
        //                            if (CheckValidWord(Lines[i].Word[k].strWord.ToUpper().Trim().Replace("'", "`"), new String[] { args[j].ToUpper() }) || Lines[i].Word[k].strWord.ToUpper().Trim().Contains(args[j].ToUpper()))
        //                            {
        //                                matchword++;
        //                            }
        //                            else
        //                            {
        //                                string[] words = Lines[i].Word[k].strWord.Split(' ');

        //                                foreach (string item in words)
        //                                {
        //                                    if (CheckValidWord(item.ToUpper().Trim().Replace("'", "`"), new String[] { args[j].ToUpper() }) || Lines[i].Word[k].strWord.ToUpper().Trim().Contains(args[j].ToUpper()))
        //                                    {
        //                                        matchword++; break;

        //                                    }
        //                                }
        //                            }
        //                        }
        //                    }


        //                }

        //                if (matchword >= 1)
        //                {
        //                    Linenos.Add(Lines[i].LineNo);
        //                    continue;
        //                }
        //            }
        //        }


        //        foreach (int Lno in Linenos)
        //        {

        //            foreach (clsCnCWord word in ObjMetaData.Page[intCurrPageNumber].Line[Lno].Word)
        //            {
        //                if (word != null)
        //                {
        //                    if (word.strWord.ToUpper().Trim().Replace("'", "`").Contains("UPC"))
        //                        continue;

        //                    if (word.strWord.ToUpper().Trim().Replace("'", "`").Contains("DESCRIPTION") ||     word.strWord.ToUpper().Trim().Replace("'", "`").Contains("OTY")|| 
        //                            word.strWord.ToUpper().Trim().Replace("'", "`").Contains("QTY") || word.strWord.ToUpper().Trim().Replace("'", "`").Contains("TYPE"))
        //                    {

        //                        DataRow dr = dtDetailLineData.NewRow();

        //                        dr["Page No"] = intCurrPageNumber;
        //                        dr["Line No"] = word.LineNo;
        //                        dr["Word No"] = word.WordNumber;
        //                        dr["Word"] = word.strWord;
        //                        dr["Left"] = word.Left;
        //                        dr["Top"] = word.Top;
        //                        dr["Right"] = word.Right;
        //                        dr["Bottom"] = word.Bottom;
        //                        dtDetailLineData.Rows.Add(dr);
        //                        continue;
        //                    }
        //                    //if (word.strWord.ToUpper().Trim().Replace("'", "`").Contains("NET WEIGHT") || word.strWord.ToUpper().Trim().Replace("'", "`").Contains("H.UNITS"))
        //                    //    continue;

        //                    for (int j = 0; j <= args.Length - 1; j++)
        //                    {

        //                        if (CheckValidWord(word.strWord.ToUpper().Trim().Replace("'", "`"), new String[] { args[j] }) || word.strWord.ToUpper().Trim().Contains(args[j].ToUpper()))
        //                        {

        //                            if (word.strWord.Length - args[j].Length <= 4 || args[j].ToUpper() == "DESCRIPTION" || args[j].ToUpper() == "DESC")
        //                            {
        //                                string filter = "Word='" + word.strWord + "'";

        //                                DataView view = new DataView(dtDetailLineData);
        //                                view.RowFilter = filter;
        //                                DataTable dtResult = view.ToTable();

        //                                if (dtResult.Rows.Count == 0)
        //                                {
        //                                    DataRow dr = dtDetailLineData.NewRow();

        //                                    dr["Page No"] = intCurrPageNumber;
        //                                    dr["Line No"] = word.LineNo;
        //                                    dr["Word No"] = word.WordNumber;
        //                                    dr["Word"] = word.strWord;
        //                                    dr["Left"] = word.Left;
        //                                    dr["Top"] = word.Top;
        //                                    dr["Right"] = word.Right;
        //                                    dr["Bottom"] = word.Bottom;
        //                                    dtDetailLineData.Rows.Add(dr);
        //                                }
        //                            }
        //                        }
        //                        else
        //                        {
        //                            string[] words = word.strWord.Split(' ');

        //                            foreach (string item in words)
        //                            {
        //                                if (CheckValidWord(item.ToUpper().Trim().Replace("'", "`"), new String[] { args[j] }) || item.ToUpper().Trim().Contains(args[j].ToUpper()))
        //                                {

        //                                    if (word.strWord.Length - args[j].Length <= 4 || args[j].ToUpper() == "DESCRIPTION" || args[j].ToUpper() == "DESC")
        //                                    {
        //                                        string filter = "Word='" + word.strWord + "'";

        //                                        DataView view = new DataView(dtDetailLineData);
        //                                        view.RowFilter = filter;
        //                                        DataTable dtResult = view.ToTable();

        //                                        if (dtResult.Rows.Count == 0)
        //                                        {
        //                                            DataRow dr = dtDetailLineData.NewRow();

        //                                            dr["Page No"] = intCurrPageNumber;
        //                                            dr["Line No"] = word.LineNo;
        //                                            dr["Word No"] = word.WordNumber;
        //                                            dr["Word"] = word.strWord;
        //                                            dr["Left"] = word.Left;
        //                                            dr["Top"] = word.Top;
        //                                            dr["Right"] = word.Right;
        //                                            dr["Bottom"] = word.Bottom;
        //                                            dtDetailLineData.Rows.Add(dr);
        //                                        }
        //                                    }
        //                                }
        //                            }

        //                        }


        //                    }
        //                }
        //            }

        //        }
        //    }

        //    return dtDetailLineData;
        //}
        public DataTable GetHeaderLine(clsCncMetaData ObjMetaData, int intCurrPageNumber, string StrArgs)
        {

            //int h = ObjMetaData.Page[intCurrPageNumber].ImageHeight;
            //int w = ObjMetaData.Page[intCurrPageNumber].ImageWidth;
            //int Wid = w;
            //int Height = h / 2;

            MakeDtDeatilLineData();

            LineNo = F3_GetValue(ObjMetaData, intCurrPageNumber, StrArgs);

            Module1.LineNo = Convert.ToInt32(LineNo);

            if (Module1.LineNo != 0)
            {

                //foreach (clsCnCWord word in ObjMetaData.Page[intCurrPageNumber].Line[lineno].Word)
                //{
                //    if (word != null)
                //    {
                //        DataRow dr = dtDetailLineData.NewRow();

                //        dr["Page No"] = intCurrPageNumber;
                //        dr["Line No"] = lineno;
                //        dr["Word No"] = word.WordNumber;
                //        dr["Word"] = word.strWord;
                //        dr["Left"] = word.Left;
                //        dr["Top"] = word.Top;
                //        dr["Right"] = word.Right;
                //        dr["Bottom"] = word.Bottom;
                //        dtDetailLineData.Rows.Add(dr);

                //    }
                //}

                ClsCNC oClsCnc = new ClsCNC();

                List<clsCnCLine> Lines = new List<clsCnCLine>();

                List<int> Linenos = new List<int>();
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo + 1]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo + 2]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo + 3]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo + 4]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo - 1]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo - 2]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo - 3]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo - 4]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo - 5]);
                Lines.Add(ObjMetaData.Page[intCurrPageNumber].Line[Module1.LineNo - 6]);

                string[] args = Module1.Synonyms.Split(new char[] { '~' });

                for (int i = 0; i <= Lines.Count - 1; i++)
                {
                    int matchword = 0;
                    if (Lines[i] != null)
                    {
                        //Possiblewords = new List<clsCnCWord>();

                        for (int j = 0; j <= args.Length - 1; j++)
                        {
                            for (int k = 0; k <= Lines[i].WordCount; k++)
                            {
                                //    if (Lines[i].WordCount <=3)
                                //        break;

                                if (Lines[i].Word[k] != null)
                                {
                                    if (CheckValidWord(Lines[i].Word[k].strWord.ToUpper().Trim().Replace("'", "`"), new String[] { args[j].ToUpper() }) || Lines[i].Word[k].strWord.ToUpper().Trim().Contains(args[j].ToUpper()))
                                    {
                                        matchword++;
                                    }
                                    else
                                    {
                                        string[] words = Lines[i].Word[k].strWord.Split(' ');

                                        foreach (string item in words)
                                        {
                                            if (CheckValidWord(item.ToUpper().Trim().Replace("'", "`"), new String[] { args[j].ToUpper() }) || Lines[i].Word[k].strWord.ToUpper().Trim().Contains(args[j].ToUpper()))
                                            {
                                                matchword++; break;

                                            }
                                        }
                                    }
                                }
                            }


                        }

                        if (matchword >= 1)
                        {
                            Linenos.Add(Lines[i].LineNo);
                            continue;
                        }
                    }
                }


                foreach (int Lno in Linenos)
                {

                    foreach (clsCnCWord word in ObjMetaData.Page[intCurrPageNumber].Line[Lno].Word)
                    {
                        if (word != null)
                        {
                            if (word.strWord.ToUpper().Trim().Replace("'", "`").Contains("UPC"))
                                continue;

                            if (word.strWord.ToUpper().Trim().Replace("'", "`").Contains("DESCRIPTION") || word.strWord.ToUpper().Trim().Replace("'", "`").Contains("OTY") ||
                                    word.strWord.ToUpper().Trim().Replace("'", "`").Contains("QTY") || word.strWord.ToUpper().Trim().Replace("'", "`").Contains("TYPE"))
                            {

                                DataRow dr = dtDetailLineData.NewRow();

                                dr["Page No"] = intCurrPageNumber;
                                dr["Line No"] = word.LineNo;
                                dr["Word No"] = word.WordNumber;
                                dr["Word"] = word.strWord;
                                dr["Left"] = word.Left;
                                dr["Top"] = word.Top;
                                dr["Right"] = word.Right;
                                dr["Bottom"] = word.Bottom;
                                dtDetailLineData.Rows.Add(dr);
                                continue;
                            }
                            //if (word.strWord.ToUpper().Trim().Replace("'", "`").Contains("NET WEIGHT") || word.strWord.ToUpper().Trim().Replace("'", "`").Contains("H.UNITS"))
                            //    continue;

                            for (int j = 0; j <= args.Length - 1; j++)
                            {

                                if (CheckValidWord(word.strWord.ToUpper().Trim().Replace("'", "`"), new String[] { args[j] }) || word.strWord.ToUpper().Trim().Contains(args[j].ToUpper()))
                                {

                                    if (word.strWord.Length - args[j].Length <= 4 || args[j].ToUpper() == "DESCRIPTION" || args[j].ToUpper() == "DESC")
                                    {
                                        string value = word.strWord.Replace("'", "");
                                        string filter = "Word='" + value + "'";

                                        DataView view = new DataView(dtDetailLineData);
                                        view.RowFilter = filter;
                                        DataTable dtResult = view.ToTable();

                                        if (dtResult.Rows.Count == 0)
                                        {
                                            DataRow dr = dtDetailLineData.NewRow();

                                            dr["Page No"] = intCurrPageNumber;
                                            dr["Line No"] = word.LineNo;
                                            dr["Word No"] = word.WordNumber;
                                            dr["Word"] = word.strWord;
                                            dr["Left"] = word.Left;
                                            dr["Top"] = word.Top;
                                            dr["Right"] = word.Right;
                                            dr["Bottom"] = word.Bottom;
                                            dtDetailLineData.Rows.Add(dr);
                                        }
                                    }
                                }
                                else
                                {
                                    string[] words = word.strWord.Split(' ');

                                    foreach (string item in words)
                                    {
                                        if (CheckValidWord(item.ToUpper().Trim().Replace("'", "`"), new String[] { args[j] }) || item.ToUpper().Trim().Contains(args[j].ToUpper()))
                                        {

                                            if (word.strWord.Length - args[j].Length <= 4 || args[j].ToUpper() == "DESCRIPTION" || args[j].ToUpper() == "DESC")
                                            {
                                                string filter = "Word='" + word.strWord + "'";

                                                DataView view = new DataView(dtDetailLineData);
                                                view.RowFilter = filter;
                                                DataTable dtResult = view.ToTable();

                                                if (dtResult.Rows.Count == 0)
                                                {
                                                    DataRow dr = dtDetailLineData.NewRow();

                                                    dr["Page No"] = intCurrPageNumber;
                                                    dr["Line No"] = word.LineNo;
                                                    dr["Word No"] = word.WordNumber;
                                                    dr["Word"] = word.strWord;
                                                    dr["Left"] = word.Left;
                                                    dr["Top"] = word.Top;
                                                    dr["Right"] = word.Right;
                                                    dr["Bottom"] = word.Bottom;
                                                    dtDetailLineData.Rows.Add(dr);
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

            return dtDetailLineData;
        }

        public DataTable GetDataLineDeail(clsCncMetaData ObjMetaData, int intCurrPageNumber, string strargs)
        {
            int h = ObjMetaData.Page[intCurrPageNumber].ImageHeight;
            int w = ObjMetaData.Page[intCurrPageNumber].ImageWidth;
            int Wid = w;
            int Height = h / 2;
            DataTable dtDetail;


            try
            {


                //possibleword = F3_GetValue(ObjMetaData, intCurrPageNumber, Synonyms, 0, 0, Wid, h);

                //LineNo = F3_GetValue(ObjMetaData, intCurrPageNumber, strargs);

                dtDetail = F3_GetDetailValue(ObjMetaData, intCurrPageNumber, "0", Module1.LineNo);

                MakeDtDeatilLineData();

                if (dtDetail != null)
                {
                    if (dtDetail.Rows.Count > 0)
                    {
                        foreach (DataRow item in dtDetail.Rows)
                        {
                            int DetailLineNo = Convert.ToInt32(item["Line No"]);
                            int DetailPageNo = Convert.ToInt32(item["Page No"]);


                            if (DetailLineNo != 0)
                            {
                                foreach (clsCnCWord word in ObjMetaData.Page[DetailPageNo].Line[DetailLineNo].Word)
                                {
                                    if (word != null)
                                    {
                                        DataRow dr = dtDetailLineData.NewRow();

                                        dr["Page No"] = DetailPageNo;
                                        dr["Line No"] = DetailLineNo;
                                        dr["Word No"] = word.WordNumber;
                                        dr["Word"] = word.strWord;
                                        dr["Left"] = word.Left;
                                        dr["Top"] = word.Top;
                                        dr["Right"] = word.Right;
                                        dr["Bottom"] = word.Bottom;
                                        dr["ObjWord"] = word;
                                        dtDetailLineData.Rows.Add(dr);

                                    }
                                }
                            }
                        }
                    }
                }
            }
            catch (Exception)
            {

                throw;
            }
            finally
            {
                GC.Collect();
            }
            return dtDetailLineData;
        }



        private string F3_GetValue(clsCncMetaData ObjMetaData, int intCurrPageNumber, string stringArgs)
        {
            ClsCNC oClsCnc = new ClsCNC();

            int x1, y1, x2, y2 = 0;

            string[] args = stringArgs.Split(new char[] { '~' });

            x1 = Convert.ToInt32(args[0]);
            y1 = Convert.ToInt32(args[1]);
            x2 = Convert.ToInt32(args[2]);
            y2 = Convert.ToInt32(args[3]);

            List<clsCnCLine> Lines = oClsCnc.GetLinesFromROI(ObjMetaData.Page, intCurrPageNumber, x1, y1, x2, y2).ToList();


            //List<clsCnCLine> Lines = ObjMetaData.Page[intCurrPageNumber].Line.ToList();

            //string Synonyms = "Quantity~Weight~WElGHT~Cl~Item Description~Commodity Description~Class~NMFC~Qty~Pieces~Units~Qty~CTNS~Pieces/Quantity~Package~WT~Pkgs~Net Weight~Gross Weight";
            //string Synonyms = "DESC~Description~DESCRIPT]ON~Pallet~Pallets~* WElGHT (SUBJECT~WElGHT~DEDCRIPTION~LBS~FRT~Freight~Packaging~Pieces~CODE CLASS~NMFC CODE~WEIGHT - LB~WGHT~Description~PC~PKG~NMFC Item~Gross Weight~PACKAGES~Type~NMFC No~NMFC number~WGT~Freight ID Description~PACKAGE TYPE~Quantity~Weight~WElGHT~Cl~Item Description~Commodity Description~Class~NMFC~Qty~Pieces~Units~Qty~CTNS~Pieces/Quantity~Package~WT~Pkgs";
            //PossibleWords1 = GetColumnHeaderLineStatus(Lines, Synonyms, intCurrPageNumber);
            string LineNo = GetColumnHeaderLineStatus(Lines, Module1.Synonyms, intCurrPageNumber);
            //returnZones.Words = PossibleWords1
            //return PossibleWords1;
            return LineNo;

        }

        private DataTable F3_GetDetailValue(clsCncMetaData ObjMetaData, int intCurrPageNumber, string Synonyms, int IntLineNo)
        {

            //string LineNo = F3_GetValue(ObjMetaData, intCurrPageNumber, Synonyms, 0, 0, Wid, h);
            //int IntLineNo = Convert.ToInt32(LineNo);
            DataRow dr;
            MakeDtDeatilLine();

            char[] seperator = { '\\' };

            string[] file = PublicComponents.PROPath.Split(seperator);


            //string connStr = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";

            string[] arrFileName = file[file.Count() - 1].Replace(".PRO", "").Split('_');

            string PROFilename;

            if (arrFileName.Count() == 1)
            {
                PROFilename = arrFileName[0];
            }
            else
                PROFilename = arrFileName[arrFileName.Count() - 1];

            string code = string.Empty;


            if (PublicComponents.htMyVariable.ContainsKey("BOLSET4_ShpCode"))
            {
                code = Convert.ToString(PublicComponents.htMyVariable["BOLSET4_ShpCode"]);

            }
            else
            {
                //MyConnection = new System.Data.OleDb.OleDbConnection(connStr);
                //MyCommand = new System.Data.OleDb.OleDbDataAdapter("select * from [Sheet1$] where [FHPRO]=" + PROFilename + "", MyConnection);


                //DtSet = new System.Data.DataSet();
                //MyCommand.Fill(DtSet);
                //MyConnection.Close();

                //MyCommand.Dispose();
                //MyConnection.Dispose();


                //if (DtSet != null)
                //{
                //    if (DtSet.Tables.Count > 0)
                //    {
                //        if (DtSet.Tables[0].Rows.Count > 0)
                //            code = Convert.ToString(DtSet.Tables[0].Rows[0][1]);
                //        //else
                //        //    return null;
                //    }
                //    //if (code.StartsWith("99"))
                //    //    return null;
                //}
            }

            //if (code.StartsWith("99") || string.IsNullOrEmpty(code))
            //    return null;

            string shipcode = string.Empty;
            //code = "0739394";

            shipcode = code.PadLeft(7, '0');
            //if (code.Trim().Length < 7)
            //{
            //    for (int k = 1; k <= 7 - code.Trim().Length; k++)
            //        shipcode = "0" + shipcode;
            //}


            //shipcode = "0107778";
            //string Query_String = "SELECT F30.SASBLC, F31.SBDES,F31.SBNMFC,F31.SBSUB,F31.SBCMCL FROM FRP030 F30 Inner join FRP031 F31 on F30.SASBLC = F31.SBSBLC where F30.SASCD ='" + batchname[batchname.Count() - 1] + "'";
            string queryString = "SELECT F30.SASBLC, F31.SBDES,F31.SBNMFC,F31.SBSUB,F31.SBCMCL FROM FRP030 F30 Inner join FRP031 F31 on F30.SASBLC = F31.SBSBLC where F30.SASCD ='" + shipcode + "'";



            DataTable dtGetData = GetDataTableByText("", queryString);

            for (int i = IntLineNo + 1; i < ObjMetaData.Page[intCurrPageNumber].LineCount; i++)
            {

                if (ObjMetaData.Page[intCurrPageNumber].Line[i].strLine.ToUpper().Contains("TOTAL") || ObjMetaData.Page[intCurrPageNumber].Line[i].strLine.ToUpper().Contains("GRAND TOTAL"))
                    break;

                dr = dtDetailLine.NewRow();
                //int NumeriCount = 0;
                //int result;
                //double dResult;
                foreach (clsCnCWord word in ObjMetaData.Page[intCurrPageNumber].Line[i].Word)
                {
                    if (word != null)
                    {
                        if (dtGetData != null)
                        {
                            foreach (DataRow item in dtGetData.Rows)
                            {
                                if ((word.strWord.Contains(Convert.ToString(item["SBNMFC"]).Trim()) || Convert.ToString(item["SBCMCL"]).Trim() == word.strWord.Trim()) && Convert.ToString(item["SBNMFC"]).Trim() != "0")
                                //if (Convert.ToString(item["SBNMFC"]).Trim() == word.strWord.Trim() || Convert.ToString(item["SBCMCL"]).Trim() == word.strWord.Trim())
                                {
                                    dr["Page No"] = intCurrPageNumber;
                                    dr["Line No"] = i;
                                    //dr["NumericWordCount"] = NumeriCount;
                                    DataRow[] drline = dtDetailLine.Select("[Line No]=" + i);
                                    if (drline != null)
                                    {
                                        if (drline.Count() == 0)
                                        {
                                            dtDetailLine.Rows.Add(dr);

                                        }

                                    }
                                    else
                                    {
                                        dtDetailLine.Rows.Add(dr);
                                    }
                                }
                            }
                        }
                        //if (int.TryParse(word.strWord.Replace(",", "").Replace("-", ""), out  result) || double.TryParse(word.strWord.Replace(",", "").Replace("-", ""), out  dResult))
                        //    NumeriCount = NumeriCount + 1;

                    }
                }
            }

            //if (dtDetailLine.Rows.Count == 0)
            //{
            int k = 0;
            for (int i = IntLineNo + 1; i < ObjMetaData.Page[intCurrPageNumber].LineCount; i++)
            {

                if (ObjMetaData.Page[intCurrPageNumber].Line[i].strLine.ToUpper().Contains("SUB WEIGHT") || ObjMetaData.Page[intCurrPageNumber].Line[i].strLine.ToUpper().Contains("TOTAL") || ObjMetaData.Page[intCurrPageNumber].Line[i].strLine.ToUpper().Contains("GRAND TOTAL"))
                    break;

                if (k >= 5 || (k >= 1 && dtDetailLine.Rows.Count > 0))
                    break;

                dr = dtDetailLine.NewRow();
                int NumeriCount = 0;
                int result;
                double dResult;

                foreach (clsCnCWord word in ObjMetaData.Page[intCurrPageNumber].Line[i].Word)
                {
                    if (word != null)
                    {
                        if (int.TryParse(word.strWord.Replace(",", "").Replace("-", ""), out  result) || double.TryParse(word.strWord.Replace(",", "").Replace("-", ""), out  dResult))
                            NumeriCount = NumeriCount + 1;

                    }
                }
                if (NumeriCount >= 2)
                {
                    dr["Page No"] = intCurrPageNumber;
                    dr["Line No"] = i;
                    dr["NumericWordCount"] = NumeriCount;

                    if (!(dtDetailLine.AsEnumerable().Any(t => t.Field<string>("Line No") == i.ToString())))
                        dtDetailLine.Rows.Add(dr);
                    k = 0;
                }
                else
                {
                    k++;
                }
            }


            return dtDetailLine;
        }

        private string GetColumnHeaderLineStatus(List<clsCnCLine> Lines, string Synonyms, int intCurrPageNumber)
        {
            string[] ArgArr = null;
            List<clsCnCWord> Possiblewords = new List<clsCnCWord>();
            DataTable dt = new DataTable();
            dt.Columns.Add("Keyword");
            dt.Columns.Add("Page No");
            dt.Columns.Add("Line No");
            dt.Columns.Add("Word No");
            dt.Columns.Add("WordText");

            try
            {
                if (Synonyms.Contains("~"))
                {
                    ArgArr = Synonyms.Split('~');
                    //Lines.Reverse();

                    for (int i = 0; i <= Lines.Count - 1; i++)
                    {
                        if (Lines[i] != null)
                        {
                            Possiblewords = new List<clsCnCWord>();

                            for (int j = 0; j <= ArgArr.Length - 1; j++)
                            {
                                for (int k = 0; k <= Lines[i].WordCount; k++)
                                {
                                    //    if (Lines[i].WordCount <=3)
                                    //        break;

                                    if (Lines[i].Word[k] != null)
                                    {
                                        if (CheckValidWord(Lines[i].Word[k].strWord.ToUpper().Trim().Replace("'", "`"), new String[] { ArgArr[j] }) || Lines[i].Word[k].strWord.ToUpper().Trim().Contains(ArgArr[j].ToUpper()))
                                        {
                                            if (!Possiblewords.Contains(Lines[i].Word[k]))
                                            {
                                                DataRow dr = dt.NewRow();

                                                Possiblewords.Add(Lines[i].Word[k]);
                                                dr["Page No"] = Lines[i].Word[k].PageNo;
                                                dr["Keyword"] = ArgArr[j];
                                                dr["Line No"] = Lines[i].Word[k].LineNo;
                                                dr["Word No"] = Lines[i].Word[k].WordNumber;
                                                dr["WordText"] = Lines[i].Word[k].strWord;
                                                dt.Rows.Add(dr);
                                            }

                                        }
                                        else
                                        {
                                            string[] items = Lines[i].Word[k].strWord.Split(' ');

                                            foreach (var str in items)
                                            {
                                                if (CheckValidWord(str.ToUpper().Trim().Replace("'", "`"), new String[] { ArgArr[j] }) || str.ToUpper().Trim().Contains(ArgArr[j].ToUpper()))
                                                {
                                                    if (!Possiblewords.Contains(Lines[i].Word[k]))
                                                    {
                                                        DataRow dr = dt.NewRow();

                                                        Possiblewords.Add(Lines[i].Word[k]);
                                                        dr["Page No"] = Lines[i].Word[k].PageNo;
                                                        dr["Keyword"] = ArgArr[j];
                                                        dr["Line No"] = Lines[i].Word[k].LineNo;
                                                        dr["Word No"] = Lines[i].Word[k].WordNumber;
                                                        dr["WordText"] = Lines[i].Word[k].strWord;
                                                        dt.Rows.Add(dr);
                                                    }

                                                }
                                            }
                                        }
                                    }

                                }
                            }

                            //if (Possiblewords.Count >= 2)t=>t.
                            //break;
                        }
                    }
                }
                var group = dt.AsEnumerable().GroupBy(t => t.Field<string>("Line No")).Select(t => new { Lineno = t.Key, WordCount = t.Count() });
                DataTable dtGroup = new DataTable();

                dtGroup.Columns.Add("LineNo");
                dtGroup.Columns.Add("Count");

                if (group.Count() > 0)
                {
                    foreach (var item in group)
                    {

                        DataRow dr = dtGroup.NewRow();
                        dr["LineNo"] = item.Lineno;
                        dr["Count"] = item.WordCount;
                        dtGroup.Rows.Add(dr);
                    }
                }

                DataRow[] drmaxCount = dtGroup.Select("Count = MAX(Count) and MAX(Count)>=3");
                DataRow[] drminCount = dtGroup.Select("Count = MAX(Count) and MAX(Count)>=2");

                if (drmaxCount.Count() > 0)
                    return Convert.ToString(drmaxCount[0]["LineNo"]);
                else if (drminCount.Count() > 0)
                {
                    int lineno = Convert.ToInt32(drminCount[0]["LineNo"]);
                    int PreviousLine = lineno + 1;
                    int nextLine = lineno - 1;
                    DataRow[] drlines = dtGroup.Select("LineNo=" + PreviousLine + " OR LineNo=" + nextLine + "");

                    if (drlines.Count() > 0)
                    {
                        return Convert.ToString(lineno);
                    }
                    else
                        return "0";
                }
                else
                    return "0";

                //if (Possiblewords != null)
                //{
                //    if (Possiblewords.Count >= 2)
                //    {
                //        return Possiblewords;
                //    }
                //    else
                //        return null;

                //}
                //else
                //    return null;



            }
            catch (Exception)
            {

                throw;
            }
            finally
            {

            }

        }

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
        //    iQPUBLIC.clsCNCSBR objRestruct056 = new iQPUBLIC.clsCNCSBR();
        //    iQPUBLIC.RetStructIQSBR058 objRestruct058;

        //    if (!string.IsNullOrEmpty(sWord))
        //    {

        //        // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
        //        List<string> arrFinalkeyword = new List<string>();
        //        lengthofword = sWord.Length;


        //        if (sWord.ToUpper() == arrcheckkeywords[0].ToUpper())
        //        {
        //            return true;
        //        }

        //        if ((lengthofword <= 5))
        //        {
        //            //for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            //{
        //            //if ((arrcheckkeywords[iLoop].Length <= 5))
        //            if ((arrcheckkeywords[0].Length <= 5))
        //                arrFinalkeyword.Add(arrcheckkeywords[0]);
        //            //}
        //        }
        //        else if ((lengthofword >= 6 & lengthofword <= 8))
        //        {
        //            //for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            //{
        //            if ((arrcheckkeywords[0].Length >= 4 & arrcheckkeywords[0].Length <= 8))
        //                arrFinalkeyword.Add(arrcheckkeywords[0]);
        //            //}
        //        }
        //        else if ((lengthofword >= 9 & lengthofword <= 11))
        //        {
        //            //for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            //{
        //            if ((arrcheckkeywords[0].Length >= 9 & arrcheckkeywords[0].Length <= 11))
        //                arrFinalkeyword.Add(arrcheckkeywords[0]);
        //            //}
        //        }
        //        else if ((lengthofword >= 12))
        //        {
        //            //for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
        //            //{
        //            if ((arrcheckkeywords[0].Length >= 8))
        //                arrFinalkeyword.Add(arrcheckkeywords[0]);
        //            //}
        //        }

        //        for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)
        //        {
        //            objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, CompareOptions.EliminateBlankNPunctuation);

        //            if (lengthofword <= 3)
        //            {
        //                if (objRestruct058.CharChanged < 1)
        //                    bFlag = true;
        //            }

        //            if ((lengthofword <= 5))
        //            {
        //                if ((objRestruct058.CharChanged < 1))
        //                    bFlag = true;
        //            }
        //            else if ((lengthofword >= 6 & lengthofword <= 8))
        //            {
        //                if ((objRestruct058.CharChanged < 2))
        //                    bFlag = true;
        //            }
        //            else if ((lengthofword >= 9 & lengthofword <= 11))
        //            {
        //                if ((objRestruct058.CharChanged < 3))
        //                    bFlag = true;
        //            }
        //            else if ((lengthofword >= 12))
        //            {
        //                if ((objRestruct058.CharChanged < 4))
        //                    bFlag = true;
        //            }
        //        }
        //    }

        //    return bFlag;
        //}

        private bool CheckValidWord(string sWord, string[] arrcheckkeywords)
        {
            bool bFlag = false;
            int lengthofword;
            // Dim AllMatches As MatchCollection = Regex.Matches(sWord, "[a-zA-Z0-9]+", RegexOptions.IgnoreCase)
            // For Each SingleMatch As Match In AllMatches
            // If Array.IndexOf(arrAllWords, SingleMatch.Value.ToUpper()) >= 0 Then
            // bFlag = True
            // End If
            // Next
            iQPUBLIC.clsCNCSBR objRestruct056 = new iQPUBLIC.clsCNCSBR();
            iQPUBLIC.RetStructIQSBR058 objRestruct058;

            if (!string.IsNullOrEmpty(sWord))
            {

                // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
                List<string> arrFinalkeyword = new List<string>();
                lengthofword = sWord.Length;


                //if (sWord.ToUpper() == arrcheckkeywords[0].ToUpper())
                //{
                //    return true;
                //}

                if ((lengthofword <= 5))
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if ((arrcheckkeywords[iLoop].Length <= 5))
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
                    }
                }
                else if ((lengthofword >= 6 & lengthofword <= 8))
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if ((arrcheckkeywords[iLoop].Length >= 4 & arrcheckkeywords[iLoop].Length <= 8))
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop].ToString());
                    }
                }
                else if ((lengthofword >= 9 & lengthofword <= 11))
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if ((arrcheckkeywords[iLoop].Length >= 9 & arrcheckkeywords[iLoop].Length <= 11))
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop].ToString());
                    }
                }
                else if ((lengthofword >= 12))
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if ((arrcheckkeywords[iLoop].Length >= 8))
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop].ToString());
                    }
                }

                for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)
                {
                    objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, CompareOptions.EliminateBlankNPunctuation);

                    if ((lengthofword <= 5))
                    {
                        if ((objRestruct058.CharChanged < 1))
                            bFlag = true;
                    }
                    else if ((lengthofword >= 6 & lengthofword <= 8))
                    {
                        if ((objRestruct058.CharChanged < 2))
                            bFlag = true;
                    }
                    else if ((lengthofword >= 9 & lengthofword <= 11))
                    {
                        if ((objRestruct058.CharChanged < 3))
                            bFlag = true;
                    }
                    else if ((lengthofword >= 12))
                    {
                        if ((objRestruct058.CharChanged < 4))
                            bFlag = true;
                    }
                }
            }

            return bFlag;
        }

        private clsCnCWord GetDefaultWord(int iCurpgno, string strval)
        {
            clsCnCWord oRetword = new clsCnCWord();
            oRetword.X1Char = "5";
            oRetword.Y1Char = "5";
            oRetword.X2Char = "10";
            oRetword.X1Char = "10";
            oRetword.Confidence = 90;
            oRetword.LineNo = 1;
            oRetword.PageNo = iCurpgno;
            oRetword.Left = 10;
            oRetword.Right = 20;
            oRetword.Top = 10;
            oRetword.Bottom = 20;
            oRetword.strWord = strval.ToString();
            oRetword.ConfString = "9".PadLeft(strval.Length);

            return oRetword;
        }

        //private DataTable MakeDtForCustomF3()
        //{
        //    dtKeyword = new DataTable();
        //    dtKeyword.Columns.Add("Keyword");
        //    dtKeyword.Columns.Add("Page No");
        //    dtKeyword.Columns.Add("Line No");
        //    dtKeyword.Columns.Add("Word No");
        //    dtKeyword.Columns.Add("WordText");

        //    return dtKeyword;
        //}

        private DataTable MakeDtDeatilLine()
        {
            dtDetailLine = new DataTable();
            dtDetailLine.Columns.Add("Page No");
            dtDetailLine.Columns.Add("Line No");
            dtDetailLine.Columns.Add("NumericWordCount");
            return dtDetailLine;
        }

        private DataTable MakeDtDeatilLineData()
        {
            dtDetailLineData = new DataTable();
            dtDetailLineData.Columns.Add("Page No");
            dtDetailLineData.Columns.Add("Line No");
            dtDetailLineData.Columns.Add("Word No");
            dtDetailLineData.Columns.Add("Word");
            dtDetailLineData.Columns.Add("Left");
            dtDetailLineData.Columns.Add("Top");
            dtDetailLineData.Columns.Add("Right");
            dtDetailLineData.Columns.Add("Bottom");
            dtDetailLineData.Columns.Add("ObjWord", typeof(clsCnCWord));
            return dtDetailLineData;
        }

        private DataTable MakedtDetailFinal()
        {
            dtDetailLineData = new DataTable();
            dtDetailLineData.Columns.Add("Line Number");
            dtDetailLineData.Columns.Add("Quantity");
            dtDetailLineData.Columns.Add("Class Code");
            dtDetailLineData.Columns.Add("HazMat code");
            dtDetailLineData.Columns.Add("Pallate Code");
            dtDetailLineData.Columns.Add("Description");
            dtDetailLineData.Columns.Add("NMFC");
            dtDetailLineData.Columns.Add("Weight");
            dtDetailLineData.Columns.Add("IsValid");
            dtDetailLineData.Columns.Add("WQuantity", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("WClass Code", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("WHazMat code", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("WPallate Code", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("WDescription", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("WNMFC", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("WWeight", typeof(clsCnCWord));
            dtDetailLineData.Columns.Add("PROLineNo");

            return dtDetailLineData;
        }

        #endregion


    }
                #endregion
}
