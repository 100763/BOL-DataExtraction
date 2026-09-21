using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.Text.RegularExpressions;
using System.Windows.Forms;
using iQDataProvider;
using iQPUBLIC;

namespace BOLSET2IQ
{
    public static class Module1
    {

        public static RetStructF3 objRetStructF3 = new RetStructF3();
        public static RetStructF5 objRetStructF5 = new RetStructF5();
        public static ClsCNC objClsCNCForROI = new ClsCNC();
        public static DataTable TableRecordss = new DataTable();
        public static DataTable TableRecordssFinal = new DataTable();
        public static string imageName = string.Empty;
        public static string ActualimageName = string.Empty;
        public static DataTable dtUSCityStateZip = new DataTable();
        public static DataTable dtUSAStateCodeName = new DataTable();
        public static DataTable dtAbbreviationTable = new DataTable();

        // For Micro ROI
        public static int MX1, MY1, MX2, MY2;
        public static bool CheckValidWord(string sWord, string[] arrcheckkeywords)
        {
            bool bFlag = false;
            int lengthofword;
            // Dim AllMatches As MatchCollection = Regex.Matches(sWord, "[a-zA-Z0-9]+", RegexOptions.IgnoreCase)
            // For Each SingleMatch As Match In AllMatches
            // If Array.IndexOf(arrAllWords, SingleMatch.Value.ToUpper()) >= 0 Then
            // bFlag = True
            // End If
            // Next
            var objRestruct056 = new iQPUBLIC.clsCNCSBR();
            iQPUBLIC.RetStructIQSBR058 objRestruct058;
            if (!string.IsNullOrEmpty(sWord))
            {

                // Dim arrcheckkeywords() As String = sAllValidWords.Split("#")
                var arrFinalkeyword = new List<string>();
                lengthofword = sWord.Length;
                if (lengthofword <= 5)
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if (arrcheckkeywords[iLoop].Length <= 5)
                        {
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
                        }
                    }
                }
                else if (lengthofword >= 6 & lengthofword <= 8)
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if (arrcheckkeywords[iLoop].Length >= 4 && arrcheckkeywords[iLoop].Length <= 8)
                        {
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
                        }
                    }
                }
                else if (lengthofword >= 9 & lengthofword <= 11)
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if (arrcheckkeywords[iLoop].Length >= 9 && arrcheckkeywords[iLoop].Length <= 11)
                        {
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
                        }
                    }
                }
                else if (lengthofword >= 12)
                {
                    for (int iLoop = 0; iLoop <= arrcheckkeywords.Length - 1; iLoop++)
                    {
                        if (arrcheckkeywords[iLoop].Length >= 8)
                        {
                            arrFinalkeyword.Add(arrcheckkeywords[iLoop]);
                        }
                    }
                }

                for (int iLoop = 0; iLoop <= arrFinalkeyword.Count - 1; iLoop++)

                {
                    objRestruct058 = objRestruct056.IQSBR058(sWord, arrFinalkeyword[iLoop], true, iQPUBLIC.CompareOptions.EliminateBlankNPunctuation);
                    if (lengthofword <= 5)
                    {
                        if (objRestruct058.CharChanged < 1)
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofword >= 6 & lengthofword <= 8)
                    {
                        if (objRestruct058.CharChanged < 2)
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofword >= 9 & lengthofword <= 11)
                    {
                        if (objRestruct058.CharChanged < 3)
                        {
                            bFlag = true;
                        }
                    }
                    else if (lengthofword >= 12)
                    {
                        if (objRestruct058.CharChanged < 4)
                        {
                            bFlag = true;
                        }
                    }
                }
            }

            return bFlag;
        }
        public static string func_RemoveSpecialCharacter(string str_Data)
        {
            //string[] array_RemoveSpecialCharacter = new[] { " ", "  ", ":", ".", "/", "!", "@", "#", "$", "%", "^", "*", "'", @"\", ";", "_", "(", ")", "|", "[", "]", ",", "-" };

            //foreach (string str_RemoveSpecialCharacter in array_RemoveSpecialCharacter)
            //    str_Data = str_Data.Replace(str_RemoveSpecialCharacter, "");

            str_Data = Regex.Replace(str_Data, "[^a-zA-Z0-9]", ""); // Above code commented and added on 18-Feb-2021
            return str_Data.ToUpper().Trim();
        }
        public static string func_RemoveSpecialCharacterwithSpace(string str_Data)
        {
            //string[] array_RemoveSpecialCharacter = new[] { " ", "  ", ":", ".", "/", "!", "@", "#", "$", "%", "^", "*", "'", @"\", ";", "_", "(", ")", "|", "[", "]", ",", "-" };

            //foreach (string str_RemoveSpecialCharacter in array_RemoveSpecialCharacter)
            //    str_Data = str_Data.Replace(str_RemoveSpecialCharacter, "");

            str_Data = Regex.Replace(str_Data, "[^a-zA-Z0-9\\s%/]", ""); // Above code commented and added on 18-Feb-2021

            return str_Data.ToUpper().Trim();
        }
        public static string func_RemoveKeywordsNameAddress_Bck(string str_Data)
        {
            //string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address:", "Address", "Address2:", "Addrest", "Ship From:", "Ship To:", "SHIP TO", ":", "Néme", "Néme:","NEME","FROM","from","From","TO:","SHIPPER","shipper","Shipper","consignee","CONSIGNEE", "CITY","STATE","ZIP","CODE" };
            // string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address:", "Address", "Address2:", "Addrest", "Ship From:", "Ship To:", "SHIP TO", ":", "Néme", "Néme:", "NEME", "FROM", "from", "From", "TO:", "SHIPPER", "shipper", "Shipper", "consignee", "CONSIGNEE", "CITY", "STATE", "ZIP", "BILLING", "ZLP" };
            //string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address:", "Address", "Address2:", "Addrest", "Ship From:", "Ship To:", "SHIP TO", ":", "Néme", "Néme:", "NEME", "FROM", "from", "From", "TO:", "SHIPPER", "consignee", "CONSIGNEE", "CITY", "STATE", "ZIP", "BILLING", "ZLP", "Prepaid" };

            string[] array_RemoveKeywords = new[] { "Address2:", "Address3:", "Address1:", "Address:", "Address", "Addrest", "Ship From:", "Ship To:", "SHIP TO", "TO:", "Néme", "Néme:", "SHIPPER", "CONSIGNEE", "CITY", "STATE", "Zip Code", "ZIP", "ZLP", "DELIVERY", ":" };
            //string[] array_RemoveKeywords = new[] { "Address2:", "Address3:", "Address1:", "Address:", "Address", "Addrest", "Ship From:", "Ship To:", "SHIP TO", "TO:", "Néme", "Néme:", "SHIPPER", "CONSIGNEE", "CITY",  "Zip Code", "ZIP", "ZLP", "DELIVERY", ":" };  // Added on 24-Nov-2021 => Removed STATE keyword
            //string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address2:", "Address3:", "Address1:", "Address:", "Address", "Addrest", "Ship From:", "Ship To:", "SHIP TO", "TO:", "Néme", "Néme:", "NEME", "FROM", "SHIPPER", "CONSIGNEE", "CITY", "STATE", "Zip Code", "ZIP", "ZLP", "DELIVERY", "DELIVER", ":" };

            foreach (string str_RemoveKeywords in array_RemoveKeywords)
                str_Data = str_Data.ToUpper().Replace(str_RemoveKeywords.ToUpper(), "");

            return str_Data.ToUpper().Trim();
        }
        // Added on 25-Nov-2021 -- Start
        public static string func_RemoveKeywordsNameAddress(string str_Data)
        {
            //string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address:", "Address", "Address2:", "Addrest", "Ship From:", "Ship To:", "SHIP TO", ":", "Néme", "Néme:","NEME","FROM","from","From","TO:","SHIPPER","shipper","Shipper","consignee","CONSIGNEE", "CITY","STATE","ZIP","CODE" };
            // string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address:", "Address", "Address2:", "Addrest", "Ship From:", "Ship To:", "SHIP TO", ":", "Néme", "Néme:", "NEME", "FROM", "from", "From", "TO:", "SHIPPER", "shipper", "Shipper", "consignee", "CONSIGNEE", "CITY", "STATE", "ZIP", "BILLING", "ZLP" };
            //string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address:", "Address", "Address2:", "Addrest", "Ship From:", "Ship To:", "SHIP TO", ":", "Néme", "Néme:", "NEME", "FROM", "from", "From", "TO:", "SHIPPER", "consignee", "CONSIGNEE", "CITY", "STATE", "ZIP", "BILLING", "ZLP", "Prepaid" };
            //string[] array_RemoveKeywords = new[] { "Name:", "Name", "Address2:", "Address3:", "Address1:", "Address:", "Address", "Addrest", "Ship From:", "Ship To:", "SHIP TO", "TO:", "Néme", "Néme:", "NEME", "FROM", "SHIPPER", "CONSIGNEE", "CITY", "STATE", "Zip Code", "ZIP", "ZLP", "DELIVERY", "DELIVER", ":" };
                       
            // string[] array_RemoveKeywords = new[] { "Address2:", "Address3:", "Address1:", "Address:", "Address", "Addrest", "Ship From:", "Ship To:", "SHIP TO", "TO:", "Néme", "Néme:", "SHIPPER", "CONSIGNEE", "CITY", "STATE", "Zip Code", "ZIP", "ZLP", "DELIVERY", ":" }; //commented on 25Nov21
            string[] array_RemoveKeywords = new[] { "Address2:", "Address3:", "Address1:", "Address:", "Address", "Addrest", "Ship From:", "Ship To:", "SHIP TO", "TO:", "Néme", "Néme:", "SHIPPER", "CONSIGNEE", "Zip Code", "ZIP", "ZLP", "DELIVERY", ":" };

            //*****************Added on 25Nov21********
            string[] arValue = str_Data.Trim().Split(' ');

            for (int i = 0; i < arValue.Length; i++)
            {
                if (Regex.Replace(arValue[i], "[^a-zA-Z0-9]", "").Trim().Length > 1)
                {
                    if (arValue[i].ToUpper().Trim() == "STATE" || arValue[i].ToUpper().Trim() == "CITY")
                    {
                        str_Data = "";
                        for (int j = i + 1; j < arValue.Length; j++)
                        {
                            str_Data = str_Data + " " + arValue[j];
                        }
                        str_Data = str_Data.Trim();
                    }
                    break;
                }
            }
            //*******************************************

            foreach (string str_RemoveKeywords in array_RemoveKeywords)
                str_Data = str_Data.ToUpper().Replace(str_RemoveKeywords.ToUpper(), "");

            return str_Data.ToUpper().Trim();
        }
        // Added on 25-Nov-2021 -- End
        public static string func_RemoveKeywordsFromName(string str_Data)
        {

            try
            {

                Match objMatch = Regex.Match(str_Data, " location#| location #| loc#| loc #", RegexOptions.IgnoreCase);
                if (objMatch.Success)
                {
                    str_Data = str_Data.Substring(0, objMatch.Index);
                }

                str_Data = str_Data.Trim().Trim(':');
                if (str_Data.Contains(":"))
                {
                    str_Data = str_Data.Substring(str_Data.IndexOf(":") + 1, str_Data.Length - str_Data.IndexOf(":") - 1).Trim();
                }
                else
                {
                    string[] array_RemoveKeywords = new[] { "NAME", "SHIPPER", "CONSIGNEE", "COSIGNEE", "PARTY", "DESTINATION", "FROM", "TO", "CHARGES",  "DELIVERY", "DELIVER", "ORIGIN" };//"FREIGHT",

                    foreach (string str_RemoveKeywords in array_RemoveKeywords)
                    {
                        int indexPos = str_Data.ToUpper().IndexOf(str_RemoveKeywords);

                        if (indexPos >= 0)
                        {
                            if (indexPos == 0 && Regex.IsMatch(str_Data.Substring(indexPos + str_RemoveKeywords.Length, 1), "[a-zA-Z0-9]") == false)
                            {
                                str_Data = str_Data.ToUpper().Substring(str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length, str_Data.ToUpper().Length - (str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length));
                            }
                            else if (Regex.IsMatch(str_Data.Substring(indexPos - 1, 1), "[a-zA-Z0-9]") == false && Regex.IsMatch(str_Data.Substring(indexPos + str_RemoveKeywords.Length, 1), "[a-zA-Z0-9]") == false)
                            {
                                str_Data = str_Data.ToUpper().Substring(str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length, str_Data.ToUpper().Length - (str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length));
                            }

                        }
                    }

                }
            }
            catch (Exception ex)
            {
            }

            return str_Data.ToUpper().Trim();
        }

        public static bool func_IgnoreKeyword(string str_Data)
        {
            bool bfound = false;

            try
            {

                if (str_Data.Trim().ToUpper().StartsWith("PO#"))
                {
                    bfound = true;
                }
                
            }
            catch (Exception ex)
            {
            }

            return bfound;
        }

        public static string func_RemoveKeywordsFromName_BackUp(string str_Data)
        {
            //string[] array_RemoveKeywords = new[] { "Name", "Ship To", "SHIPPER", "CONSIGNEE", "FROM", "BILL TO", "REMIT TO", "FREIGHT TO","THIRD PARTY", "3RD PARTY", "FREIGHT CHARGES", "DESTINATION", "DELIVERY", "DELIVER" };
            //string[] array_RemoveKeywords = new[] { "Name", "To", "SHIPPER", "CONSIGNEE", "FROM", "PARTY",  "CHARGES", "DESTINATION", "DELIVERY", "DELIVER" };
            //str_Data = Regex.Replace(str_Data, "[^a-zA-Z0-9\\s]", " ");
            //str_Data = str_Data.Trim();
            //str_Data = " " + str_Data + " ";

            try
            {
                if (str_Data.Contains(":"))
                {
                    str_Data = str_Data.Trim().Trim(':');
                    str_Data = str_Data.Substring(str_Data.IndexOf(":") + 1, str_Data.Length - str_Data.IndexOf(":") - 1).Trim();
                }
                else
                {
                    string[] array_RemoveKeywords = new[] { "NAME", "SHIPPER", "CONSIGNEE", "PARTY", "DESTINATION", "FROM", "TO", "CHARGES", "FREIGHT", "DELIVERY", "DELIVER" };

                    // string[] array_RemoveKeywords = new[] { " NAME ", " SHIPPER ", " CONSIGNEE ", " PARTY ", " DESTINATION ", " FROM ", " TO ", " CHARGES ", " DELIVERY ", " DELIVER " };

                    foreach (string str_RemoveKeywords in array_RemoveKeywords)
                    {
                        if (str_Data.ToUpper().IndexOf(str_RemoveKeywords) >= 0)
                        {
                            str_Data = str_Data.ToUpper().Substring(str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length, str_Data.ToUpper().Length - (str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length));
                        }
                    }

                    //foreach (string str_RemoveKeywords in array_RemoveKeywords)
                    //    str_Data = str_Data.ToUpper().Replace(str_RemoveKeywords.ToUpper(), "");
                }
            }
            catch (Exception ex)
            {
            }
            return str_Data.ToUpper().Trim();
        }

        public static string func_RemoveKeywordsFromAddress(string str_Data)
        {
            try
            {
                string[] array_RemoveKeywords = new[] { "NAME", "SHIPPER", "CONSIGNEE", "DESTINATION", "ADDRESS3", "ADDRESS2", "ADDRESS1", "ADDRESS", "TO", "ORIGIN" };

                foreach (string str_RemoveKeywords in array_RemoveKeywords)
                {
                    int indexPos = str_Data.ToUpper().IndexOf(str_RemoveKeywords);

                    if (indexPos >= 0)
                    {
                        if (indexPos == 0 && Regex.IsMatch(str_Data.Substring(indexPos + str_RemoveKeywords.Length, 1), "[a-zA-Z0-9]") == false)
                        {
                            str_Data = str_Data.ToUpper().Substring(str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length, str_Data.ToUpper().Length - (str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length));
                        }
                        else if (Regex.IsMatch(str_Data.Substring(indexPos - 1, 1), "[a-zA-Z0-9]") == false && Regex.IsMatch(str_Data.Substring(indexPos + str_RemoveKeywords.Length, 1), "[a-zA-Z0-9]") == false)
                        {
                            str_Data = str_Data.ToUpper().Substring(str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length, str_Data.ToUpper().Length - (str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length));
                        }

                    }
                }
                //foreach (string str_RemoveKeywords in array_RemoveKeywords)
                //{
                //    if (str_Data.ToUpper().IndexOf(str_RemoveKeywords) >= 0)
                //    {
                //        str_Data = str_Data.ToUpper().Substring(str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length, str_Data.ToUpper().Length - (str_Data.ToUpper().IndexOf(str_RemoveKeywords) + str_RemoveKeywords.Length));
                //    }
                //}
            }
            catch (Exception)
            {
            }
            return str_Data.ToUpper().Trim();
        }

        public static string func_RemoveKeywordsTelephone(string str_Data)
        {

            string Regxtel = "[^0-9]";
            string telTemp = Regex.Replace(str_Data, Regxtel, "", RegexOptions.IgnoreCase);
            return telTemp;

            //string[] array_RemoveKeywords = new[] { "Service,"," ", "  ", ":", ".", "/", "!", "@", "#", "$", "%", "^", "*", "'", @"\", ";", "_", "(", ")", "|", "[", "]", ",", "-","X", "x,","x" };

            //foreach (string str_RemoveKeywords in array_RemoveKeywords)
            //    str_Data = str_Data.Replace(str_RemoveKeywords, "");

            //return str_Data.ToUpper().Trim();

        }
        public static DataTable GetUSCityStateZipCode()
        {
            try
            {
                string mdbCityStateZipFilePath = Application.StartupPath + "\\Combine.mdb";
                clsOleDBDataAccess OdbcConnectionObject = new clsOleDBDataAccess();
                OdbcConnectionObject.propConnection = new System.Data.OleDb.OleDbConnection();
                OdbcConnectionObject.propConnection.ConnectionString = @"Provider=Microsoft.Jet.OLEDB.4.0;Data Source='" + mdbCityStateZipFilePath + "'";
                DataTable USCityStateZipTable = clsParseUSAddress.GetUSAZipCodeStateCityTable(OdbcConnectionObject, "USCityStateZip");
                DataTable USAStateCodeNameTable = clsParseUSAddress.GetUSAZipCodeStateCityTable(OdbcConnectionObject, "USAStateCodeName");
                dtUSCityStateZip = USCityStateZipTable;
                dtUSAStateCodeName = USAStateCodeNameTable;
                DataTable AbbreviationTable = clsSearchCustomer.GetAbbreviationsTable(OdbcConnectionObject);
                dtAbbreviationTable = AbbreviationTable;
                //AddressSection objDBAddress = clsParseUSAddress.ParseUSAddress(txtOCRAddress.Text.Trim(), USCityStateZipTable, USAStateCodeNameTable);

                //string cmdString = "select * from USCityStateZip where zipcode = '" + ZipCode + "' AND state = '" + StateCode + "'";
                //string cmdString = "select * from USCityStateZip";
                //OdbcConnectionObject.propConnection.Open();
                //DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USCityStateZip");
                //if (dataTable != null && dataTable.Rows != null && dataTable.Rows.Count > 0)
                //{
                //    dtUSCityStateZip = dataTable;
                //DataRow[] drRows = dataTable.Select("[zipcode] = '" + ZipCode + "' AND [state] = '" + StateCode + "'");
                //if (drRows != null && drRows.Length > 0)
                //{
                //    if (drRows.Length > 1)
                //    {
                //        foreach (DataRow dr in drRows)
                //        {
                //            if (Convert.ToString(dr["city"]).ToUpper().Contains(CityFromAddress.ToUpper()))
                //            {
                //                cityFoundValue = Convert.ToString(dr["city"]);
                //                break;
                //            }
                //        }
                //    }
                //    else
                //    {
                //        cityFoundValue = Convert.ToString(drRows[0]["city"]);
                //    }
                //}
                //}
                //OdbcConnectionObject.propConnection.Close();
                //OdbcConnectionObject = null;
            }
            catch (Exception ex)
            {

                //MessageBox.Show("GetUSCityStateZipCode:" + ex.Message);
                //   cityFoundValue = "";
            }
            return dtUSCityStateZip;
        }

        public static string func_RemoveKeywordsforAllName(string str_Data)
        {
           // string[] array_RemoveKeywords = new[] { "THIRD PARTY FREIGHT CHARGES BILL TO", "FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID", "BILL TO OR REMIT TO", "3RD PARTY BILL FREIGHT PREPAID", "BILL FREIGHT CHARGES", "BILL THIRD PARTY INFORMATION", "FREIGHT PAYMELT PLAN", "THIRD PARTY FREIGHT CHARGES BILL", "3RD PARTY BILL", "BILL OR REMIT", "FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT CHARGES", "THIRD PARTY FREIGHT", "FREIGHT PREPAID", "THIRD PARTY", "FREIGHT CHARGES", "BILLING ADDRESS", "REMIT", "BILL OR", "BILL TO", "BILL  TO", "STREET", "TO:", "SHIP TO", "SOLD TO", "NAME", "SHIPPING", "CONSIGNEE", "DESTINATION", "SHIP FROM", "SOLD FROM", "THIRD", "3RD", "PARTY", "FREIGHT", "CHARGES", "SHIPPED", "MANIFEST", "BILL FREIGHT TO", "SHIP", "TO:", "BILL:", "Prepaid", "Néme", "Néme:", "NEME", "C.O.D", "Address" };
            string[] array_RemoveKeywords = new[] { "THIRD PARTY FREIGHT CHARGES BILL TO", "FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID", "BILL TO OR REMIT TO", "3RD PARTY BILL FREIGHT PREPAID", "BILL FREIGHT CHARGES", "BILL THIRD PARTY INFORMATION", "FREIGHT PAYMELT PLAN", "THIRD PARTY FREIGHT CHARGES BILL", "3RD PARTY BILL", "BILL OR REMIT", "FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT CHARGES", "THIRD PARTY FREIGHT", "FREIGHT PREPAID", "THIRD PARTY", "FREIGHT CHARGES", "BILLING ADDRESS", "BILL OR", "BILL TO", "BILL  TO", "STREET", "TO:", "SHIP TO", "SOLD TO", "SHIPPER", "CONSIGNEE", "DESTINATION", "SHIP FROM", "SOLD FROM", "SHIPPED", "MANIFEST", "BILL FREIGHT TO",  "TO:", "BILL:",  "C.O.D", "Address" };

            //string[] array_RemoveKeywords = new[] { "THIRD PARTY FREIGHT CHARGES BILL TO", "BILL TO OR REMIT TO", "3RD PARTY BILL", "3RD PARTY BILL FREIGHT PREPAID", "BILL OR REMIT", "REMIT", "BILL OR", "FREIGHT PREPAID", "BILL FREIGHT CHARGES", "BILL THIRD PARTY INFORMATION", "FREIGHT PAYMELT PLAN", "FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID", "FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT", "THIRD PARTY FREIGHT CHARGES", "FREIGHT CHARGES", "THIRD PARTY", "FREIGHT CHARGES", "BILLING ADDRESS", "BILL TO", "BILL  TO", "STREET", "TO:", "SHIP TO", "SOLD TO", "NAME", "SHIPPING", "SHIPPER", "CONSIGNEE", "DELIVER", "DESTINATION", "SHIP FROM", "SOLD FROM", "THIRD", "3RD", "PARTY", "FREIGHT", "CHARGES", "SHIPPED", "MANIFEST", "BILL FREIGHT TO", "SHIP", "TO:", "BILL:", "BILLING" };
            //string[] array_RemoveKeywords = new[] { "THIRD PARTY FREIGHT CHARGES BILL TO", "BILL TO OR REMIT TO", "3RD PARTY BILL", "3RD PARTY BILL FREIGHT PREPAID", "BILL OR REMIT", "REMIT", "BILL OR", "FREIGHT PREPAID", "BILL FREIGHT CHARGES", "BILL THIRD PARTY INFORMATION", "FREIGHT PAYMELT PLAN", "FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID", "FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT", "THIRD PARTY FREIGHT CHARGES", "FREIGHT CHARGES", "THIRD PARTY", "FREIGHT CHARGES", "BILLING ADDRESS", "BILL TO", "BILL  TO", "STREET", "TO:", "SHIP TO", "SOLD TO", "NAME", "SHIPPING", "SHIPPER", "CONSIGNEE", "DELIVER", "DESTINATION", "SHIP FROM", "SOLD FROM", "THIRD", "3RD", "PARTY", "FREIGHT", "CHARGES", "SHIPPED", "MANIFEST", "BILL FREIGHT TO", "SHIP", "TO:", "BILL:", "BILLING" };
            //string[] array_RemoveKeywords = new[] { "THIRD PARTY FREIGHT CHARGES BILL TO", "BILL TO OR REMIT TO", "3RD PARTY BILL", "3RD PARTY BILL FREIGHT PREPAID", "BILL OR REMIT", "REMIT", "BILL OR", "FREIGHT PREPAID", "BILL FREIGHT CHARGES", "BILL THIRD PARTY INFORMATION", "FREIGHT PAYMELT PLAN", "FREIGHT CHARGE TERMS (FREIGHT CHARGES ARE PREPAID", "FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT CHARGES BILL", "THIRD PARTY FREIGHT", "THIRD PARTY FREIGHT CHARGES", "FREIGHT CHARGES", "THIRD PARTY", "FREIGHT CHARGES", "BILL TO", "BILL  TO", "STREET", "TO:", "SHIP TO", "SOLD TO", "NAME", "SHIPPING", "SHIPPER", "CONSIGNEE", "DELIVER", "DESTINATION", "SHIP FROM", "SOLD FROM", "THIRD", "3RD", "PARTY", "FREIGHT" };

            foreach (string str_RemoveKeywords in array_RemoveKeywords)
                str_Data = str_Data.ToUpper().Replace(str_RemoveKeywords.ToUpper(), "");

            return str_Data.ToUpper().Trim();
        }

        public static string func_RemoveSpecialWordsFromName(string str_Data)
        {
            //if (str_Data.EndsWith(" COMPANY") || str_Data.EndsWith(" CO") || str_Data.EndsWith(" INCORPORATED") || str_Data.EndsWith(" INC") || str_Data.EndsWith(" ENTERPRISES") || str_Data.EndsWith(" ENTERPRISE") || str_Data.EndsWith(" LLC"))
            //{ 
            //str_Data = str_Data.Replace(" COMPANY", "").Replace(" CO", "").Replace(" INCORPORATED", "").Replace(" INC", "").Replace(" ENTERPRISES", "").Replace(" ENTERPRISE", "").Replace(" LLC", "");
            //}
            // added on 31-March-2021 eg :- Shipper Name OBL = WYNDHAM COLLECTION, LLC => due to this it replaces CO and LLC and forms =>  Name = WYNDHAMLLECTION

            if (str_Data.EndsWith(" COMPANY"))
            {
                str_Data = str_Data.Replace(" COMPANY", "");
            }
            else if (str_Data.EndsWith(" CO"))
            {
                str_Data = str_Data.Replace(" CO", "");
            }
            else if (str_Data.EndsWith(" INCORPORATED"))
            {
                str_Data = str_Data.Replace(" INCORPORATED", "");
            }
            else if (str_Data.EndsWith(" INC"))
            {
                str_Data = str_Data.Replace(" INC", "");
            }
            else if (str_Data.EndsWith(" ENTERPRISES"))
            {
                str_Data = str_Data.Replace(" ENTERPRISES", "");
            }
            else if (str_Data.EndsWith(" ENTERPRISE"))
            {
                str_Data = str_Data.Replace(" ENTERPRISE", "");
            }
            else if (str_Data.EndsWith(" LLC"))
            {
                str_Data = str_Data.Replace(" LLC", "");
            }
            else if (str_Data.EndsWith(" SUPPLY"))
            {
                str_Data = str_Data.Replace(" SUPPLY", "");
            }
            else if (str_Data.EndsWith(" CORP"))
            {
                str_Data = str_Data.Replace(" CORP", "");
            }
            else if (str_Data.EndsWith(" CORPORATION"))
            {
                str_Data = str_Data.Replace(" CORPORATION", "");
            }
            return str_Data;
        }
    }
}