using iQPUBLIC;
using System;
using System.Data;
using System.Data.Odbc;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using iQPUBLIC;
using System.Collections.Generic;
using System.Configuration;

namespace BOLSET2IQ
{
    public static class clsSearchCustomer
    {
        //private static DataTable ATLTerminalData { get; set; }
        private static string GenerateSearchQueryForARP001(SearchParameters CustomerSearchParamaters)
        {
            string finalQuery = "";
            try
            {
                StringBuilder tempQeury = new StringBuilder();
                tempQeury.Append("SELECT * FROM ARP001 WHERE 1 = 1 AND CMSTAT = 'A' ");

                if (!string.IsNullOrEmpty(CustomerSearchParamaters.CustomerName))
                {
                    tempQeury.Append(" AND CMNAME LIKE '" + CustomerSearchParamaters.CustomerName.ToUpper() + "%'");
                }
                if (!string.IsNullOrEmpty(CustomerSearchParamaters.TerminalID))
                {
                    tempQeury.Append(" AND CMTID LIKE '" + CustomerSearchParamaters.TerminalID.ToUpper() + "%'");
                }
                if (!string.IsNullOrEmpty(CustomerSearchParamaters.StreetNumber))
                {
                    tempQeury.Append(" AND CMADRN = " + CustomerSearchParamaters.StreetNumber);
                }
                if (!string.IsNullOrEmpty(CustomerSearchParamaters.DirectionFirstLetter))
                {
                    tempQeury.Append(" AND CMADRU LIKE '" + CustomerSearchParamaters.DirectionFirstLetter.ToUpper() + "%'");
                }
                if (!string.IsNullOrEmpty(CustomerSearchParamaters.ZipCode))
                {
                    tempQeury.Append(" AND CMZIP LIKE '" + CustomerSearchParamaters.ZipCode + "%'");
                }
                if (!string.IsNullOrEmpty(CustomerSearchParamaters.CustomerFirstLetter))
                {
                    tempQeury.Append(" AND CMNAME LIKE '" + CustomerSearchParamaters.CustomerFirstLetter.ToUpper() + "%'");
                }
                if (!string.IsNullOrEmpty(CustomerSearchParamaters.City))
                {
                    tempQeury.Append(" AND CMCITY LIKE '" + CustomerSearchParamaters.City.ToUpper() + "%'");
                }

                if (!string.IsNullOrEmpty(Convert.ToString(tempQeury)))
                {
                    finalQuery = Convert.ToString(tempQeury);
                }
            }
            catch
            {
                finalQuery = "";
            }
            return finalQuery;
        }

        public static DataTable SearchCustomerInARP001(SearchParameters CustomerSearchParmaters, OdbcConnection OdbcConnectionObject)
        {
            DataTable dtCustomer = new DataTable();

            try
            {
                string searchQuery = GenerateSearchQueryForARP001(CustomerSearchParmaters);

                if (!string.IsNullOrEmpty(searchQuery))
                {
                    dtCustomer = ExecuteSearchQueryOnARP001(searchQuery, OdbcConnectionObject);
                }
            }
            catch
            {
                dtCustomer = null;
            }

            return dtCustomer;
        }

        public static DataTable MatchCustomerData(DataTable CustomerRecordsFromDB, AddressSection AddressSectionsFromOCR, clsOleDBDataAccess MDBOdbcConnectionObject)
        {
            DataTable dtResults = new DataTable();

            try
            {
                if (CustomerRecordsFromDB != null && CustomerRecordsFromDB.Rows.Count > 0)
                {
                    DataColumn recordNoColumn = new DataColumn("Sr.No.");
                    dtResults.Columns.Add(recordNoColumn);
                    DataColumn zipCodeDBColumn = new DataColumn("Zip DB");
                    dtResults.Columns.Add(zipCodeDBColumn);
                    DataColumn zipCodeOCRColumn = new DataColumn("Zip OCR");
                    dtResults.Columns.Add(zipCodeOCRColumn);
                    DataColumn zipCodeColumn = new DataColumn("ZipPercent", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(zipCodeColumn);
                    DataColumn stateCodeDBColumn = new DataColumn("State DB");
                    dtResults.Columns.Add(stateCodeDBColumn);
                    DataColumn stateCodeOCRColumn = new DataColumn("State OCR");
                    dtResults.Columns.Add(stateCodeOCRColumn);
                    DataColumn stateCodeColumn = new DataColumn("StatePercent", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(stateCodeColumn);
                    DataColumn cityDBColumn = new DataColumn("City DB");
                    dtResults.Columns.Add(cityDBColumn);
                    DataColumn cityOCRColumn = new DataColumn("City OCR");
                    dtResults.Columns.Add(cityOCRColumn);
                    DataColumn cityColumn = new DataColumn("CityPercent", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(cityColumn);
                    DataColumn addrLine1DBColumn = new DataColumn("AddrLine1 DB");
                    dtResults.Columns.Add(addrLine1DBColumn);
                    DataColumn addrLine1OCRColumn = new DataColumn("AddrLine1 OCR");
                    dtResults.Columns.Add(addrLine1OCRColumn);
                    DataColumn addrLine1Column = new DataColumn("AddrLine1Percent", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(addrLine1Column);
                    DataColumn directionDBColumn = new DataColumn("Direction DB");
                    dtResults.Columns.Add(directionDBColumn);
                    DataColumn directionOCRColumn = new DataColumn("Direction OCR");
                    dtResults.Columns.Add(directionOCRColumn);
                    DataColumn directionColumn = new DataColumn("DirectionPercent", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(directionColumn);
                    DataColumn streetNumberDBColumn = new DataColumn("StreetNo DB");
                    dtResults.Columns.Add(streetNumberDBColumn);
                    DataColumn streetNumberOCRColumn = new DataColumn("StreetNo OCR");
                    dtResults.Columns.Add(streetNumberOCRColumn);
                    DataColumn streetNumberColumn = new DataColumn("StreetNoPercent", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(streetNumberColumn);
                    DataColumn streetNoCharColumn = new DataColumn("StreetNoChar", Type.GetType("System.Int32"));
                    dtResults.Columns.Add(streetNoCharColumn);

                    DataTable AbbreviationTable = GetAbbreviationsTable(MDBOdbcConnectionObject);

                    for (int index = 0; index < CustomerRecordsFromDB.Rows.Count; index++)
                    {
                        DataRow drRow = CustomerRecordsFromDB.Rows[index];
                        DataRow drOutput = dtResults.NewRow();
                        clsCNCSBR clsCNCSBR = new clsCNCSBR();

                        drOutput["Sr.No."] = (index + 1);
                        drOutput["Zip DB"] = Convert.ToString(drRow["CMZIP"]);
                        drOutput["Zip OCR"] = AddressSectionsFromOCR.ZipCode;
                        drOutput["State DB"] = Convert.ToString(drRow["CMST"]);
                        drOutput["State OCR"] = AddressSectionsFromOCR.State;
                        drOutput["City DB"] = Convert.ToString(drRow["CMCITY"]);
                        drOutput["City OCR"] = AddressSectionsFromOCR.City;
                        drOutput["AddrLine1 DB"] = Convert.ToString(drRow["CMADR1"]);
                        drOutput["AddrLine1 OCR"] = AddressSectionsFromOCR.AddressLine1;
                        drOutput["Direction DB"] = Convert.ToString(drRow["CMADRU"]);
                        drOutput["Direction OCR"] = AddressSectionsFromOCR.Direction;
                        drOutput["StreetNo DB"] = Convert.ToString(drRow["CMADRN"]);
                        drOutput["StreetNo OCR"] = AddressSectionsFromOCR.StreetNumber;

                        // Compare Zip
                        if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMZIP"])) && !string.IsNullOrEmpty(AddressSectionsFromOCR.ZipCode))
                        {
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMZIP"]).Trim(), AddressSectionsFromOCR.ZipCode.Trim());
                            drOutput["ZipPercent"] = retStruct.PercentageMatch;
                        }
                        else
                        {
                            drOutput["ZipPercent"] = 0;
                        }
                        // Compare State
                        if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMST"])) && !string.IsNullOrEmpty(AddressSectionsFromOCR.State))
                        {
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMST"]).Trim(), AddressSectionsFromOCR.State.Trim());
                            drOutput["StatePercent"] = retStruct.PercentageMatch;
                        }
                        else
                        {
                            drOutput["StatePercent"] = 0;
                        }
                        // Compare City
                        if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMCITY"])) && !string.IsNullOrEmpty(AddressSectionsFromOCR.City))
                        {
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMCITY"]).Trim(), AddressSectionsFromOCR.City.Trim());
                            drOutput["CityPercent"] = retStruct.PercentageMatch;
                        }
                        else
                        {
                            drOutput["CityPercent"] = 0;
                        }
                        // Compare Address Line 1
                        if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMADR1"])) && !string.IsNullOrEmpty(AddressSectionsFromOCR.AddressLine1))
                        {
                            // CMADR2 - If DB record has address 1 and address 2 both, then for comparison, combine both the values and match with the OCR text
                            string dBValue = "";
                            if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMADR2"])))
                            {
                                dBValue = string.Concat(Convert.ToString(drRow["CMADR1"]).Trim(), " ", Convert.ToString(drRow["CMADR2"]).Trim());
                            }
                            else
                            {
                                dBValue = Convert.ToString(drRow["CMADR1"]).Trim();
                            }

                            // Before matching Address Line text, replace all abbreviations with full name and then compare
                            //dBValue = ReplaceAddressAbbreviation(dBValue.Trim(), MDBOdbcConnectionObject);
                            dBValue = ReplaceAddressAbbreviation(dBValue.Trim(), AbbreviationTable);
                            //string ocrValue = ReplaceAddressAbbreviation(AddressSectionsFromOCR.AddressLine1.Trim(), MDBOdbcConnectionObject);
                            string ocrValue = ReplaceAddressAbbreviation(AddressSectionsFromOCR.AddressLine1.Trim(), AbbreviationTable);

                            // Before matching remove the "house #" if present in the DB address line 1, since OCR address line 1 does not have house # section
                            string houseNumber = Convert.ToString(drRow["CMADRN"]).Trim();
                            if (!string.IsNullOrEmpty(houseNumber))
                            {
                                if (dBValue.Contains(houseNumber.Trim()))
                                {
                                    dBValue = dBValue.Replace(houseNumber, "").Trim();
                                    // Remove "Direction" keyword after "Street Number" => E / W / S / N / EAST / WEST / NORTH / SOUTH / N E / N W / S E / S W /
                                    string direction = Convert.ToString(drRow["CMADRU"]).Trim();
                                    if (!string.IsNullOrEmpty(direction))
                                    {
                                        if (dBValue.Contains(direction.Trim()))
                                        {
                                            dBValue = dBValue.Replace(direction, "").Trim();
                                        }
                                    }
                                }
                            }

                            // Remove "SUITE AND #" from DB value -> CMADR1 OR CMADR2 before comparison
                            if (dBValue.Contains("SUITE") || dBValue.Contains("STE"))
                            {
                                if (dBValue.Contains("SUITE"))
                                {
                                    int startIndex = dBValue.IndexOf("SUITE");
                                    if (startIndex >= 0)
                                    {
                                        string tempString = dBValue.Remove(startIndex, 6);

                                        int nextIndex = tempString.IndexOf(" ", startIndex);
                                        if (nextIndex >= 0)
                                        {
                                            int charToRemove = (nextIndex - startIndex) + 1;

                                            string finalString = tempString.Remove(startIndex, charToRemove);
                                            if (!string.IsNullOrEmpty(finalString))
                                            {
                                                dBValue = finalString;
                                            }
                                        }
                                    }
                                }
                                else if (dBValue.Contains("STE"))
                                {
                                    int startIndex = dBValue.IndexOf("STE");
                                    if (startIndex >= 0)
                                    {
                                        string tempString = dBValue.Remove(startIndex, 4);

                                        int nextIndex = tempString.IndexOf(" ", startIndex);
                                        if (nextIndex >= 0)
                                        {
                                            int charToRemove = (nextIndex - startIndex) + 1;

                                            string finalString = tempString.Remove(startIndex, charToRemove);
                                            if (!string.IsNullOrEmpty(finalString))
                                            {
                                                dBValue = finalString;
                                            }
                                        }
                                    }
                                }
                            }

                            // Remove "P O BOX AND #" from DB value -> CMADR1 OR CMADR2 before comparison
                            if (dBValue.Contains("P O BOX"))
                            {
                                int startIndex = dBValue.IndexOf("P O BOX");
                                if (startIndex >= 0)
                                {
                                    string tempString = dBValue.Remove(startIndex, 8);

                                    int nextIndex = tempString.IndexOf(" ", startIndex);
                                    if (nextIndex >= 0)
                                    {
                                        int charToRemove = (nextIndex - startIndex) + 1;

                                        string finalString = tempString.Remove(startIndex, charToRemove);
                                        if (!string.IsNullOrEmpty(finalString))
                                        {
                                            dBValue = finalString;
                                        }
                                    }
                                }
                            }

                            //RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMADR1"]).Trim(), AddressSectionsFromOCR.AddressLine1.Trim());
                            // Before comparing the address lines, remove all blank spaces and then compare
                            dBValue = Regex.Replace(dBValue, @"\s", ""); //dBValue.Replace(" ", String.Empty);
                            ocrValue = Regex.Replace(ocrValue, @"\s", ""); //ocrValue.Replace(" ", String.Empty);
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(dBValue, ocrValue);
                            drOutput["AddrLine1Percent"] = retStruct.PercentageMatch;
                        }
                        else
                        {
                            drOutput["AddrLine1Percent"] = 0;
                        }
                        // Compare Direction
                        if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMADRU"])) && !string.IsNullOrEmpty(AddressSectionsFromOCR.Direction))
                        {
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMADRU"]).Trim(), AddressSectionsFromOCR.Direction.Trim());
                            drOutput["DirectionPercent"] = retStruct.PercentageMatch;
                        }
                        else
                        {
                            drOutput["DirectionPercent"] = 0;
                        }
                        // Compare Street Number
                        if (!string.IsNullOrEmpty(Convert.ToString(drRow["CMADRN"])) && !string.IsNullOrEmpty(AddressSectionsFromOCR.StreetNumber))
                        {
                            RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMADRN"]).Trim(), AddressSectionsFromOCR.StreetNumber.Trim());
                            drOutput["StreetNoPercent"] = retStruct.PercentageMatch;
                            drOutput["StreetNoChar"] = retStruct.CharChanged;
                        }
                        else
                        {
                            drOutput["StreetNoPercent"] = 0;
                            drOutput["StreetNoChar"] = 0;
                        }
                        dtResults.Rows.Add(drOutput);
                    }
                }
            }
            catch
            {
                dtResults = null;
            }

            return dtResults;
        }

        public static DataTable MatchRuleEngine(DataTable SearchResults)
        {
            DataTable dtOutput = SearchResults.Clone();
            try
            {
                if (SearchResults != null && SearchResults.Rows.Count > 0)
                {
                    // Find the highest value of address matching percentage from all search results to be considered for valid address matching
                    int addressHighestMatchPercentage = Convert.ToInt32(SearchResults.Compute("Max(AddrLine1Percent)", ""));

                    for (int index = 0; index < SearchResults.Rows.Count; index++)
                    {
                        DataRow currentRow = SearchResults.Rows[index];

                        // If any one of these 6 rules is satisfied, return the matching records

                        // Rule 1 - Zip - 100%, State - 100%, City - 100%, Address Line 1 - 100%, Street# - 100%
                        if (Convert.ToInt32(currentRow["ZipPercent"]) == 100 && Convert.ToInt32(currentRow["StatePercent"]) == 100 &&
                            Convert.ToInt32(currentRow["CityPercent"]) == 100 && Convert.ToInt32(currentRow["AddrLine1Percent"]) == 100 && Convert.ToInt32(currentRow["StreetNoPercent"]) == 100)
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            //for(int colIndex = 0; index < currentRow.ItemArray.Length; colIndex++)
                            //{
                            //    dataRow[colIndex] = currentRow[colIndex];
                            //}
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 2 - Zip - 100%, State - 100%, City - 100%, Address Line 1 - 100%, Street# - 1 digit changed
                        else if (Convert.ToInt32(currentRow["ZipPercent"]) == 100 && Convert.ToInt32(currentRow["StatePercent"]) == 100 &&
                            Convert.ToInt32(currentRow["CityPercent"]) == 100 && Convert.ToInt32(currentRow["AddrLine1Percent"]) == 100 && Convert.ToInt32(currentRow["StreetNoChar"]) == 1)
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 3 - Zip - 100%, State - 100%, City - 100%, Address Line 1 >= 70 (and highest from all records), Street# - 100%
                        else if (Convert.ToInt32(currentRow["ZipPercent"]) == 100 && Convert.ToInt32(currentRow["StatePercent"]) == 100 &&
                            Convert.ToInt32(currentRow["CityPercent"]) == 100 &&
                            (Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                            Convert.ToInt32(currentRow["StreetNoPercent"]) == 100)
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 4 - Zip - 100%, State - 100%, City - BLANK, Address Line 1 >= 70 (and highest from all records), Street# - 100%
                        else if (Convert.ToInt32(currentRow["ZipPercent"]) == 100 && Convert.ToInt32(currentRow["StatePercent"]) == 100 &&
                             (Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                             Convert.ToInt32(currentRow["StreetNoPercent"]) == 100)
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 5 - Zip - 100%, State - BLANK, City - 100%, Address Line 1 >= 70 (and highest from all records), Street# - 100%
                        else if (Convert.ToInt32(currentRow["ZipPercent"]) == 100 && Convert.ToInt32(currentRow["CityPercent"]) == 100 &&
                            (Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                            Convert.ToInt32(currentRow["StreetNoPercent"]) == 100)
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 6 - Zip - BLANK, State - 100%, City - 100%, Address Line 1 >= 70 (and highest from all records), Street# - 100%
                        else if (Convert.ToInt32(currentRow["StatePercent"]) == 100 && Convert.ToInt32(currentRow["CityPercent"]) == 100 &&
                             (Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                             Convert.ToInt32(currentRow["StreetNoPercent"]) == 100)
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                    }
                }
            }
            catch
            {
                dtOutput = null;
            }
            return dtOutput;
        }

        #region FILTER CUSTOMER RECORDS METHODS
        private static DataTable FilterBaseResultsUsingCustomerName(DataTable BaseResults, string CustomerName)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMNAME] LIKE '%" + CustomerName.ToUpper() + "%'").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }

        private static DataTable FilterBaseResultsUsingZipCode(DataTable BaseResults, string ZipCode)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMZIP] LIKE '" + ZipCode.ToUpper() + "%'").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }

        private static DataTable FilterBaseResultsUsingCity(DataTable BaseResults, string City)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMCITY] LIKE '" + City.ToUpper() + "%'").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }

        private static DataTable FilterBaseResultsUsingStreetNumber(DataTable BaseResults, string StreetNumber)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMADRN] = " + StreetNumber + "").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }

        #endregion

        #region MAIN QUERY AND EXECUTION ON SAIA DATABASE METHODS
        private static string SearchCustomerUsingTerminalIdQuery(string TerminalID, SearchParameters.CustomerType CustomerType, AddressSection AddressSectionsFromOCR)
        {
            string customerClassCondition = "";
            if (CustomerType == SearchParameters.CustomerType.Shipper)
            {
                customerClassCondition = " AND CMCLS = 'F'";
            }
            else if (CustomerType == SearchParameters.CustomerType.Consignee)
            {
                customerClassCondition = " AND CMCLS = 'F'";
            }
            else if (CustomerType == SearchParameters.CustomerType.ThirdParty)
            {
                customerClassCondition = " AND (CMCLS = 'F' OR CMCLS = 'B')";
            }

            string query = string.Empty;
            // For Third Party customer, Terminal ID is not relevant, hence, do not consider Terminal ID for base data
            if (CustomerType == SearchParameters.CustomerType.ThirdParty)
            {
                query = "select * from ARP001 WHERE  CMSTAT = 'A'" + customerClassCondition + " AND CMST = '" + AddressSectionsFromOCR.State + "'";
            }
            else
            {
                query = "select * from ARP001 WHERE  CMSTAT = 'A' AND CMST= '" + AddressSectionsFromOCR.State + "' AND CMTID LIKE '" + TerminalID.ToUpper() + "%'" + customerClassCondition;
            }

            return query;
        }
        private static DataTable ExecuteSearchQueryOnARP001(string SearchQuery, OdbcConnection OdbcConnectionObject)
        {
            DataTable dtCustomer = new DataTable();

            try
            {
                //using (OdbcConnectionObject)
                {
                    using (OdbcCommand cmd = new OdbcCommand(SearchQuery, OdbcConnectionObject))
                    {
                        if (OdbcConnectionObject != null && !string.IsNullOrEmpty(SearchQuery))
                        {
                            cmd.CommandType = CommandType.Text;
                            using (OdbcDataAdapter da = new OdbcDataAdapter(cmd))
                            {
                                da.Fill(dtCustomer);
                            }
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                dtCustomer = null;
            }
            return dtCustomer;
        }
        #endregion

        #region MAIN METHOD
        public static DataTable SearchAndFetchMatchingCustomers(AddressSection AddressSectionsFromOCR, SearchParameters CustomerSearchParmaters,
            OdbcConnection SAIAOdbcConnectionObject, clsOleDBDataAccess MDBOdbcConnectionObject, SearchParameters.CustomerType CustomerAccountType,
            SearchParameters.SearchType SearchType)
        {
            DataTable dtOutput = new DataTable();

            DataTable AbbreviationTable = GetAbbreviationsTable(MDBOdbcConnectionObject);

            // If customer name in parameter contains hyphen "-", remove and join the words // Example => WAL-MART => WALMART
            bool isHyphenInCustomerName = false;
            string originalCustomerNameWithHyphen = "";
            string splittedCustomerNameWithoutHyphen = "";
            if (CustomerSearchParmaters.CustomerName.Contains("-"))
            {
                isHyphenInCustomerName = true;
                originalCustomerNameWithHyphen = CustomerSearchParmaters.CustomerName;
                splittedCustomerNameWithoutHyphen = string.Concat(originalCustomerNameWithHyphen.Split('-')[0], " ", originalCustomerNameWithHyphen.Split('-')[1]);
                CustomerSearchParmaters.CustomerName = Regex.Replace(CustomerSearchParmaters.CustomerName, @"-", "");
            }

            if (SearchType == SearchParameters.SearchType.Normal) // If normal account search - 4 Search parameters - Name OR ZipCode OR City OR House Number
            {
                try
                {
                    // For Third Party customer, Terminal ID is not relevant, hence, do not consider Terminal ID for base data
                    // Step 1 - Base data using "Terminal ID" to be used for search
                    // Added CustomerType as one more condition to fetch valid customer accounts from database
                    {
                        string searchBaseQuery = SearchCustomerUsingTerminalIdQuery(CustomerSearchParmaters.TerminalID, CustomerAccountType, AddressSectionsFromOCR);
                        if (!string.IsNullOrEmpty(searchBaseQuery))
                        {
                            //Commented on 06th Jan 2021
                            //if (dtStepOneOutput != null)
                            {
                                //Commented on 06th Jan 2021
                                //dtOutput = dtStepOneOutput.Clone();
                                // Step 2 - Generate a random number to proceed with either "Customer Name" search OR "Zip Code" search OR "City" Search
                                Random random = new Random();
                                double randomNumber = random.NextDouble();

                                bool isCustomerNameSearch = false;
                                bool isZipCodeSearch = false;
                                bool isCitySearch = false;

                                DataTable dtStepTwoOutput = new DataTable();

                                // Added on 14-Jan-2021 --start
                                string CustomerNameFormatted = string.Empty;
                                string tempBaseQuery = string.Empty;
                                string CustomerNameFormattedThreeLetter = string.Empty;
                                // Master Name = H D SUPPLY FACILITIES MAINTENA
                                // OCR Name =   HD SUPPLY FACILITIES MAINTENA  PROLINE ABC XYZ ==>  PROLINE 
                                // CustomerSearchParmaters.CustomerName = HD SUPPLY
                                string[] arrCustomerName = CustomerSearchParmaters.CustomerName.Split(' '); // query will be HD SUPPLY/ PROLINE  OR H D SUPPLY // 
                                bool isNamewithORcondition = false;
                                string QuerywithORcondition = string.Empty;
                                string QuerywithOUTORcondition = string.Empty;
                                if (arrCustomerName.Length <= 2) // check name is concat with two words  case 1 :-  HD SUPPLY case 2:-  HD SUPPLY 
                                {
                                    if (arrCustomerName[0].Length == 2) // split the concat word and check the len if it is 2 then add space in between 
                                    {
                                        char[] letters = arrCustomerName[0].Trim().ToCharArray();
                                        CustomerNameFormatted = letters[0] + " " + letters[1] + " " + arrCustomerName[1];
                                        isNamewithORcondition = true;
                                        QuerywithORcondition = " AND (CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%' OR CMNAME LIKE '" + CustomerNameFormatted.ToUpper() + "%')";

                                    }
                                    else // if splitted concat word len is  
                                    {
                                        isNamewithORcondition = false;
                                        if (isHyphenInCustomerName)
                                        {
                                            QuerywithOUTORcondition = " AND (CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%' OR CMNAME LIKE '" +
                                                splittedCustomerNameWithoutHyphen + "%')";
                                        }
                                        // added on 1-Feb-2021 when customer name is 3 letters then add space between each letter and form a query eg:- ORS => O R S
                                        else if (CustomerSearchParmaters.CustomerName.Trim().Length == 3)
                                        {
                                            isNamewithORcondition = false;
                                            char[] letters = CustomerSearchParmaters.CustomerName.ToUpper().Trim().ToCharArray();
                                            CustomerNameFormattedThreeLetter = letters[0] + " " + letters[1] + " " + letters[2];
                                            QuerywithOUTORcondition = " AND (CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%' OR CMNAME LIKE '" +
                                                    CustomerNameFormattedThreeLetter + "%')";

                                        } //end
                                        else
                                        {
                                            QuerywithOUTORcondition = " AND CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%'";
                                        }
                                    }
                                }
                                else
                                {
                                    isNamewithORcondition = false;
                                    if (isHyphenInCustomerName)
                                    {
                                        QuerywithOUTORcondition = " AND (CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%' OR CMNAME LIKE '" +
                                            splittedCustomerNameWithoutHyphen + "%')";
                                    }
                                    // added on 1-Feb-2021 when customer name is 3 letters then add space between each letter and form a query eg:- ORS => O R S
                                    else if (CustomerSearchParmaters.CustomerName.Trim().Length == 3)
                                    {
                                        isNamewithORcondition = false;
                                        char[] letters = CustomerSearchParmaters.CustomerName.ToUpper().Trim().ToCharArray();
                                        CustomerNameFormattedThreeLetter = letters[0] + " " + letters[1] + " " + letters[2];
                                        QuerywithOUTORcondition = " AND (CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%' OR CMNAME LIKE '" +
                                                CustomerNameFormattedThreeLetter + "%')";
                                    }
                                    else
                                    {
                                        QuerywithOUTORcondition = " AND CMNAME LIKE '" + CustomerSearchParmaters.CustomerName.ToUpper() + "%'";
                                    }
                                }
                                // Added on 14-Jan-2021 --// Added on 14-Jan-2021 --start

                                // Customer Name / Zip code / City Search => First pass
                                if (randomNumber <= 0.5 && !string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName)) // Customer Name Search
                                {
                                    isCustomerNameSearch = true;

                                    // Added on 14-Jan-2021 -- start
                                    if (isNamewithORcondition == true)
                                    {
                                        tempBaseQuery = searchBaseQuery + QuerywithORcondition;
                                    }
                                    else
                                    {
                                        tempBaseQuery = searchBaseQuery + QuerywithOUTORcondition;
                                    }

                                    // Added on 14-Jan-2021 -- end
                                    dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                }
                                else if (randomNumber > 0.5 && !string.IsNullOrEmpty(CustomerSearchParmaters.ZipCode)) // Zip Code Search
                                {
                                    isZipCodeSearch = true;
                                    tempBaseQuery = searchBaseQuery + " AND CMZIP LIKE '" + CustomerSearchParmaters.ZipCode.ToUpper() + "%'";
                                    dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                }
                                else if (!string.IsNullOrEmpty(CustomerSearchParmaters.City)) // City Search - IF no value for search - CustName/ZipCode
                                {
                                    isCitySearch = true;
                                    tempBaseQuery = searchBaseQuery + " AND CMCITY LIKE '" + CustomerSearchParmaters.City.ToUpper() + "%'";
                                    dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                }
                                // Using the output of step 2, match and find out if any matched customer data is returned, if not, then go with next parameter search                           
                                if (dtStepTwoOutput != null && dtStepTwoOutput.Rows != null && dtStepTwoOutput.Rows.Count > 0)
                                {
                                    //// added on 06th Jan 2021
                                    //dtOutput = dtStepTwoOutput.Clone();
                                    // Add for-loop and parse each searched customer address with address from OCR and check if it is matched
                                    // Add the record to output - if matched

                                    dtOutput = FetchFinalMatchedCustomer(dtStepTwoOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerSearchParmaters, SAIAOdbcConnectionObject,
                                        CustomerAccountType, isNamewithORcondition, isHyphenInCustomerName, CustomerNameFormatted, splittedCustomerNameWithoutHyphen, CustomerNameFormattedThreeLetter);

                                    #region Commented on 06th Feb 2021
                                    //for (int rowIndex = 0; rowIndex < dtStepTwoOutput.Rows.Count; rowIndex++)
                                    //{
                                    //    DataRow currentDBRow = dtStepTwoOutput.Rows[rowIndex];

                                    //    AddressSection objDBAddress = new AddressSection();
                                    //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                    //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                    //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                    //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                    //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                    //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                    //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                    //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                    //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                    //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                    //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                    //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                    //    if (!string.IsNullOrEmpty(matchResults))
                                    //    {
                                    //        if (matchResults.Split('~')[0] == "Y")
                                    //        {

                                    //            #region If street number && address && state code with city or zip match then pick the current row as it is with out matching name
                                    //            // Check - Street# == 100%, Address Line 1 >= 70%, State = 100%, Zip = 100% OR City = 100%
                                    //            int ZipPercentage = Convert.ToInt32(matchResults.Split('~')[1].Split('#')[1]);
                                    //            int StatePercentage = Convert.ToInt32(matchResults.Split('~')[2].Split('#')[1]);
                                    //            int CityPercentage = Convert.ToInt32(matchResults.Split('~')[3].Split('#')[1]);
                                    //            int AddressLine1Percentage = Convert.ToInt32(matchResults.Split('~')[4].Split('#')[1]);
                                    //            int streetPercentage = Convert.ToInt32(matchResults.Split('~')[5].Split('#')[1]);
                                    //            if (streetPercentage == 100 && AddressLine1Percentage >= 70 && StatePercentage == 100 && CityPercentage == 100 && ZipPercentage == 100)
                                    //            {
                                    //                DataRow dataRow = dtOutput.NewRow();
                                    //                dataRow.ItemArray = currentDBRow.ItemArray;
                                    //                dtOutput.Rows.Add(dataRow);
                                    //                // Check for alise case
                                    //                string strQueryAlise = "select * from ARP001 WHERE  CMSTAT = 'A' AND CMST ='" + objDBAddress.State + "' AND CMZIP='" + objDBAddress.ZipCode + "' AND CMADR1 like '" + objProcessAddressParameters.AddressLine1 + "%' AND CMCITY like '" + objDBAddress.City + "%'";
                                    //                DataTable dtAlise = new DataTable();
                                    //                dtAlise = ExecuteSearchQueryOnARP001(strQueryAlise, SAIAOdbcConnectionObject);
                                    //                if (dtAlise != null && dtAlise.Rows.Count > 1)
                                    //                {
                                    //                    if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                                    //                    {
                                    //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseShipperName = "Y";
                                    //                    }
                                    //                    else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                                    //                    {
                                    //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseConsigneeName = "Y";
                                    //                    }
                                    //                    else
                                    //                    {
                                    //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseBillToName = "Y";
                                    //                    }
                                    //                }

                                    //            }
                                    //            #endregion
                                    //            else
                                    //            {
                                    //                if (!string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName))
                                    //                {
                                    //                    if (isNamewithORcondition == true)
                                    //                    {
                                    //                        if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormatted.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                    //                        {
                                    //                            DataRow dataRow = dtOutput.NewRow();
                                    //                            dataRow.ItemArray = currentDBRow.ItemArray;
                                    //                            dtOutput.Rows.Add(dataRow);
                                    //                        }
                                    //                    }
                                    //                    else
                                    //                    {
                                    //                        if (isHyphenInCustomerName) // Added for special case S-One on BOL and S One in DB, check both values "SONE" AND "S ONE"
                                    //                        {
                                    //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                    //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(splittedCustomerNameWithoutHyphen.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                    //                            {
                                    //                                DataRow dataRow = dtOutput.NewRow();
                                    //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                    //                                dtOutput.Rows.Add(dataRow);
                                    //                            }
                                    //                        }
                                    //                        else if (CustomerSearchParmaters.CustomerName.Trim().Length == 3) // added on 1-Feb-2021 to handle case like ORS then send both words like ORS and O R S
                                    //                        {
                                    //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                    //                                    Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormattedThreeLetter.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                    //                            {
                                    //                                DataRow dataRow = dtOutput.NewRow();
                                    //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                    //                                dtOutput.Rows.Add(dataRow);
                                    //                            }
                                    //                        }
                                    //                        else
                                    //                        {
                                    //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                    //                            {
                                    //                                DataRow dataRow = dtOutput.NewRow();
                                    //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                    //                                dtOutput.Rows.Add(dataRow);
                                    //                            }
                                    //                        }
                                    //                    }

                                    //                }
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    #endregion
                                }

                                // One parameter search did not give result, then check for other paramater for search - Customer Name / Zip code / City
                                // Customer Name / Zip code / City Search => Second pass
                                if (dtOutput == null || (dtOutput != null && dtOutput.Rows.Count == 0))
                                {
                                    if (isCustomerNameSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName)) // Customer Name Search
                                    {
                                        isCustomerNameSearch = true;
                                        // Added on 14-Jan-2021 -- start
                                        if (isNamewithORcondition == true)
                                        {
                                            tempBaseQuery = searchBaseQuery + QuerywithORcondition;
                                        }
                                        else
                                        {
                                            tempBaseQuery = searchBaseQuery + QuerywithOUTORcondition;
                                        }

                                        // Added on 14-Jan-2021 -- end
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isZipCodeSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.ZipCode)) // Zip Code Search
                                    {
                                        isZipCodeSearch = true;
                                        tempBaseQuery = searchBaseQuery + " AND CMZIP LIKE '" + CustomerSearchParmaters.ZipCode.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isCitySearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.City)) // City Search - IF no value for search - CustName/ZipCode
                                    {
                                        isCitySearch = true;
                                        tempBaseQuery = searchBaseQuery + " AND CMCITY LIKE '" + CustomerSearchParmaters.City.ToUpper() + "%'"; // commented on 14-Jan-2021                                      
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }

                                    // Using the output of step 2, match and find out if any matched customer data is returned, if not, then go with next parameter search
                                    if (dtStepTwoOutput != null && dtStepTwoOutput.Rows != null && dtStepTwoOutput.Rows.Count > 0)
                                    {
                                        //// added on 06th Jan 2021
                                        //dtOutput = dtStepTwoOutput.Clone();

                                        dtOutput = FetchFinalMatchedCustomer(dtStepTwoOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerSearchParmaters, SAIAOdbcConnectionObject,
                                        CustomerAccountType, isNamewithORcondition, isHyphenInCustomerName, CustomerNameFormatted, splittedCustomerNameWithoutHyphen, CustomerNameFormattedThreeLetter);

                                        #region Commented on 06th Feb 2021
                                        //for (int rowIndex = 0; rowIndex < dtStepTwoOutput.Rows.Count; rowIndex++)
                                        //{
                                        //    DataRow currentDBRow = dtStepTwoOutput.Rows[rowIndex];

                                        //    AddressSection objDBAddress = new AddressSection();
                                        //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                        //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                        //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                        //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                        //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                        //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                        //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                        //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                        //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                        //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                        //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                        //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                        //    if (!string.IsNullOrEmpty(matchResults))
                                        //    {
                                        //        if (matchResults.Split('~')[0] == "Y")
                                        //        {
                                        //            #region If street number && address && state code with city or zip match then pick the current row as it is with out matching name
                                        //            // Check - Street# == 100%, Address Line 1 >= 70%, State = 100%, Zip = 100% OR City = 100%
                                        //            int ZipPercentage = Convert.ToInt32(matchResults.Split('~')[1].Split('#')[1]);
                                        //            int StatePercentage = Convert.ToInt32(matchResults.Split('~')[2].Split('#')[1]);
                                        //            int CityPercentage = Convert.ToInt32(matchResults.Split('~')[3].Split('#')[1]);
                                        //            int AddressLine1Percentage = Convert.ToInt32(matchResults.Split('~')[4].Split('#')[1]);
                                        //            int streetPercentage = Convert.ToInt32(matchResults.Split('~')[5].Split('#')[1]);
                                        //            if (streetPercentage == 100 && AddressLine1Percentage >= 70 && StatePercentage == 100 && CityPercentage == 100 && ZipPercentage == 100)
                                        //            {
                                        //                DataRow dataRow = dtOutput.NewRow();
                                        //                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                dtOutput.Rows.Add(dataRow);
                                        //                // Check for alise case
                                        //                string strQueryAlise = "select * from ARP001 WHERE  CMSTAT = 'A' AND CMST ='" + objDBAddress.State + "' AND CMZIP='" + objDBAddress.ZipCode + "' AND CMADR1 like '" + objProcessAddressParameters.AddressLine1 + "%' AND CMCITY like '" + objDBAddress.City + "%'";
                                        //                DataTable dtAlise = new DataTable();
                                        //                dtAlise = ExecuteSearchQueryOnARP001(strQueryAlise, SAIAOdbcConnectionObject);
                                        //                if (dtAlise != null && dtAlise.Rows.Count > 1)
                                        //                {
                                        //                    if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseShipperName = "Y";
                                        //                    }
                                        //                    else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseConsigneeName = "Y";
                                        //                    }
                                        //                    else
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseBillToName = "Y";
                                        //                    }
                                        //                }

                                        //            }
                                        //            #endregion
                                        //            else
                                        //            {
                                        //                if (!string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName))
                                        //                {
                                        //                    if (isNamewithORcondition == true)
                                        //                    {
                                        //                        if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormatted.ToUpper()))
                                        //                        {
                                        //                            DataRow dataRow = dtOutput.NewRow();
                                        //                            dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                            dtOutput.Rows.Add(dataRow);
                                        //                        }
                                        //                    }
                                        //                    else
                                        //                    {
                                        //                        if (isHyphenInCustomerName) // Added for special case S-One on BOL and S One in DB, check both values "SONE" AND "S ONE"
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                        //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(splittedCustomerNameWithoutHyphen.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                        else if (CustomerSearchParmaters.CustomerName.Trim().Length == 3) // added on 1-Feb-2021 to handle case like ORS then send both words like ORS and O R S
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                        //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormattedThreeLetter.ToUpper()))
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                        else
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()))
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                    }
                                        //                }
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        #endregion
                                    }
                                }

                                // One parameter search did not give result, then check for other paramater for search - Customer Name / Zip code / City
                                // Customer Name / Zip code / City Search => Third pass
                                if (dtOutput == null || (dtOutput != null && dtOutput.Rows.Count == 0))
                                {
                                    if (isCustomerNameSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName)) // Customer Name Search
                                    {
                                        isCustomerNameSearch = true;
                                        // Added on 14-Jan-2021 -- start
                                        if (isNamewithORcondition == true)
                                        {
                                            tempBaseQuery = searchBaseQuery + QuerywithORcondition;
                                        }
                                        else
                                        {
                                            tempBaseQuery = searchBaseQuery + QuerywithOUTORcondition;
                                        }

                                        // Added on 14-Jan-2021 -- end
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isZipCodeSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.ZipCode)) // Zip Code Search
                                    {
                                        isZipCodeSearch = true;
                                        tempBaseQuery = searchBaseQuery + " AND CMZIP LIKE '" + CustomerSearchParmaters.ZipCode.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isCitySearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.City)) // City Search - IF no value for search - CustName/ZipCode
                                    {
                                        isCitySearch = true;
                                        tempBaseQuery = searchBaseQuery + " AND CMCITY LIKE '" + CustomerSearchParmaters.City.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }

                                    // Using the output of step 2, match and find out if any matched customer data is returned, if not, then go with next parameter search
                                    if (dtStepTwoOutput != null && dtStepTwoOutput.Rows != null && dtStepTwoOutput.Rows.Count > 0)
                                    {
                                        //// added on 06th Jan 2021
                                        //dtOutput = dtStepTwoOutput.Clone();

                                        dtOutput = FetchFinalMatchedCustomer(dtStepTwoOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerSearchParmaters, SAIAOdbcConnectionObject,
                                        CustomerAccountType, isNamewithORcondition, isHyphenInCustomerName, CustomerNameFormatted, splittedCustomerNameWithoutHyphen, CustomerNameFormattedThreeLetter);

                                        #region Commented on 06th Feb 2021
                                        //for (int rowIndex = 0; rowIndex < dtStepTwoOutput.Rows.Count; rowIndex++)
                                        //{
                                        //    DataRow currentDBRow = dtStepTwoOutput.Rows[rowIndex];

                                        //    AddressSection objDBAddress = new AddressSection();
                                        //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                        //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                        //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                        //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                        //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                        //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                        //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                        //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                        //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                        //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                        //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                        //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                        //    if (!string.IsNullOrEmpty(matchResults))
                                        //    {
                                        //        if (matchResults.Split('~')[0] == "Y")
                                        //        {
                                        //            #region If street number && address && state code with city or zip match then pick the current row as it is with out matching name
                                        //            // Check - Street# == 100%, Address Line 1 >= 70%, State = 100%, Zip = 100% OR City = 100%
                                        //            int ZipPercentage = Convert.ToInt32(matchResults.Split('~')[1].Split('#')[1]);
                                        //            int StatePercentage = Convert.ToInt32(matchResults.Split('~')[2].Split('#')[1]);
                                        //            int CityPercentage = Convert.ToInt32(matchResults.Split('~')[3].Split('#')[1]);
                                        //            int AddressLine1Percentage = Convert.ToInt32(matchResults.Split('~')[4].Split('#')[1]);
                                        //            int streetPercentage = Convert.ToInt32(matchResults.Split('~')[5].Split('#')[1]);
                                        //            if (streetPercentage == 100 && AddressLine1Percentage >= 70 && StatePercentage == 100 && CityPercentage == 100 && ZipPercentage == 100)
                                        //            {
                                        //                DataRow dataRow = dtOutput.NewRow();
                                        //                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                dtOutput.Rows.Add(dataRow);
                                        //                // Check for alise case
                                        //                string strQueryAlise = "select * from ARP001 WHERE  CMSTAT = 'A' AND CMST ='" + objDBAddress.State + "' AND CMZIP='" + objDBAddress.ZipCode + "' AND CMADR1 like '" + objProcessAddressParameters.AddressLine1 + "%' AND CMCITY like '" + objDBAddress.City + "%'";
                                        //                DataTable dtAlise = new DataTable();
                                        //                dtAlise = ExecuteSearchQueryOnARP001(strQueryAlise, SAIAOdbcConnectionObject);
                                        //                if (dtAlise != null && dtAlise.Rows.Count > 1)
                                        //                {
                                        //                    if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseShipperName = "Y";
                                        //                    }
                                        //                    else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseConsigneeName = "Y";
                                        //                    }
                                        //                    else
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseBillToName = "Y";
                                        //                    }
                                        //                }

                                        //            }
                                        //            #endregion
                                        //            else
                                        //            {
                                        //                if (!string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName))
                                        //                {
                                        //                    if (isNamewithORcondition == true)
                                        //                    {
                                        //                        if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormatted.ToUpper()))
                                        //                        {
                                        //                            DataRow dataRow = dtOutput.NewRow();
                                        //                            dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                            dtOutput.Rows.Add(dataRow);
                                        //                        }
                                        //                    }
                                        //                    else
                                        //                    {
                                        //                        if (isHyphenInCustomerName) // Added for special case S-One on BOL and S One in DB, check both values "SONE" AND "S ONE"
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                        //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(splittedCustomerNameWithoutHyphen.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                        else if (CustomerSearchParmaters.CustomerName.Trim().Length == 3) // added on 1-Feb-2021 to handle case like ORS then send both words like ORS and O R S
                                        //                        {

                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                        //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormattedThreeLetter.ToUpper())) // ORS or O R S 
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                        else
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()))
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                    }

                                        //                }
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        #endregion
                                    }
                                }

                                // Step 3 - Proceed with "Street Number" search
                                if (dtOutput == null || (dtOutput != null && dtOutput.Rows.Count == 0))
                                {
                                    DataTable dtStepThreeOutput = new DataTable();

                                    // Street Number
                                    if (!string.IsNullOrEmpty(CustomerSearchParmaters.StreetNumber)) // Street Number Search
                                    {
                                        tempBaseQuery = searchBaseQuery + " AND CMADRN = " + CustomerSearchParmaters.StreetNumber + "";
                                        dtStepThreeOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    // Using the output of step 3, match and find out if any matched customer data is returned, if not, then return empty                         
                                    if (dtStepThreeOutput != null && dtStepThreeOutput.Rows != null && dtStepThreeOutput.Rows.Count > 0)
                                    {
                                        //// added on 06th Jan 2021
                                        //dtOutput = dtStepThreeOutput.Clone();

                                        dtOutput = FetchFinalMatchedCustomer(dtStepThreeOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerSearchParmaters, SAIAOdbcConnectionObject,
                                        CustomerAccountType, isNamewithORcondition, isHyphenInCustomerName, CustomerNameFormatted, splittedCustomerNameWithoutHyphen, CustomerNameFormattedThreeLetter);

                                        #region Commented on 06th Feb 2021
                                        //for (int rowIndex = 0; rowIndex < dtStepThreeOutput.Rows.Count; rowIndex++)
                                        //{
                                        //    DataRow currentDBRow = dtStepThreeOutput.Rows[rowIndex];

                                        //    AddressSection objDBAddress = new AddressSection();
                                        //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                        //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                        //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                        //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                        //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                        //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                        //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                        //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                        //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                        //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                        //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                        //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                        //    if (!string.IsNullOrEmpty(matchResults))
                                        //    {
                                        //        if (matchResults.Split('~')[0] == "Y")
                                        //        {
                                        //            #region If street number && address && state code with city or zip match then pick the current row as it is with out matching name
                                        //            // Check - Street# == 100%, Address Line 1 >= 70%, State = 100%, Zip = 100% OR City = 100%
                                        //            int ZipPercentage = Convert.ToInt32(matchResults.Split('~')[1].Split('#')[1]);
                                        //            int StatePercentage = Convert.ToInt32(matchResults.Split('~')[2].Split('#')[1]);
                                        //            int CityPercentage = Convert.ToInt32(matchResults.Split('~')[3].Split('#')[1]);
                                        //            int AddressLine1Percentage = Convert.ToInt32(matchResults.Split('~')[4].Split('#')[1]);
                                        //            int streetPercentage = Convert.ToInt32(matchResults.Split('~')[5].Split('#')[1]);
                                        //            if (streetPercentage == 100 && AddressLine1Percentage >= 70 && StatePercentage == 100 && CityPercentage == 100 && ZipPercentage == 100)
                                        //            {
                                        //                DataRow dataRow = dtOutput.NewRow();
                                        //                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                dtOutput.Rows.Add(dataRow);
                                        //                // Check for alise case
                                        //                string strQueryAlise = "select * from ARP001 WHERE  CMSTAT = 'A' AND CMST ='" + objDBAddress.State + "' AND CMZIP='" + objDBAddress.ZipCode + "' AND CMADR1 like '" + objProcessAddressParameters.AddressLine1 + "%' AND CMCITY like '" + objDBAddress.City + "%'";
                                        //                DataTable dtAlise = new DataTable();
                                        //                dtAlise = ExecuteSearchQueryOnARP001(strQueryAlise, SAIAOdbcConnectionObject);
                                        //                if (dtAlise != null && dtAlise.Rows.Count > 1)
                                        //                {
                                        //                    if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseShipperName = "Y";
                                        //                    }
                                        //                    else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseConsigneeName = "Y";
                                        //                    }
                                        //                    else
                                        //                    {
                                        //                        clsF27.ClsSearchEngin.ClsSearchEnginAliseBillToName = "Y";
                                        //                    }
                                        //                }

                                        //            }
                                        //            #endregion
                                        //            else
                                        //            {
                                        //                if (!string.IsNullOrEmpty(CustomerSearchParmaters.CustomerName))
                                        //                {
                                        //                    if (isNamewithORcondition == true)
                                        //                    {
                                        //                        if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormatted.ToUpper()))
                                        //                        {
                                        //                            DataRow dataRow = dtOutput.NewRow();
                                        //                            dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                            dtOutput.Rows.Add(dataRow);
                                        //                        }
                                        //                    }
                                        //                    else
                                        //                    {
                                        //                        if (isHyphenInCustomerName) // Added for special case S-One on BOL and S One in DB, check both values "SONE" AND "S ONE"
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                        //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(splittedCustomerNameWithoutHyphen.ToUpper())) // HD SUPPLY or H D SUPPLY  
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                        else if (CustomerSearchParmaters.CustomerName.Trim().Length == 3) // added on 1-Feb-2021 to handle case like ORS then send both words like ORS and O R S
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()) ||
                                        //                                Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerNameFormattedThreeLetter.ToUpper())) // ORS or O R S  
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                        else
                                        //                        {
                                        //                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains(CustomerSearchParmaters.CustomerName.ToUpper()))
                                        //                            {
                                        //                                DataRow dataRow = dtOutput.NewRow();
                                        //                                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                                dtOutput.Rows.Add(dataRow);
                                        //                            }
                                        //                        }
                                        //                    }
                                        //                }
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        #endregion
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    dtOutput = null;
                }

                // If output contains data, and if selected customer type is "Third Party", check if there are 2 records for same criteria - one with type "F" and other "B"
                // If present, return only "B" type customer record
                if (CustomerAccountType == SearchParameters.CustomerType.ThirdParty && dtOutput != null && dtOutput.Rows.Count > 1)
                {
                    DataRow[] filteredData = dtOutput.Select("[CMCLS] = 'B'");
                    if (filteredData != null && filteredData.Count() > 0)
                    {
                        dtOutput = filteredData.CopyToDataTable();
                    }
                }
            }
            else // If Blind Shipper account search - 3 Search Parameters - ZipCode OR City OR House Number - Only Address Search
            {
                try
                {
                    // Step 1 - Base data using "Terminal ID" to be used for search
                    // Added CustomerType as one more condition to fetch valid customer accounts from database
                    if (!string.IsNullOrEmpty(CustomerSearchParmaters.TerminalID))
                    {
                        string searchBaseQuery = SearchCustomerUsingTerminalIdQuery(CustomerSearchParmaters.TerminalID, CustomerAccountType, AddressSectionsFromOCR);
                        if (!string.IsNullOrEmpty(searchBaseQuery))
                        {
                            {
                                // Step 2 - Generate a random number to proceed with either "Zip Code" search OR "City" Search OR "House Number" Search
                                Random random = new Random();
                                double randomNumber = random.NextDouble();

                                bool isHouseNumberSearch = false;
                                bool isZipCodeSearch = false;
                                bool isCitySearch = false;

                                DataTable dtStepTwoOutput = new DataTable();

                                // Zip code / City / Street Number => First pass
                                if (randomNumber <= 0.5 && !string.IsNullOrEmpty(CustomerSearchParmaters.ZipCode)) // Zip Code Search
                                {
                                    isZipCodeSearch = true;
                                    string tempBaseQuery = searchBaseQuery + " AND CMZIP LIKE '" + CustomerSearchParmaters.ZipCode.ToUpper() + "%'";
                                    dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                }
                                else if (randomNumber > 0.5 && !string.IsNullOrEmpty(CustomerSearchParmaters.City)) // City Search 
                                {
                                    isCitySearch = true;
                                    string tempBaseQuery = searchBaseQuery + " AND CMCITY LIKE '" + CustomerSearchParmaters.City.ToUpper() + "%'";
                                    dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                }
                                else if (!string.IsNullOrEmpty(CustomerSearchParmaters.StreetNumber)) // Street Number Search
                                {
                                    isHouseNumberSearch = true;
                                    string tempBaseQuery = searchBaseQuery + " AND CMADRN = " + CustomerSearchParmaters.StreetNumber + "";
                                    dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                }
                                // Using the output of step 2, match and find out if any matched customer data is returned, if not, then go with next parameter search                           
                                if (dtStepTwoOutput != null)
                                {
                                    #region Commented on 06th Feb 2021
                                    //// added on 06th Jan 2021
                                    //dtOutput = dtStepTwoOutput.Clone();

                                    //for (int rowIndex = 0; rowIndex < dtStepTwoOutput.Rows.Count; rowIndex++)
                                    //{
                                    //    DataRow currentDBRow = dtStepTwoOutput.Rows[rowIndex];

                                    //    AddressSection objDBAddress = new AddressSection();
                                    //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                    //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                    //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                    //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                    //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                    //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                    //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                    //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                    //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                    //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                    //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                    //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                    //    if (!string.IsNullOrEmpty(matchResults))
                                    //    {
                                    //        if (matchResults.Split('~')[0] == "Y")
                                    //        {
                                    //            if (Convert.ToString(currentDBRow["CMNAME"]).Contains("(BS)"))
                                    //            {
                                    //                DataRow dataRow = dtOutput.NewRow();
                                    //                dataRow.ItemArray = currentDBRow.ItemArray;
                                    //                dtOutput.Rows.Add(dataRow);
                                    //            }
                                    //        }
                                    //    }
                                    //}
                                    #endregion
                                    dtOutput = FetchFinalMatchedBlindShipper(dtStepTwoOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerAccountType);
                                }

                                // One parameter search did not give result, then check for other paramater for search - Zip code / City / Street Number
                                // Zip code / City / Street Number => Second pass
                                if (dtOutput == null || (dtOutput != null && dtOutput.Rows.Count == 0))
                                {
                                    if (isZipCodeSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.ZipCode)) // Zip Code Search
                                    {
                                        isZipCodeSearch = true;
                                        string tempBaseQuery = searchBaseQuery + " AND CMZIP LIKE '" + CustomerSearchParmaters.ZipCode.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isCitySearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.City)) // City Search
                                    {
                                        isCitySearch = true;
                                        string tempBaseQuery = searchBaseQuery + " AND CMCITY LIKE '" + CustomerSearchParmaters.City.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isHouseNumberSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.StreetNumber)) //  Street Number Search
                                    {
                                        isHouseNumberSearch = true;
                                        string tempBaseQuery = searchBaseQuery + " AND CMADRN = " + CustomerSearchParmaters.StreetNumber + "";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }

                                    // Using the output of step 2, match and find out if any matched customer data is returned, if not, then go with next parameter search
                                    if (dtStepTwoOutput != null)
                                    {
                                        #region Commented on 06th Feb 2021
                                        //// added on 06th Jan 2021
                                        //dtOutput = dtStepTwoOutput.Clone();

                                        //for (int rowIndex = 0; rowIndex < dtStepTwoOutput.Rows.Count; rowIndex++)
                                        //{
                                        //    DataRow currentDBRow = dtStepTwoOutput.Rows[rowIndex];

                                        //    AddressSection objDBAddress = new AddressSection();
                                        //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                        //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                        //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                        //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                        //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                        //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                        //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                        //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                        //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                        //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                        //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                        //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                        //    if (!string.IsNullOrEmpty(matchResults))
                                        //    {
                                        //        if (matchResults.Split('~')[0] == "Y")
                                        //        {
                                        //            if (Convert.ToString(currentDBRow["CMNAME"]).Contains("(BS)"))
                                        //            {
                                        //                DataRow dataRow = dtOutput.NewRow();
                                        //                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                dtOutput.Rows.Add(dataRow);
                                        //            }
                                        //        }
                                        //    }
                                        //}

                                        #endregion
                                        dtOutput = FetchFinalMatchedBlindShipper(dtStepTwoOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerAccountType);
                                    }
                                }

                                // One parameter search did not give result, then check for other paramater for search - Zip code / City / Street Number
                                // Zip code / City / Street Number => Third pass
                                if (dtOutput == null || (dtOutput != null && dtOutput.Rows.Count == 0))
                                {
                                    if (isZipCodeSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.ZipCode)) // Zip Code Search
                                    {
                                        isZipCodeSearch = true;
                                        string tempBaseQuery = searchBaseQuery + " AND CMZIP LIKE '" + CustomerSearchParmaters.ZipCode.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isCitySearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.City)) // City Search
                                    {
                                        isCitySearch = true;
                                        string tempBaseQuery = searchBaseQuery + " AND CMCITY LIKE '" + CustomerSearchParmaters.City.ToUpper() + "%'";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }
                                    else if (isHouseNumberSearch == false && !string.IsNullOrEmpty(CustomerSearchParmaters.StreetNumber)) //  Street Number Search
                                    {
                                        isHouseNumberSearch = true;
                                        string tempBaseQuery = searchBaseQuery + " AND CMADRN = " + CustomerSearchParmaters.StreetNumber + "";
                                        dtStepTwoOutput = ExecuteSearchQueryOnARP001(tempBaseQuery, SAIAOdbcConnectionObject);
                                    }

                                    // Using the output of step 2, match and find out if any matched customer data is returned, if not, then go with next parameter search
                                    if (dtStepTwoOutput != null)
                                    {
                                        #region Commented on 06th Feb 2021
                                        //// added on 06th Jan 2021
                                        //dtOutput = dtStepTwoOutput.Clone();

                                        //for (int rowIndex = 0; rowIndex < dtStepTwoOutput.Rows.Count; rowIndex++)
                                        //{
                                        //    DataRow currentDBRow = dtStepTwoOutput.Rows[rowIndex];

                                        //    AddressSection objDBAddress = new AddressSection();
                                        //    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                                        //    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                                        //    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                                        //    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                                        //    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                                        //    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                                        //    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                                        //    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                                        //    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                                        //    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                                        //    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                                        //    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                                        //    if (!string.IsNullOrEmpty(matchResults))
                                        //    {
                                        //        if (matchResults.Split('~')[0] == "Y")
                                        //        {
                                        //            if (Convert.ToString(currentDBRow["CMNAME"]).Contains("(BS)"))
                                        //            {
                                        //                DataRow dataRow = dtOutput.NewRow();
                                        //                dataRow.ItemArray = currentDBRow.ItemArray;
                                        //                dtOutput.Rows.Add(dataRow);
                                        //            }
                                        //        }
                                        //    }
                                        //}
                                        #endregion

                                        dtOutput = FetchFinalMatchedBlindShipper(dtStepTwoOutput, AbbreviationTable, AddressSectionsFromOCR, CustomerAccountType);
                                    }
                                }
                            }
                        }
                    }
                }
                catch
                {
                    dtOutput = null;
                }
            }
            return dtOutput;
        }
        #endregion

        #region MAIN ADDRESS MATCHING METHOD
        public static string MatchAddressUsingRules(AddressSection AddressSectionsFromDB, AddressSection AddressSectionsFromOCR, SearchParameters.CustomerType CustomerAccountType,
            DataTable AbbreviationTable)
        {
            // Final result will be string containing the matching percentage of the address sections AND result as "Matched" Or "Not Matched"
            string finalMatchedResults;
            try
            {
                clsCNCSBR clsCNCSBR = new clsCNCSBR();
                StringBuilder tempResults = new StringBuilder();

                int zipCodePercentMatch = 0;
                int statePercentMatch = 0;
                int cityPercentMatch = 0;
                int addrLine1PercentMatch = 0;
                int streetNumberPercentMatch = 0;
                int streetCharChanged = 0;
                // Remove all special characters except space "\s"
                AddressSectionsFromOCR.AddressLine1 = Regex.Replace(AddressSectionsFromOCR.AddressLine1, "[^a-zA-Z0-9\\s%/]", ""); // Added on 9-Feb-2021
                AddressSectionsFromOCR.AddressLine1 = ReplaceAddressAbbreviation(AddressSectionsFromOCR.AddressLine1, AbbreviationTable); // Added on 9-Feb-2021 Eg=> 2727 S W 13TH ST = 2727 S W 13TH ST for this expanded STREET 
                #region COMPARE ADDRESS SECTIONS
                // Compare Zip
                if (!string.IsNullOrEmpty(AddressSectionsFromDB.ZipCode) && !string.IsNullOrEmpty(AddressSectionsFromOCR.ZipCode))
                {
                    RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(AddressSectionsFromDB.ZipCode.Trim(), AddressSectionsFromOCR.ZipCode.Trim());
                    zipCodePercentMatch = retStruct.PercentageMatch;
                }
                else
                {
                    zipCodePercentMatch = 0;
                }
                // Compare State
                if (!string.IsNullOrEmpty(AddressSectionsFromDB.State) && !string.IsNullOrEmpty(AddressSectionsFromOCR.State))
                {
                    RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(AddressSectionsFromDB.State.Trim(), AddressSectionsFromOCR.State.Trim());
                    statePercentMatch = retStruct.PercentageMatch;
                }
                else
                {
                    statePercentMatch = 0;
                }
                // Compare City
                if (!string.IsNullOrEmpty(AddressSectionsFromDB.City) && !string.IsNullOrEmpty(AddressSectionsFromOCR.City))
                {
                    RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(AddressSectionsFromDB.City.Trim(), AddressSectionsFromOCR.City.Trim());
                    cityPercentMatch = retStruct.PercentageMatch;
                }
                else
                {
                    cityPercentMatch = 0;
                }
                // Compare Address Line 1
                if (!string.IsNullOrEmpty(AddressSectionsFromDB.AddressLine1) && !string.IsNullOrEmpty(AddressSectionsFromOCR.AddressLine1))
                {
                    // Before comparing the address lines, remove all blank spaces and then compare
                    string dbValue = Regex.Replace(AddressSectionsFromDB.AddressLine1, @"\s", ""); //AddressSectionsFromDB.AddressLine1.Replace(" ", String.Empty);
                    string ocrValue = Regex.Replace(AddressSectionsFromOCR.AddressLine1, @"\s", ""); //AddressSectionsFromOCR.AddressLine1.Replace(" ", String.Empty);

                    RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(dbValue.Trim(), ocrValue.Trim());
                    addrLine1PercentMatch = retStruct.PercentageMatch;

                    // If the matching percentage is less than threshold 70, then, check by removing the expanded abbreviations and then compare both the address lines
                    if (addrLine1PercentMatch < 70)
                    {
                        string dbAddrLine1 = RemoveExpandedAbbreviation(AddressSectionsFromDB.AddressLine1.Trim(), AbbreviationTable);
                        string ocrAddrLine1 = RemoveExpandedAbbreviation(AddressSectionsFromOCR.AddressLine1.Trim(), AbbreviationTable);

                        if (!string.IsNullOrEmpty(dbAddrLine1) && !string.IsNullOrEmpty(ocrAddrLine1))
                        {
                            string tempDBAddressLine1 = dbAddrLine1;
                            string tempOCRAddressLine1 = ocrAddrLine1;

                            // Before comparing the address lines, remove all blank spaces and then compare
                            dbAddrLine1 = Regex.Replace(dbAddrLine1, @"\s", ""); //dbAddrLine1.Replace(" ", String.Empty);
                            ocrAddrLine1 = Regex.Replace(ocrAddrLine1, @"\s", ""); //ocrAddrLine1.Replace(" ", String.Empty);

                            retStruct = clsCNCSBR.IQSBR050(dbAddrLine1, ocrAddrLine1);
                            addrLine1PercentMatch = retStruct.PercentageMatch;

                            // Case 3 =>  // If the matching percentage is less than threshold 70, then, remove all full numeric words and then compare both the address lines
                            if (addrLine1PercentMatch < 70)
                            {
                                string tempDBAddressLine1WithDigit = tempDBAddressLine1;
                                string tempOCRAddressLine1WithDigit = tempOCRAddressLine1;
                                // Remove numeric words from DB address
                                string pattern = string.Format(@"\b{0}\b", "[\\d]+");
                                bool contains = Regex.IsMatch(tempDBAddressLine1, pattern);
                                if (contains)
                                {
                                    string strWord = Regex.Replace(tempDBAddressLine1, pattern, "");
                                    tempDBAddressLine1 = strWord;
                                }

                                // Remove numeric words from OCR address
                                string pattern1 = string.Format(@"\b{0}\b", "[\\d]+");
                                bool contains1 = Regex.IsMatch(tempOCRAddressLine1, pattern1);
                                if (contains1)
                                {
                                    string strWord = Regex.Replace(tempOCRAddressLine1, pattern1, "");
                                    tempOCRAddressLine1 = strWord;
                                }

                                if (!string.IsNullOrEmpty(tempDBAddressLine1) && !string.IsNullOrEmpty(tempOCRAddressLine1))
                                {
                                    // Before removing blank spaces, copy the value in another variable to be used for keyword present check when match % is less than 70
                                    string tempDBAddressLine1Before = tempDBAddressLine1;
                                    string tempOCRAddressLine1Before = tempOCRAddressLine1;

                                    // Before comparing the address lines, remove all blank spaces and then compare
                                    tempDBAddressLine1 = Regex.Replace(tempDBAddressLine1, @"\s", ""); //tempDBAddressLine1.Replace(" ", String.Empty);
                                    tempOCRAddressLine1 = Regex.Replace(tempOCRAddressLine1, @"\s", ""); //tempOCRAddressLine1.Replace(" ", String.Empty);

                                    retStruct = clsCNCSBR.IQSBR050(tempDBAddressLine1, tempOCRAddressLine1);
                                    addrLine1PercentMatch = retStruct.PercentageMatch;

                                    // Compare words within address - If any keyword from OCR address found in DB , assign the address percentage to 70
                                    if (addrLine1PercentMatch < 70)
                                    {
                                        // Remove all special characters except space "\s"
                                        tempDBAddressLine1Before = Regex.Replace(tempDBAddressLine1Before, "[^a-zA-Z\\s]", "");
                                        tempOCRAddressLine1Before = Regex.Replace(tempOCRAddressLine1Before, "[^a-zA-Z\\s]", "");

                                        string[] dbWords = tempDBAddressLine1Before.Split(' ');
                                        string[] ocrWords = tempOCRAddressLine1Before.Split(' ');

                                        // Remove empty spaces
                                        dbWords = dbWords.Where(item => !string.IsNullOrEmpty(item)).ToArray();
                                        ocrWords = ocrWords.Where(item => !string.IsNullOrEmpty(item)).ToArray();

                                        bool isWordFound = false;
                                        if (dbWords.Length > 0 && ocrWords.Length > 0)
                                        {
                                            for (int wordIndex1 = 0; wordIndex1 < dbWords.Length; wordIndex1++)
                                            {
                                                // DB address word to check should be of Minimum length = 3
                                                if (dbWords[wordIndex1].Length >= 3)
                                                {
                                                    for (int wordIndex2 = 0; wordIndex2 < ocrWords.Length; wordIndex2++)
                                                    {
                                                        // OCR address word to compare should be of Minimum length = 3
                                                        if (ocrWords[wordIndex2].Length >= 3)
                                                        {
                                                            if (dbWords[wordIndex1].ToUpper() == ocrWords[wordIndex2].ToUpper())
                                                            {
                                                                isWordFound = true;
                                                                addrLine1PercentMatch = 70;
                                                                break;
                                                            }
                                                        }
                                                    }
                                                }
                                                if (isWordFound)
                                                    break;
                                            }
                                            // As discussed, removed the case on 03rd Feb 2021, since incorrect address was matched using only basis as Street #
                                            //// added on 2-feb-2021 OCRAddress = 740 N WLSE DR BLDG B   ZLP Sumter SC 29153  => ParseAddressLine1 = WLSE DR BLDG B ZLP => MatchAddressUsingRules= WLSEDRIVEBLDGBZLP 
                                            ////MasterAddress = 740 N WISE DR SUMTER SC 29153  => ParseAddressLine1 = WISE DR => MatchAddressUsingRules = WISEDRIVE
                                            //if (!isWordFound)
                                            //{
                                            //    if (!string.IsNullOrEmpty(AddressSectionsFromDB.StreetNumber) && !string.IsNullOrEmpty(AddressSectionsFromOCR.StreetNumber))
                                            //    {
                                            //        RetStructIQSBR050 retStruct1 = clsCNCSBR.IQSBR050(AddressSectionsFromDB.StreetNumber.Trim(), AddressSectionsFromOCR.StreetNumber.Trim());
                                            //        if (retStruct1.PercentageMatch == 100)
                                            //        {
                                            //            isWordFound = true;
                                            //            addrLine1PercentMatch = 70;
                                            //        }

                                            //    }
                                            //}

                                            
                                        }

                                        // Added on 19-march-2021 -- Start 
                                        if (!isWordFound && addrLine1PercentMatch < 70)
                                        {
                                            // check numeric words from DB address
                                            string patternWithDigit = string.Format(@"\b{0}\b", "[\\d]+");
                                            bool containsWithDigit = Regex.IsMatch(tempDBAddressLine1WithDigit, patternWithDigit);
                                            // check numeric words from OCR address
                                            string pattern1WithDigit = string.Format(@"\b{0}\b", "[\\d]+");
                                            bool contains1withDigit = Regex.IsMatch(tempOCRAddressLine1WithDigit, pattern1WithDigit);
                                            // if both contains digit then check only if digit match in address line
                                            if (containsWithDigit && contains1withDigit)
                                            {
                                                string[] dbWordsWithDigit = tempDBAddressLine1WithDigit.Split(' ');
                                                string[] ocrWordsWithDigit = tempOCRAddressLine1WithDigit.Split(' ');

                                                // Remove empty spaces
                                                dbWordsWithDigit = dbWordsWithDigit.Where(item => !string.IsNullOrEmpty(item)).ToArray();
                                                ocrWordsWithDigit = ocrWordsWithDigit.Where(item => !string.IsNullOrEmpty(item)).ToArray();
                                                bool isWordFoundWithDigit = false;
                                                if (dbWordsWithDigit.Length > 0 && ocrWordsWithDigit.Length > 0)
                                                {
                                                    for (int wordIndex1 = 0; wordIndex1 < dbWordsWithDigit.Length; wordIndex1++)
                                                    {
                                                        // DB address word should be digit only
                                                        if (Regex.IsMatch(dbWordsWithDigit[wordIndex1].ToString().Trim(), patternWithDigit) == true)
                                                        {
                                                            for (int wordIndex2 = 0; wordIndex2 < ocrWordsWithDigit.Length; wordIndex2++)
                                                            {
                                                                // OCR address word should be digit only
                                                                if (Regex.IsMatch(ocrWordsWithDigit[wordIndex2].ToString().Trim(), pattern1WithDigit) == true)
                                                                {
                                                                    if (dbWordsWithDigit[wordIndex1].ToUpper() == ocrWordsWithDigit[wordIndex2].ToUpper())
                                                                    {
                                                                        isWordFoundWithDigit = true;
                                                                        addrLine1PercentMatch = 70;
                                                                        break;
                                                                    }
                                                                }
                                                            }
                                                        }
                                                        if (isWordFoundWithDigit)
                                                            break;
                                                    }
                                                   
                                                }
                                            }

                                        }

                                        // Added on 19-march-2021 -- end
                                    }
                                }
                                else
                                {
                                    addrLine1PercentMatch = 0;
                                }
                            }
                        }
                        else
                        {
                            addrLine1PercentMatch = 0;
                        }
                    }
                }
                else
                {
                    addrLine1PercentMatch = 0;
                }
                // Compare Street Number
                if (!string.IsNullOrEmpty(AddressSectionsFromDB.StreetNumber) && !string.IsNullOrEmpty(AddressSectionsFromOCR.StreetNumber))
                {
                    RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(AddressSectionsFromDB.StreetNumber.Trim(), AddressSectionsFromOCR.StreetNumber.Trim());
                    streetNumberPercentMatch = retStruct.PercentageMatch;
                    streetCharChanged = retStruct.CharChanged;
                }
                else
                {
                    streetNumberPercentMatch = 0;
                    streetCharChanged = 0;
                }
                #endregion

                #region EXECUTE MATCH RULES
                //Execute Match Rules AND append in output as "matched" or "not matched"

                // Rule 1 - Street# - 100%, Address Line 1 - 100%, State - 100%, Zip - 100% OR City - 100%
                if (streetNumberPercentMatch == 100 && addrLine1PercentMatch == 100 &&
                    statePercentMatch == 100 && (zipCodePercentMatch == 100 || cityPercentMatch == 100))
                {
                    tempResults.Append("Y~");
                }
                // Rule 2 - Street# - 100%, Address Line 1 >= 70 (and highest from all records), State - 100%, Zip - 100% OR City - 100%
                else if (streetNumberPercentMatch == 100 && (addrLine1PercentMatch == 100 || addrLine1PercentMatch >= 70) &&
                    statePercentMatch == 100 && (zipCodePercentMatch == 100 || cityPercentMatch == 100))
                {
                    tempResults.Append("Y~");
                }
                #region Commented on 08th Feb 2021 -> Not required
                //// Rule 3 - Street# - 1 digit changed, Address Line 1 - 100%, State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                //else if (streetCharChanged == 1 && addrLine1PercentMatch == 100 &&
                //    statePercentMatch == 100 && (zipCodePercentMatch == 100 || cityPercentMatch == 100))
                //{
                //    //if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty) // Commented on 4-Feb-2021 for shipper as Street # needs to be an exact match
                //    //if (CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                //    //{
                //    //    tempResults.Append("Y~");
                //    //}
                //    //else
                //    //{
                //    tempResults.Append("N~");
                //    //}
                //}
                //// Rule 4 - Street# - 1 digit changed, Address Line 1 >= 70 (and highest from all records), State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                //else if (streetCharChanged == 1 &&
                //     (addrLine1PercentMatch == 100 || addrLine1PercentMatch >= 70) &&
                //    statePercentMatch == 100 && (zipCodePercentMatch == 100 || cityPercentMatch == 100))
                //{
                //    //if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                //    //if (CustomerAccountType == SearchParameters.CustomerType.ThirdParty) // Commented on 4-Feb-2021 for shipper as Street # needs to be an exact match
                //    //{
                //    //    tempResults.Append("Y~");
                //    //}
                //    //else
                //    //{
                //    tempResults.Append("N~");
                //    // }
                //}
                //// Rule 5 - Address Line 1 - 100%, State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                //else if (addrLine1PercentMatch == 100 &&
                //    statePercentMatch == 100 && (zipCodePercentMatch == 100 || cityPercentMatch == 100))
                //{
                //    //if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty) // Commented on 4-Feb-2021 for shipper as Street # needs to be an exact match
                //    //if (CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                //    //{
                //    //    tempResults.Append("Y~");
                //    //}
                //    //else
                //    //{
                //    tempResults.Append("N~");
                //    //}
                //}
                //// Rule 6 - Address Line 1 >= 70 (and highest from all records), State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                //else if ((addrLine1PercentMatch == 100 || addrLine1PercentMatch >= 70) &&
                //    statePercentMatch == 100 && (zipCodePercentMatch == 100 || cityPercentMatch == 100))
                //{
                //    //if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty) // Commented on 4-Feb-2021 for shipper as Street # needs to be an exact match
                //    //if (CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                //    //{
                //    //    tempResults.Append("Y~");
                //    //}
                //    //else
                //    //{
                //    tempResults.Append("N~");
                //    //}
                //}
                #endregion
                else
                {
                    tempResults.Append("N~");
                }

                #endregion

                #region PREPARE RETURN RESULT STRING
                // Append all the address sections percentage match results and send as output string
                tempResults.Append(string.Concat("ZipCode#", zipCodePercentMatch, "~"));
                tempResults.Append(string.Concat("State#", statePercentMatch, "~"));
                tempResults.Append(string.Concat("City#", cityPercentMatch, "~"));
                tempResults.Append(string.Concat("AddressLine1#", addrLine1PercentMatch, "~"));
                tempResults.Append(string.Concat("StreetNumber#", streetNumberPercentMatch, "~"));
                tempResults.Append(string.Concat("StreetCharChanged#", streetCharChanged, "~"));
                #endregion 

                finalMatchedResults = Convert.ToString(tempResults);
            }
            catch
            {
                finalMatchedResults = string.Empty;
            }
            return finalMatchedResults;
        }
        #endregion

        #region ADDRESS MATCHING RELATED SUPPORTING METHODS
        public static string ReplaceAddressAbbreviation(string AddressLine1, DataTable AbbreviationTable)
        {
            string updatedValue = AddressLine1;
            try
            {
                //string cmdString = "select * from USPostalAbbreviations";
                //OdbcConnectionObject.propConnection.Open();
                //DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USPostalAbbreviations");
                if (AbbreviationTable != null && AbbreviationTable.Rows != null && AbbreviationTable.Rows.Count > 0)
                {
                    // Check if the current address line text contains any one of the postal abbreviations
                    // if found, replace the abbreviation with the actual name          
                    string tempAddress = AddressLine1;
                    foreach (DataRow drRow in AbbreviationTable.Rows)
                    {
                        string pattern = string.Format(@"\b{0}\b", Convert.ToString(drRow["Abbreviation"]));//  string.Concat("@", "\b", Convert.ToString(drRow["Abbreviation"]), "\b");
                        bool contains = Regex.IsMatch(tempAddress, pattern);
                        if (contains)
                        {
                            string strWord = Regex.Replace(tempAddress, pattern, Convert.ToString(drRow["Name"]));
                            tempAddress = strWord;
                        }
                    }
                    updatedValue = tempAddress;
                }
                //OdbcConnectionObject.propConnection.Close();
                //OdbcConnectionObject = null;
            }
            catch
            {
                updatedValue = AddressLine1;
            }
            return updatedValue;
        }
        public static string ProcessAddressBeforeComparison(ProcessAddressParameters AddressParameters, DataTable AbbreviationTable)
        {
            // CMADR2 - If DB record has address 1 and address 2 both, then for comparison, combine both the values and match with the OCR text
            string dBValue = "";
            if (!string.IsNullOrEmpty(Convert.ToString(AddressParameters.AddressLine2)))
            {
                //dBValue = string.Concat(Convert.ToString(AddressParameters.AddressLine1).Trim(), " ", Convert.ToString(AddressParameters.AddressLine2).Trim()); // Commented on 10-Feb-2021 and added below if..else block 
                if (Convert.ToString(AddressParameters.AddressLine1).Trim().StartsWith("%") || Convert.ToString(AddressParameters.AddressLine1).Trim().StartsWith("C/O")) // added on 10-Feb-2021
                {
                    dBValue = Convert.ToString(AddressParameters.AddressLine2).Trim(); // added on 10-Feb-2021
                }
                else
                {
                    dBValue = string.Concat(Convert.ToString(AddressParameters.AddressLine1).Trim(), " ", Convert.ToString(AddressParameters.AddressLine2).Trim());
                }
            }
            else
            {
                dBValue = Convert.ToString(AddressParameters.AddressLine1).Trim();
            }

            // Before matching Address Line text, replace all abbreviations with full name and then compare
            dBValue = ReplaceAddressAbbreviation(dBValue.Trim(), AbbreviationTable);

            // Before matching remove the "house #" if present in the DB address line 1, since OCR address line 1 does not have house # section
            // string houseNumber = Convert.ToString(AddressParameters.StreetNumber).Trim(); // commented and added below on 26-Dec-2020 to handle exception
            string houseNumber = string.Empty;
            if (!string.IsNullOrEmpty(AddressParameters.StreetNumber))
            {
                houseNumber = Convert.ToString(AddressParameters.StreetNumber).Trim();
            }
            if (!string.IsNullOrEmpty(houseNumber))
            {
                if (dBValue.Contains(houseNumber.Trim()))
                {
                    dBValue = dBValue.Replace(houseNumber, "").Trim();
                    // Remove "Direction" keyword after "Street Number" => E / W / S / N / EAST / WEST / NORTH / SOUTH / N E / N W / S E / S W /
                    if (!string.IsNullOrEmpty(AddressParameters.StreetDirection))
                    {
                        string direction = Convert.ToString(AddressParameters.StreetDirection).Trim();
                        if (!string.IsNullOrEmpty(direction))
                        {
                            //if (dBValue.Contains(direction.Trim()))
                            {
                                //dBValue = dBValue.Replace(direction, "").Trim();

                                string pattern = string.Format(@"\b{0}\b", direction);
                                bool contains = Regex.IsMatch(dBValue, pattern);
                                if (contains)
                                {
                                    string strWord = Regex.Replace(dBValue, pattern, "");
                                    dBValue = strWord;
                                }

                            }
                        }
                    }
                }
            }

            // Remove "SUITE AND #" from DB value -> CMADR1 OR CMADR2 before comparison
            string pattern1 = string.Format(@"\b{0}\b", "SUITE");
            bool contains1 = Regex.IsMatch(dBValue, pattern1);

            string pattern2 = string.Format(@"\b{0}\b", "STE");
            bool contains2 = Regex.IsMatch(dBValue, pattern2);

            //if (dBValue.Contains("SUITE") || dBValue.Contains("STE"))
            if (contains1 == true || contains2 == true)
            {
                //if (dBValue.Contains("SUITE"))
                if (contains1 == true)
                {
                    int startIndex = dBValue.IndexOf("SUITE");
                    if (startIndex >= 0)
                    {
                        string temp = dBValue.Substring(startIndex + "SUITE".Length).Trim();
                        string tempString = dBValue.Remove(startIndex, 6);

                        string[] parts = temp.Split(' ');
                        string value = parts[0];

                        string pattern = string.Format(@"\b{0}\b", value);
                        bool contains = Regex.IsMatch(tempString, pattern);
                        if (contains)
                        {
                            string strWord = Regex.Replace(tempString, pattern, "");
                            dBValue = strWord;
                        }

                        //int nextIndex = tempString.IndexOf(" ", startIndex);
                        //if (nextIndex >= 0)
                        //{
                        //    int charToRemove = (nextIndex - startIndex) + 1;

                        //    string finalString = tempString.Remove(startIndex, charToRemove);
                        //    if (!string.IsNullOrEmpty(finalString))
                        //    {
                        //        dBValue = finalString;
                        //    }
                        //}
                    }
                }
                //else if (dBValue.Contains("STE"))
                else if (contains2 == true)
                {
                    int startIndex = dBValue.IndexOf("STE");
                    if (startIndex >= 0)
                    {
                        string temp = dBValue.Substring(startIndex + "STE".Length).Trim();
                        string tempString = dBValue.Remove(startIndex, 4);

                        string[] parts = temp.Split(' ');
                        string value = parts[0];

                        string pattern = string.Format(@"\b{0}\b", value);
                        bool contains = Regex.IsMatch(tempString, pattern);
                        if (contains)
                        {
                            string strWord = Regex.Replace(tempString, pattern, "");
                            dBValue = strWord;
                        }

                        //int nextIndex = tempString.IndexOf(" ", startIndex);
                        //if (nextIndex >= 0)
                        //{
                        //    int charToRemove = (nextIndex - startIndex) + 1;

                        //    string finalString = tempString.Remove(startIndex, charToRemove);
                        //    if (!string.IsNullOrEmpty(finalString))
                        //    {
                        //        dBValue = finalString;
                        //    }
                        //}
                    }
                }
            }

            // Commented PO BOX removal code as discussed on 28th Jan 2021 => will keep PO Box and number for address comparison
            //// Remove "P O BOX AND #" from DB value -> CMADR1 OR CMADR2 before comparison
            //string pattern3 = string.Format(@"\b{0}\b", "P O BOX");
            //bool contains3 = Regex.IsMatch(dBValue, pattern3);

            ////if (dBValue.Contains("P O BOX"))
            //if (contains3 == true)
            //{
            //    int startIndex = dBValue.IndexOf("P O BOX");
            //    if (startIndex >= 0)
            //    {
            //        string tempString = dBValue.Remove(startIndex, 8);

            //        int nextIndex = tempString.IndexOf(" ", startIndex);
            //        if (nextIndex >= 0)
            //        {
            //            int charToRemove = (nextIndex - startIndex) + 1;

            //            string finalString = tempString.Remove(startIndex, charToRemove);
            //            if (!string.IsNullOrEmpty(finalString))
            //            {
            //                dBValue = finalString;
            //            }
            //        }
            //        else // If no space found after P O BOX Number in DB Value, then directly remove the characters after start index till end of string length - may be only "Number" is present
            //        {
            //            //int charToRemove = (nextIndex - startIndex) + 1;

            //            //string finalString = tempString.Remove(startIndex, charToRemove);
            //            //if (!string.IsNullOrEmpty(finalString))
            //            //{
            //            //    dBValue = finalString;
            //            //}
            //        }
            //    }
            //}

            //RetStructIQSBR050 retStruct = clsCNCSBR.IQSBR050(Convert.ToString(drRow["CMADR1"]).Trim(), AddressSectionsFromOCR.AddressLine1.Trim());
            // Before comparing the address lines, remove all blank spaces and then compare
            //dBValue = dBValue.Replace(" ", String.Empty);
            return dBValue;
        }
        public static DataTable GetAbbreviationsTable(clsOleDBDataAccess OdbcConnectionObject)
        {
            string cmdString = "select * from USPostalAbbreviations";
            OdbcConnectionObject.propConnection.Open();
            DataTable dataTable = OdbcConnectionObject.ReturnDataTable(cmdString, "USPostalAbbreviations");
            OdbcConnectionObject.propConnection.Close();
            return dataTable;
        }
        public static string RemoveExpandedAbbreviation(string AddressLine1, DataTable AbbreviationTable)
        {
            string updatedValue = AddressLine1;
            try
            {
                if (AbbreviationTable != null && AbbreviationTable.Rows != null && AbbreviationTable.Rows.Count > 0)
                {
                    // Check if the current address line text contains any one of the postal expanded abbreviations
                    // if found, replace the abbreviation with the actual name          
                    string tempAddress = AddressLine1;
                    foreach (DataRow drRow in AbbreviationTable.Rows)
                    {
                        string pattern = string.Format(@"\b{0}\b", Convert.ToString(drRow["Name"]));
                        bool contains = Regex.IsMatch(tempAddress, pattern);
                        if (contains)
                        {
                            string strWord = Regex.Replace(tempAddress, pattern, "");
                            tempAddress = strWord;
                        }
                    }
                    updatedValue = tempAddress;
                }
            }
            catch
            {
                updatedValue = AddressLine1;
            }
            return updatedValue;
        }

        public static DataTable FetchFinalMatchedCustomer(DataTable MatchResults, DataTable AbbreviationTable, AddressSection AddressSectionsFromOCR,
            SearchParameters CustomerSearchParmaters, OdbcConnection SAIAOdbcConnectionObject, SearchParameters.CustomerType CustomerAccountType,
            bool IsNamewithORcondition, bool IsHyphenInCustomerName, string CustomerNameFormatted, string SplittedCustomerNameWithoutHyphen, string CustomerNameFormattedThreeLetter)
        {
            DataTable dtOutput = new DataTable();
            dtOutput = MatchResults.Clone();
            // create temp data table and store those records where address matched
            DataTable dtOutputTemp = new DataTable();
            dtOutputTemp = dtOutput.Clone();
            // set variables
            clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = string.Empty; // F = Correct code FOUND , N = Correct code NOT FOUND , A = Alias code FOUND , M = Multiple code FOUND
            //bool isAllMatchExceptName = false;
            // STEP :- 1 => Filter and Get those record where ONLY address is matched
            for (int rowIndex = 0; rowIndex < MatchResults.Rows.Count; rowIndex++)
            {
                try
                {
                    DataRow currentDBRow = MatchResults.Rows[rowIndex];
                    //if (Convert.ToString(currentDBRow["CMCUST"]) == "0930102") // temp added
                    //{ 
                    AddressSection objDBAddress = new AddressSection();
                    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                    #region  Added on 11-Feb-2021 -- start => Temp code commented
                    //    => if we found PO BOX in DB address then set street number as zero for OCR Address 
                    //if ((!string.IsNullOrEmpty(objProcessAddressParameters.AddressLine1) 
                    //    && objProcessAddressParameters.AddressLine1.Contains("P O BOX") 
                    //    && string.IsNullOrEmpty(objProcessAddressParameters.AddressLine2) ) || 
                    //    (!string.IsNullOrEmpty(objProcessAddressParameters.AddressLine2) 
                    //    && objProcessAddressParameters.AddressLine2.Contains("P O BOX") 
                    //    && string.IsNullOrEmpty(objProcessAddressParameters.AddressLine1))
                    //    )
                    //{
                    //    AddressSectionsFromOCR.StreetNumber = "0";
                    //}
                    #endregion
                    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                    // Added on 08th Feb 2021 ==> Check if the Street Number is 0 or blank, 
                    // add a check in CMADR1 AND/OR CMADR2 to find first numeric word and consider it as Street Number for comparison
                    if (string.IsNullOrEmpty(objDBAddress.StreetNumber) || objDBAddress.StreetNumber == "0")
                    {
                        bool isStreetNumberFound = false;
                        if (!string.IsNullOrEmpty(objProcessAddressParameters.AddressLine1))
                        {
                            bool isStreetNumberPresent = Regex.IsMatch(objProcessAddressParameters.AddressLine1.Split(' ')[0], @"[\d]+");
                            if (isStreetNumberPresent)
                            {
                                isStreetNumberFound = true;
                                objDBAddress.StreetNumber = objProcessAddressParameters.AddressLine1.Split(' ')[0];
                            }
                        }
                        if (isStreetNumberFound == false && !string.IsNullOrEmpty(objProcessAddressParameters.AddressLine2))
                        {
                            bool isStreetNumberPresent = Regex.IsMatch(objProcessAddressParameters.AddressLine2.Split(' ')[0], @"[\d]+");
                            if (isStreetNumberPresent)
                            {
                                isStreetNumberFound = true;
                                objDBAddress.StreetNumber = objProcessAddressParameters.AddressLine2.Split(' ')[0];
                            }
                        }
                    }

                    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                    if (!string.IsNullOrEmpty(matchResults))
                    {
                        if (matchResults.Split('~')[0] == "Y")
                        {
                            #region Commented => If street number && address && state code with city or zip match then pick the current row as it is with out matching name
                            //// Check - Street# == 100%, Address Line 1 >= 70%, State = 100%, Zip = 100% OR City = 100%
                            //int ZipPercentage = Convert.ToInt32(matchResults.Split('~')[1].Split('#')[1]);
                            //int StatePercentage = Convert.ToInt32(matchResults.Split('~')[2].Split('#')[1]);
                            //int CityPercentage = Convert.ToInt32(matchResults.Split('~')[3].Split('#')[1]);
                            //int AddressLine1Percentage = Convert.ToInt32(matchResults.Split('~')[4].Split('#')[1]);
                            //int streetPercentage = Convert.ToInt32(matchResults.Split('~')[5].Split('#')[1]);
                            //if (streetPercentage == 100 && AddressLine1Percentage >= 70 && StatePercentage == 100 && CityPercentage == 100 && ZipPercentage == 100)
                            //{
                            //    isAllMatchExceptName = true;
                            //}
                            #endregion
                            // add records in datatable where only address is match
                            DataRow dataRowTemp = dtOutputTemp.NewRow();
                            dataRowTemp.ItemArray = currentDBRow.ItemArray;
                            dtOutputTemp.Rows.Add(dataRowTemp);
                        }
                    }
                    //}
                }
                catch
                {
                    continue;
                }
            }
            #region added on 4-Aug-2021 => To reduce false positive fetch record from DB using AddressLine1, City, State and Zip
            if (dtOutputTemp.Rows.Count > 0)
            {
                // Added on 18-Nov-2021 -- Start
                string customerClassCondition = "";
                if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                {
                    customerClassCondition = " AND CMCLS = 'F'";
                }
                else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                {
                    customerClassCondition = " AND CMCLS = 'F'";
                }
                else if (CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                {
                    customerClassCondition = " AND (CMCLS = 'F' OR CMCLS = 'B')";
                }
                // Added on 18-Nov-2021 -- End
                //string strQueryAliasForSingleRow = string.Empty;
                string strQueryForSameAddress = "select * from ARP001 WHERE CMSTAT = 'A'" + customerClassCondition + " AND CMADR1 = '" + dtOutputTemp.Rows[0]["CMADR1"].ToString().Trim() + "' AND CMCITY= '" + dtOutputTemp.Rows[0]["CMCITY"].ToString().Trim() + "' AND CMST = '" + dtOutputTemp.Rows[0]["CMST"].ToString().Trim() + "' AND CMZIP ='" + dtOutputTemp.Rows[0]["CMZIP"].ToString().Trim() + "'";
                DataTable dtForSameAddress = new DataTable();
                dtForSameAddress = ExecuteSearchQueryOnARP001(strQueryForSameAddress, SAIAOdbcConnectionObject);
                if (dtForSameAddress != null && dtForSameAddress.Rows.Count > 0)
                {
                    for (int rowIndex = 0; rowIndex < dtForSameAddress.Rows.Count; rowIndex++)
                    {
                        try
                        {
                            DataRow currentDBRowSA = dtForSameAddress.Rows[rowIndex];
                            DataRow[] foundSA = dtOutputTemp.Select("CMCUST = '" + Convert.ToString(currentDBRowSA["CMCUST"]).Trim() + "'");
                            if (foundSA.Length == 0)
                            {
                                DataRow dataRowTempSA = dtOutputTemp.NewRow();
                                dataRowTempSA.ItemArray = currentDBRowSA.ItemArray;
                                dtOutputTemp.Rows.Add(dataRowTempSA);
                            }
                        }
                        catch
                        {
                            continue;
                        }
                    }
                }
            }
            #endregion
            // STEP :- 2 => if we found zero record then return datatable as null
            if (dtOutputTemp.Rows.Count == 0)
            {
                dtOutput = null;
            }
            // STEP :- 3 => If we found ONLY ONE record with address match then return that single record
            else if (dtOutputTemp.Rows.Count == 1)
            {
                // STEP :- 3_0 => if we get signle record from ARP001 then check for alias record in ARP033 using customer code
                string strQueryAliasForSingleRow = "select * from ARP033 WHERE NACUST ='" + dtOutputTemp.Rows[0]["CMCUST"].ToString().Trim() + "' ";
                DataTable dtAliasForSingleRow = new DataTable();
                dtAliasForSingleRow = ExecuteSearchQueryOnARP001(strQueryAliasForSingleRow, SAIAOdbcConnectionObject);

                // STEP :- 3_0_1 =>  Check ARP033 record count => Zero => Existing logic => IF Match APR001 name with OBL name = 50 % THEN "Code Found"   and take ARP001 Name(DBName) ELSE Code not found
                if (dtAliasForSingleRow != null && dtAliasForSingleRow.Rows.Count == 0) // ============= Need to discuss in else part
                {
                    // STEP :- 3_1 => Match Name => if match is greater than 50 % then consider valid record and return data table else set the variable as CODE NOT FOUND
                    if (!string.IsNullOrEmpty(CustomerSearchParmaters.CustomerNameOBL))
                    {
                        string NameFinalDB = dtOutputTemp.Rows[0]["CMNAME"].ToString().Trim().ToUpper();
                        NameFinalDB = Module1.func_RemoveSpecialWordsFromName(NameFinalDB);
                        string NameFinalOCR = CustomerSearchParmaters.CustomerNameOBL.ToString().Trim().ToUpper();
                        NameFinalOCR = Module1.func_RemoveSpecialWordsFromName(NameFinalOCR);
                        if (!string.IsNullOrEmpty(NameFinalDB) && !string.IsNullOrEmpty(NameFinalOCR))
                        {
                            clsCNCSBR ObjSbr = new clsCNCSBR();
                            RetStructIQSBR050 objRet50 = new RetStructIQSBR050();
                            objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(NameFinalDB.ToString()), Module1.func_RemoveSpecialCharacter(NameFinalOCR));
                            if (objRet50.PercentageMatch >= 50)
                            {
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = dtOutputTemp.Rows[0].ItemArray;
                                dtOutput.Rows.Add(dataRow);
                                clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "F"; // SET variable CODE FOUND
                            }
                            else
                            {
                                // if we found single record and name is not matched then check street 100, addressline 70, city 100 zip 100 and state 100
                                //if (isAllMatchExceptName == true)
                                //{
                                //    DataRow dataRow = dtOutput.NewRow();
                                //    dataRow.ItemArray = dtOutputTemp.Rows[0].ItemArray;
                                //    dtOutput.Rows.Add(dataRow);
                                //    //TO DO => SET variable CODE FOUND
                                //    clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "F";
                                //}
                                //else
                                //{
                                //TO DO => SET variable CODE NOT FOUND
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = dtOutputTemp.Rows[0].ItemArray;
                                dtOutput.Rows.Add(dataRow);
                                clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "AF"; // Only Address Match
                                //}
                            }
                        }
                        else
                        {
                            // SET variable CODE NOT FOUND
                            clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "N";
                        }
                    }
                    else
                    {
                        // SET variable CODE NOT FOUND
                        clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "N";
                    }
                }

                // STEP :- 3_0_2 =>  Check ARP033 record count => Multiple Count => Match OBL name with each row from DB and take hight match
                if (dtAliasForSingleRow != null && dtAliasForSingleRow.Rows.Count > 0)
                {
                    //string strCustomerCode = dtOutputTemp.Rows[0]["CMCUST"].ToString().Trim();
                    string strCustomerCode = "'" + dtOutputTemp.Rows[0]["CMCUST"].ToString().Trim() + "'";

                    //DataTable dtMatchedAliasFromARP033 = GetMatchedAliasFromARP033(strCustomerCode, CustomerSearchParmaters.CustomerNameOBL, SAIAOdbcConnectionObject);
                    DataTable dtMatchedAliasFromARP033 = GetMatchedAliasFromARP033(strCustomerCode, CustomerSearchParmaters.CustomerNameOBL);

                    if (dtMatchedAliasFromARP033 != null && dtMatchedAliasFromARP033.Rows.Count > 0) //  "Code found" and take Collection Name , & Code      
                    {
                        DataRow dataRow = dtOutput.NewRow();
                        dataRow.ItemArray = dtOutputTemp.Rows[0].ItemArray;
                        dtOutput.Rows.Add(dataRow);
                        clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "AD"; // SET variable CODE FOUND and take Name from ARP033 
                        if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                        {
                            clsF27.ClsSearchEngin.ClsSearchEnginAliseShipperNameARP033 = dtMatchedAliasFromARP033.Rows[0]["NANAME"].ToString();
                        }
                        else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                        {
                            clsF27.ClsSearchEngin.ClsSearchEnginAliseConsigneeNameARP033 = dtMatchedAliasFromARP033.Rows[0]["NANAME"].ToString();
                        }
                        else
                        {
                            clsF27.ClsSearchEngin.ClsSearchEnginAliseBillToNameARP033 = dtMatchedAliasFromARP033.Rows[0]["NANAME"].ToString();
                        }
                    }
                    else  //"Code Found" - Take "OBL Name"
                    {
                        DataRow dataRow = dtOutput.NewRow();
                        dataRow.ItemArray = dtOutputTemp.Rows[0].ItemArray;
                        dtOutput.Rows.Add(dataRow);
                        clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "AO"; // SET variable CODE FOUND and take OBL Name 
                    }
                    
                }

            }
            // STEP :- 3 => If we found multiple record then check ARP033 against each record and match percentage
            else
            {
                // Get all customer codes in string
                string strCustomerCode = "";
                for (int rowIndexTemp = 0; rowIndexTemp < dtOutputTemp.Rows.Count; rowIndexTemp++)
                {
                    if (rowIndexTemp == dtOutputTemp.Rows.Count-1)
                    {
                        strCustomerCode += "'" + dtOutputTemp.Rows[rowIndexTemp]["CMCUST"].ToString() + "'";
                    }
                    else
                    {
                        strCustomerCode += "'" + dtOutputTemp.Rows[rowIndexTemp]["CMCUST"].ToString() + "'" + ",";
                    }
                }
                //DataTable dtMatchedAliasFromARP033 = GetMatchedAliasFromARP033(strCustomerCode.ToString(), CustomerSearchParmaters.CustomerNameOBL, SAIAOdbcConnectionObject);
                DataTable dtMatchedAliasFromARP033 = GetMatchedAliasFromARP033(strCustomerCode.ToString(), CustomerSearchParmaters.CustomerNameOBL);

                if (dtMatchedAliasFromARP033 != null && dtMatchedAliasFromARP033.Rows.Count > 0) //  "Code found" and take Collection Name , & Code      
                {
                    // get the row index of dtOutputTemp against which code highest match is found
                    string QueyExpression = " CMCUST = '" + dtMatchedAliasFromARP033.Rows[0]["NACUST"].ToString() + "'";
                    if (!string.IsNullOrEmpty(QueyExpression))
                    {
                        DataRow[] foundRows;
                        foundRows = dtOutputTemp.Select(QueyExpression);
                        if (foundRows != null && foundRows.Length == 1) // if we found single row then return as valid customer code 
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = foundRows[0].ItemArray;
                            dtOutput.Rows.Add(dataRow);

                            clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "AD"; // TO DO => SET variable CODE FOUND and take Name from ARP033 
                            if (CustomerAccountType == SearchParameters.CustomerType.Shipper)
                            {
                                clsF27.ClsSearchEngin.ClsSearchEnginAliseShipperNameARP033 = dtMatchedAliasFromARP033.Rows[0]["NANAME"].ToString();
                            }
                            else if (CustomerAccountType == SearchParameters.CustomerType.Consignee)
                            {
                                clsF27.ClsSearchEngin.ClsSearchEnginAliseConsigneeNameARP033 = dtMatchedAliasFromARP033.Rows[0]["NANAME"].ToString();
                            }
                            else
                            {
                                clsF27.ClsSearchEngin.ClsSearchEnginAliseBillToNameARP033 = dtMatchedAliasFromARP033.Rows[0]["NANAME"].ToString();
                            }
                        }
                        else
                        {
                            clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "N";
                        }
                    }     
                   
                }
                else  //If no match in alias then "No Cod Found"
                {
                    clsF27.ClsSearchEngin.ClsSearchCommonStatusFlag = "N";
                }
            }
            return dtOutput;
        }

        public static DataTable FetchFinalMatchedBlindShipper(DataTable MatchResults, DataTable AbbreviationTable, AddressSection AddressSectionsFromOCR,
            SearchParameters.CustomerType CustomerAccountType)
        {
            DataTable dtOutput = new DataTable();
            dtOutput = MatchResults.Clone();

            for (int rowIndex = 0; rowIndex < MatchResults.Rows.Count; rowIndex++)
            {
                try
                {
                    DataRow currentDBRow = MatchResults.Rows[rowIndex];

                    AddressSection objDBAddress = new AddressSection();
                    objDBAddress.ZipCode = Convert.ToString(currentDBRow["CMZIP"]).Trim();
                    objDBAddress.State = Convert.ToString(currentDBRow["CMST"]).Trim();
                    objDBAddress.City = Convert.ToString(currentDBRow["CMCITY"]).Trim();

                    ProcessAddressParameters objProcessAddressParameters = new ProcessAddressParameters();
                    objProcessAddressParameters.AddressLine1 = Convert.ToString(currentDBRow["CMADR1"]);
                    objProcessAddressParameters.AddressLine2 = Convert.ToString(currentDBRow["CMADR2"]);
                    objProcessAddressParameters.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]);
                    objProcessAddressParameters.StreetDirection = Convert.ToString(currentDBRow["CMADRU"]);

                    objDBAddress.AddressLine1 = ProcessAddressBeforeComparison(objProcessAddressParameters, AbbreviationTable); // process address line of DB and then pass it as parameter

                    objDBAddress.Direction = Convert.ToString(currentDBRow["CMADRU"]).Trim();
                    objDBAddress.StreetNumber = Convert.ToString(currentDBRow["CMADRN"]).Trim();

                    // Added on 08th Feb 2021 ==> Check if the Street Number is 0 or blank, 
                    // add a check in CMADR1 AND/OR CMADR2 to find first numeric word and consider it as Street Number for comparison
                    if (string.IsNullOrEmpty(objDBAddress.StreetNumber) || objDBAddress.StreetNumber == "0")
                    {
                        bool isStreetNumberFound = false;
                        if (!string.IsNullOrEmpty(objProcessAddressParameters.AddressLine1))
                        {
                            bool isStreetNumberPresent = Regex.IsMatch(objProcessAddressParameters.AddressLine1.Split(' ')[0], @"[\d]+");
                            if (isStreetNumberPresent)
                            {
                                isStreetNumberFound = true;
                                objDBAddress.StreetNumber = objProcessAddressParameters.AddressLine1.Split(' ')[0];
                            }
                        }
                        if (isStreetNumberFound == false && !string.IsNullOrEmpty(objProcessAddressParameters.AddressLine2))
                        {
                            bool isStreetNumberPresent = Regex.IsMatch(objProcessAddressParameters.AddressLine2.Split(' ')[0], @"[\d]+");
                            if (isStreetNumberPresent)
                            {
                                isStreetNumberFound = true;
                                objDBAddress.StreetNumber = objProcessAddressParameters.AddressLine2.Split(' ')[0];
                            }
                        }
                    }

                    string matchResults = MatchAddressUsingRules(objDBAddress, AddressSectionsFromOCR, CustomerAccountType, AbbreviationTable);

                    if (!string.IsNullOrEmpty(matchResults))
                    {
                        if (matchResults.Split('~')[0] == "Y")
                        {
                            if (Convert.ToString(currentDBRow["CMNAME"]).Contains("(BS)"))
                            {
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = currentDBRow.ItemArray;
                                dtOutput.Rows.Add(dataRow);
                            }
                        }
                    }
                }
                catch
                {
                    continue;
                }
            }

            return dtOutput;
        }

        #region 12-Feb-2021
        public static DataTable GetMatchedAliasFromARP033(string StrCustomerCode, string CustomerNameOBL)
        {
            DataTable dtMatchedAliasFromARP033 = new DataTable();
            try
            {
                #region "Use Connection Object"
                // added on 17-Feb-2021 to use from f27
                string As400_ConnectionString = "";
                if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == false)
                {
                    OdbcConnection db2Connection = new OdbcConnection();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("ConnectionString_DB2Conn") == true)
                    {
                        As400_ConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                    }
                    else
                    {
                        As400_ConnectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                    }
                    db2Connection.ConnectionString = As400_ConnectionString; //iQPUBLIC.PublicComponents.DbConnectionString1; 

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
                        As400_ConnectionString = (string)iQPUBLIC.PublicComponents.htMyVariable["ConnectionString_DB2Conn"];
                    }
                    else
                    {
                        As400_ConnectionString = ConfigurationManager.ConnectionStrings["AS400DBConnString"].ToString();

                    }
                    connection.ConnectionString = As400_ConnectionString; //iQPUBLIC.PublicComponents.DbConnectionString1; 
                    connection.Open();
                    if (iQPUBLIC.PublicComponents.htMyVariable.ContainsKey("DB2Conn") == true)
                    {
                        iQPUBLIC.PublicComponents.htMyVariable.Remove("DB2Conn");
                    }
                    iQPUBLIC.PublicComponents.htMyVariable.Add("DB2Conn", connection);
                }
                #endregion
                string strQueryAliasFromARP033 = "select NACUST,NANAME FROM ARP033 WHERE NACUST in (" + StrCustomerCode + ") UNION SELECT CMCUST,CMNAME FROM ARP001 WHERE CMCUST in (" + StrCustomerCode + ")";
                DataTable dtAliasTemp = new DataTable();
                dtAliasTemp = ExecuteSearchQueryOnARP001(strQueryAliasFromARP033, connection);
                if (dtAliasTemp != null && dtAliasTemp.Rows.Count > 0)
                {
                    dtMatchedAliasFromARP033 = dtAliasTemp.Clone();
                    Dictionary<int, int> dicthightestMatchARP003 = new Dictionary<int, int>();
                    for (int rowIndexTemp = 0; rowIndexTemp < dtAliasTemp.Rows.Count; rowIndexTemp++)
                    {
                        try
                        {
                            DataRow currentDBRowTemp = dtAliasTemp.Rows[rowIndexTemp];
                            if (!string.IsNullOrEmpty(CustomerNameOBL))
                            {
                                string NameFinalDB = currentDBRowTemp["NANAME"].ToString().Trim().ToUpper();
                                NameFinalDB = Module1.func_RemoveSpecialWordsFromName(NameFinalDB);
                                string NameFinalOCR = CustomerNameOBL.ToString().Trim().ToUpper();
                                NameFinalOCR = Module1.func_RemoveSpecialWordsFromName(NameFinalOCR);
                                if (!string.IsNullOrEmpty(NameFinalDB) && !string.IsNullOrEmpty(NameFinalOCR))
                                {
                                    clsCNCSBR ObjSbr = new clsCNCSBR();
                                    RetStructIQSBR050 objRet50 = new RetStructIQSBR050();
                                    objRet50 = ObjSbr.IQSBR050(Module1.func_RemoveSpecialCharacter(NameFinalDB.ToString()), Module1.func_RemoveSpecialCharacter(NameFinalOCR));
                                    dicthightestMatchARP003.Add(rowIndexTemp, objRet50.PercentageMatch);
                                }
                            }
                        }
                        catch (Exception ex)
                        {
                            continue;
                        }

                    }
                    // Take highest match percentage 
                    if (dicthightestMatchARP003 != null && dicthightestMatchARP003.Count > 0)
                    {
                        var HighestPercentageOrderedAlias = dicthightestMatchARP003.OrderByDescending(item => item.Value);
                        int HighestPercentageAlias = Convert.ToInt32(HighestPercentageOrderedAlias.FirstOrDefault().Value);
                        int HighestPercentageRowIndexAlias = Convert.ToInt32(HighestPercentageOrderedAlias.FirstOrDefault().Key);
                        // if highest percentage is match then return that currect record
                        if (HighestPercentageAlias >= 90) // if it matched then reurn that record
                        {

                            DataRow dataRow = dtMatchedAliasFromARP033.NewRow();
                            dataRow.ItemArray = dtAliasTemp.Rows[HighestPercentageRowIndexAlias].ItemArray;
                            dtMatchedAliasFromARP033.Rows.Add(dataRow);

                        }
                    }
                }
                else
                {
                    dtMatchedAliasFromARP033 = null;
                }
            }
            catch (Exception ex)
            {

                dtMatchedAliasFromARP033 = null;
            }
            return dtMatchedAliasFromARP033;
        }
        #endregion
        #endregion

        #region UNUSED METHODS
        public static DataTable MatchRuleEngine2(DataTable SearchResults, SearchParameters.CustomerType CustomerAccountType)
        {
            DataTable dtOutput = SearchResults.Clone();
            try
            {
                if (SearchResults != null && SearchResults.Rows.Count > 0)
                {
                    // Find the highest value of address matching percentage from all search results to be considered for valid address matching
                    int addressHighestMatchPercentage = Convert.ToInt32(SearchResults.Compute("Max(AddrLine1Percent)", ""));

                    for (int index = 0; index < SearchResults.Rows.Count; index++)
                    {
                        DataRow currentRow = SearchResults.Rows[index];

                        // If any one of the following rules is satisfied, return the matching records

                        // Rule 1 - Street# - 100%, Address Line 1 - 100%, State - 100%, Zip - 100% OR City - 100%
                        if (Convert.ToInt32(currentRow["StreetNoPercent"]) == 100 && Convert.ToInt32(currentRow["AddrLine1Percent"]) == 100 &&
                            Convert.ToInt32(currentRow["StatePercent"]) == 100 && (Convert.ToInt32(currentRow["ZipPercent"]) == 100 || Convert.ToInt32(currentRow["CityPercent"]) == 100))
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 2 - Street# - 100%, Address Line 1 >= 70 (and highest from all records), State - 100%, Zip - 100% OR City - 100%
                        else if (Convert.ToInt32(currentRow["StreetNoPercent"]) == 100 &&
                            (Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                            Convert.ToInt32(currentRow["StatePercent"]) == 100 && (Convert.ToInt32(currentRow["ZipPercent"]) == 100 || Convert.ToInt32(currentRow["CityPercent"]) == 100))
                        {
                            DataRow dataRow = dtOutput.NewRow();
                            dataRow.ItemArray = currentRow.ItemArray;
                            dtOutput.Rows.Add(dataRow);
                        }
                        // Rule 3 - Street# - 1 digit changed, Address Line 1 - 100%, State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                        else if (Convert.ToInt32(currentRow["StreetNoChar"]) == 1 && Convert.ToInt32(currentRow["AddrLine1Percent"]) == 100 &&
                            Convert.ToInt32(currentRow["StatePercent"]) == 100 && (Convert.ToInt32(currentRow["ZipPercent"]) == 100 || Convert.ToInt32(currentRow["CityPercent"]) == 100))
                        {
                            if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                            {
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = currentRow.ItemArray;
                                dtOutput.Rows.Add(dataRow);
                            }
                        }
                        // Rule 4 - Street# - 1 digit changed, Address Line 1 >= 70 (and highest from all records), State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                        else if (Convert.ToInt32(currentRow["StreetNoChar"]) == 1 &&
                             (Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                            Convert.ToInt32(currentRow["StatePercent"]) == 100 && (Convert.ToInt32(currentRow["ZipPercent"]) == 100 || Convert.ToInt32(currentRow["CityPercent"]) == 100))
                        {
                            if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                            {
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = currentRow.ItemArray;
                                dtOutput.Rows.Add(dataRow);
                            }
                        }
                        // Rule 5 - Address Line 1 - 100%, State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                        else if (Convert.ToInt32(currentRow["AddrLine1Percent"]) == 100 &&
                            Convert.ToInt32(currentRow["StatePercent"]) == 100 && (Convert.ToInt32(currentRow["ZipPercent"]) == 100 || Convert.ToInt32(currentRow["CityPercent"]) == 100))
                        {
                            if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                            {
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = currentRow.ItemArray;
                                dtOutput.Rows.Add(dataRow);
                            }
                        }
                        // Rule 6 - Address Line 1 >= 70 (and highest from all records), State - 100%, Zip - 100% OR City - 100% => Not applicable for Consignee as Street # needs to be an exact match
                        else if ((Convert.ToInt32(currentRow["AddrLine1Percent"]) == addressHighestMatchPercentage && Convert.ToInt32(currentRow["AddrLine1Percent"]) >= 70) &&
                            Convert.ToInt32(currentRow["StatePercent"]) == 100 && (Convert.ToInt32(currentRow["ZipPercent"]) == 100 || Convert.ToInt32(currentRow["CityPercent"]) == 100))
                        {
                            if (CustomerAccountType == SearchParameters.CustomerType.Shipper || CustomerAccountType == SearchParameters.CustomerType.ThirdParty)
                            {
                                DataRow dataRow = dtOutput.NewRow();
                                dataRow.ItemArray = currentRow.ItemArray;
                                dtOutput.Rows.Add(dataRow);
                            }
                        }
                    }
                }
            }
            catch
            {
                dtOutput = null;
            }
            return dtOutput;
        }

        private static DataTable FilterBaseResultsUsingStreetNumberAndDirection(DataTable BaseResults, string StreetNumber, string DirectionFirstLetter)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMADRN] = " + StreetNumber + " AND CMADRU LIKE '" + DirectionFirstLetter.ToUpper() + "%'").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }

        private static DataTable FilterBaseResultsUsingZipCodeAndCustNameFirstLetter(DataTable BaseResults, string ZipCode, string CustomerNameFirstLetter)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMZIP] LIKE '" + ZipCode.ToUpper() + "%' AND [CMNAME] LIKE '" + CustomerNameFirstLetter.ToUpper() + "%'").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }

        private static DataTable FilterBaseResultsUsingCityAndCustNameFirstLetter(DataTable BaseResults, string City, string CustomerNameFirstLetter)
        {
            DataTable filteredResults = BaseResults.Clone();
            try
            {
                BaseResults.Select("[CMCITY] LIKE '" + City.ToUpper() + "%' AND [CMNAME] LIKE '" + CustomerNameFirstLetter.ToUpper() + "%'").CopyToDataTable(filteredResults, LoadOption.PreserveChanges);
            }
            catch
            {
                filteredResults = null;
            }
            return filteredResults;
        }
        #endregion
    }

    public class SearchParameters
    {
        public string CustomerName { get; set; }
        public string TerminalID { get; set; }
        public string StreetNumber { get; set; }
        public string DirectionFirstLetter { get; set; }
        public string ZipCode { get; set; }
        public string CustomerFirstLetter { get; set; }
        public string City { get; set; }
        // added on 9-Feb-2021
        public string CustomerNameOBL { get; set; }

        public enum CustomerType
        {
            Shipper,
            Consignee,
            ThirdParty
        }
        public enum SearchType
        {
            Normal,
            BlindShipper
        }
    }

    public class ProcessAddressParameters
    {
        public string AddressLine1 { get; set; }
        public string AddressLine2 { get; set; }
        public string StreetNumber { get; set; }
        public string StreetDirection { get; set; }
    }
}
